
# Initialize platform flags
$IsPlatformWindows = $false
$IsPlatformLinux = $false
$IsPlatformMacOS = $false

# Detect the operating system
if ([System.Environment]::OSVersion.VersionString -match "Windows") {
    $IsPlatformWindows = $true
} elseif ($PSVersionTable.OS -match "Linux") {
    $IsPlatformLinux = $true
} elseif ($PSVersionTable.OS -match "Darwin") {
    $IsPlatformMacOS  = $true
}

# Set environment paths
$0 = $MyInvocation.MyCommand.Definition
$env:dp0 = [System.IO.Path]::GetDirectoryName($0)
$env:AMSScriptPath = $env:dp0
$bits = Split-Path -Path $env:dp0 -Parent

# Display paths
Write-Host -ForegroundColor Green "Script environment paths initialized:"
Write-Host "Script Path: $env:dp0"
Write-Host "AMS Script Path: $env:AMSScriptPath"
Write-Host "Azure Excel Path: $env:AzureExcelPath"

# Function to ensure a module is installed
function EnsureModule {
    param ([string]$ModuleName)
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "Installing module: $ModuleName" -ForegroundColor Cyan
        Install-Module -Name $ModuleName -Scope CurrentUser -AllowClobber -Force
    }
}

#Function to extract user attributes from a row
function Extract-UserAttributes {
    param ($row)
    $userPrincipalName = $row.'Work Contact: Work Email'
    $displayName = if (($row.Name -split ',').Count -eq 2) {
        ($row.Name -split ',')[1].Trim() + " " + ($row.Name -split ',')[0].Trim()
    }
    else {
        $row.Name
    }
    return @{
        userPrincipalName = $userPrincipalName
        displayName       = $displayName
        jobTitle          = $row.'Job Title Description'
        department        = $row.'Home Department Description'
        phoneNumber       = $row.'Work Contact: Work Phone'
        physicalLocation  = $row.'Location Description'
        manager           = $row.'Reports To Email'
    }
}

#Function to get current user attributes from Azure
function Get-CurrentUserAttributes {
    param (
        [string]$userPrincipalName
    )

    Write-Host "Fetching attributes for user:" -ForegroundColor Cyan
    Write-Host $userPrincipalName -ForegroundColor Yellow

    # Initialize return variables
    $currentUser = $null
    $currentManager = $null
    $errorDetails = @()

    try {
        # Attempt to get the current user details
        $currentUser = Get-MgUser -UserId $userPrincipalName -Property DisplayName, Department, JobTitle, OfficeLocation, BusinessPhones -ErrorAction Stop
        Write-Host "User details retrieved successfully." -ForegroundColor Green
    } 
    catch {
        #Write-Host "Error retrieving user details for " $userPrincipalName $($_.Exception.Message) -ForegroundColor Red
        $errorDetails += [PSCustomObject]@{
            Operation    = "Get-MgUser"
            UserId       = $userPrincipalName
            ErrorMessage = $_.Exception.Message
        }
    }

    try {
        # Attempt to get the manager details
        if(-not ($currentUser.JobTitle.ToLower() -match "chief executive officer"))
        {
            $managerObject = Get-MgUserManager -UserId $userPrincipalName -ErrorAction Stop
            if ($managerObject) {
                
                $currentManager = $managerObject.AdditionalProperties["mail"]
                Write-Host "Manager details retrieved successfully." -ForegroundColor Green
            } 
            else {
                Write-Host "No manager found for user $userPrincipalName." -ForegroundColor Yellow
            }
        }
    } 
    catch {
        #Write-Host "Error retrieving manager details for " $userPrincipalName $($_.Exception.Message) -ForegroundColor Red
        $errorDetails += [PSCustomObject]@{
            Operation    = "Get-MgUserManager"
            UserId       = $userPrincipalName
            ErrorMessage = $_.Exception.Message
        }
    }

    # Return the result as a PSCustomObject
    return [PSCustomObject]@{
        CurrentUser    = $currentUser
        CurrentManager = $currentManager
        Errors         = if ($errorDetails.Count -gt 0) { $errorDetails } else { $null }
    }
}

#Function to update current user attributes to Azure
function Update-UserAttributes {
    param (
        [string]$userPrincipalName,
        [hashtable]$userUpdatedAttributes
    )

    Write-Host "Processing update for user:" -ForegroundColor Cyan
    Write-Host $userPrincipalName -ForegroundColor Yellow
    Write-Host "Attributes to update:" -ForegroundColor Cyan
    Write-Host $userUpdatedAttributes

    try {
        # Convert to a new hashtable with property names and their NewValue
        $newValues = @{}
        foreach ($key in $userUpdatedAttributes.Keys) {
            $newValues[$key] = $userUpdatedAttributes[$key].NewValue
        }

        # Validate and log the UserId
        if (-not [string]::IsNullOrWhiteSpace($userPrincipalName)) {
            Write-Host "Updating user with UserId: $userPrincipalName" -ForegroundColor Green
            Write-Host "Properties to update:" -ForegroundColor Cyan
            Write-Host $newValues   

            # Attempt to update the user
            Update-MgUser -UserId $userPrincipalName -BodyParameter $newValues -ErrorAction Stop
            if($newValues.ContainsKey("Manager") -and (-not [string]::IsNullOrWhiteSpace($newValues.Manager)))
            {
               Update-UserMgrAttribute -UserPrincipalName $userPrincipalName -ManagerPrincipalName $newValues.Manager
            }
            Write-Host "User update successful for $userPrincipalName." -ForegroundColor Green

        } 
        else {
            Write-Host "Invalid UserId: $userPrincipalName" -ForegroundColor Red
        }
    }
    catch {
        # Log the error
        #Write-Host "Error updating user " $userPrincipalName $($_.Exception.Message) -ForegroundColor Red
        return [PSCustomObject]@{
            UserPrincipalName = $userPrincipalName
            Status            = "Failed"
            Error             = $_.Exception.Message
        }

    }
}

#Funtion to update manager
function Update-UserMgrAttribute {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserPrincipalName, # UPN of the user whose manager you want to update
        
        [Parameter(Mandatory = $true)]
        [string]$ManagerPrincipalName # UPN of the new manager
    )
    try{
        # Define the user and manager details
        $userPrincipalName = $UserPrincipalName    
        # User whose manager you want to update
        $managerPrincipalName = $ManagerPrincipalName
         # Manager's UPN or Object ID

        # Get User and Manager Object IDs
        $userId = (Get-MgUser -Filter "userPrincipalName eq '$userPrincipalName'").Id
        if (-not $userId) {
            throw "User '$UserPrincipalName' not found."
        }
        $managerId = (Get-MgUser -Filter "userPrincipalName eq '$managerPrincipalName'").Id
        if (-not $managerId) {
            throw "Manager '$ManagerPrincipalName' not found."
        }

        # Construct the @odata.id reference for the manager
        $managerReference = @{
            "@odata.id" = "https://graph.microsoft.com/v1.0/users/$managerId"
        }

        # Update the user's manager
        Set-MgUserManagerByRef -UserId $userId -BodyParameter $managerReference

        Write-Host "Manager updated successfully for $userPrincipalName."
    }
    catch
    {
        # Log the error
        throw
    }

}
    
#Function to compare old and new values of the user
function Compare-AndTrackChanges {
    param ($currentUser, $currentManager, $userAttributes)
    $changes = @{}
    
    if ($currentUser.DisplayName -ne $userAttributes.displayName) {
        $changes['DisplayName'] = @{
            OldValue = $currentUser.DisplayName
            NewValue = $userAttributes.displayName
        }
    }
    if ($currentUser.JobTitle -ne $userAttributes.jobTitle) {
        $changes['JobTitle'] = @{
            OldValue = $currentUser.JobTitle
            NewValue = $userAttributes.jobTitle
        }
    }
    if ($currentUser.Department -ne $userAttributes.department) {
        $changes['Department'] = @{
            OldValue = $currentUser.Department
            NewValue = $userAttributes.department
        }
    }
    if (-not [string]::IsNullOrWhiteSpace($userAttributes.phoneNumber)) {
        $phoneNumber = $userAttributes.phoneNumber -replace '[^\d+]', ''
        if ($phoneNumber -match '^\+?[1-9]\d{6,14}$') {
            if ($currentUser.BusinessPhones -ne $phoneNumber) {
                $changes['BusinessPhones'] = @{                             #Changed Work Phone to Business Phone
                    OldValue = $currentUser.BusinessPhones
                    NewValue = $phoneNumber
                }
            }
        }
        else {
            Write-Warning "Invalid phone number for user ${userAttributes.userPrincipalName}: $phoneNumber"
        }
    }
    if ($currentUser.OfficeLocation -ne $userAttributes.physicalLocation) {
        $changes['OfficeLocation'] = @{
            OldValue = $currentUser.OfficeLocation
            NewValue = $userAttributes.physicalLocation
        }
    }
    if ($currentManager -ne $userAttributes.manager) {
        $changes['Manager'] = @{
            OldValue = $currentManager
            NewValue = $userAttributes.manager
        }
    }
    return $changes
}

#Function to process user data and generate changes
function Process-Users {
    param ($data, $previewOnly = $true)
    $updatedUsers = @()
    $failedUpdates = @()
    $previewChanges = @()
    

    foreach ($row in $data) {
        try {
            $userAttributes = Extract-UserAttributes -row $row

            #$currentUser, $currentManager = Get-CurrentUserAttributes -userPrincipalName $userAttributes.userPrincipalName
            $currentUserResults = Get-CurrentUserAttributes -userPrincipalName $userAttributes.userPrincipalName

            if ($currentUserResults.Errors -eq $null) {

                $changes = Compare-AndTrackChanges -currentUser $currentUserResults.CurrentUser -currentManager $currentUserResults.CurrentManager -userAttributes $userAttributes

                if ($changes.Count -gt 0) {
                    if (-not $previewOnly) {
                        # Attempt to update user attributes
                        $updateResult = Update-UserAttributes -userPrincipalName $userAttributes.userPrincipalName -userUpdatedAttributes $changes #Changed User Attribute to Changes

                        if ($updateResult.Error -ne $null) {
                            # Log the failed update if there is an error
                            $failedUpdates += [PSCustomObject]@{
                                UserPrincipalName = $userAttributes.userPrincipalName
                                Status            = "Failed"
                                Message           = $updateResult.Error
                            }
                            Write-Host "Failed to update user $($userAttributes.userPrincipalName): $($updateResult.Error)" -ForegroundColor Red
                        } 
                        else {
                            # Log successful update
                            $updatedUsers += [PSCustomObject]@{
                                UserPrincipalName = $userAttributes.userPrincipalName
                                Status            = "Updated"
                                Message           = $changes | ConvertTo-Json -Depth 1
                            }
                            Write-Host "User $($userAttributes.userPrincipalName) updated successfully." -ForegroundColor Green
                        }

                    } 

                    $previewChanges += ConvertHTToDict $changes

                }
               
                #$previewChanges += ConvertHTToDict $changes
            }
            else {
                $concatMessage = $null
                if ($currentUserResults.Errors) {
                    foreach ($error in $currentUserResults.Errors) {
                        $concatMessage += "Operation  $($error.Operation)" + " ErrorMessage $($error.ErrorMessage)" + "`n"
                    }
                }
                
                $failedUpdates += [PSCustomObject]@{
                    UserPrincipalName = $userAttributes.userPrincipalName
                    Status            = "Failed"
                    Message           = $concatMessage
                }

            }
            
        }
        catch {
            $failedUpdates += [PSCustomObject]@{
                UserPrincipalName = $row.'Work Contact: Work Email'
                Status            = "Failed"
                Message           = $_.Exception.Message
            }
        }
    }
    

    # Export results
    if (-not $previewOnly) {
        $timestamp = Get-Date -Format 'yyyyMMddHHmmss'
        $exportPath = Join-Path $env:AMSExportPath "UpdateResults_$timestamp.csv"

        if (($updatedUsers.Count -gt 0) -or ($failedUpdates.Count -gt 0))
        {
            $consolidatedResults = $updatedUsers + $failedUpdates
            $consolidatedResults | Export-Csv -Path $exportPath -NoTypeInformation
            Write-Host "`nThe status update has been successfully exported to $exportPath" -ForegroundColor Yellow
            Write-Host "Changes applied successfully." -ForegroundColor Yellow
        }
        else
        {
            Write-Host "`nThere were no changes to be made." -ForegroundColor Yellow
        }
    }
    else {
        return $previewChanges
    }

}

#Function to convert HashTable to Dictionary for better preview
Function ConvertHTToDict() {
    param(
        [hashtable] $changes
    )

    Write-Host "$($changes.Count) changes found." -ForegroundColor Green
    # Assume $changes is an array of hashtables or dictionaries
    $tempChanges = $changes.Clone() # Create a copy of the original $changes

    # Create a dictionary to maintain the order
    $tempChangesDict = New-Object 'System.Collections.Generic.Dictionary[Object, Object]'

    # First, add the "UserPrincipalName" key-value pair to the dictionary
    $tempChangesDict["UserPrincipalName"] = $userAttributes.userPrincipalName

    # Add the existing keys from $tempChanges to the dictionary (in order)
    foreach ($key in $tempChanges.Keys) {
        $tempChangesDict[$key] = $tempChanges[$key]
    }

    # Now add this dictionary back to the preview changes
    return $tempChangesDict
}

# Function to process each change and generate a preview string
function Get-PreviewString {
    param (
        [System.Collections.Generic.Dictionary[Object, Object]]$Change
    )

    $PreviewValues = $null

    # Loop through keys of the hashtable
    foreach ($key in $Change.Keys) {
        if ($key -eq "UserPrincipalName") {
            # Add the UserPrincipalName first
            $PreviewValues += $Change[$key] + " "
        }
        else {
            # Retrieve OldValue and NewValue for other keys
            $oldValue = $Change[$key].OldValue
            $newValue = $Change[$key].NewValue

            $PreviewValues += "$($key): OldValue='$oldValue', NewValue='$newValue'; "
        }
    }

    return $PreviewValues
}

# Display-Changes function to process all changes
function Display-Changes {
    param (
        [array]$Changes
    )

    $PreviewResult = ""

    # Process each change in the input array
    foreach ($change in $Changes) {
        $PreviewResult += Get-PreviewString -Change $change
        $PreviewResult += "`n" # Add a new line after each change
    }

    return $PreviewResult
}


# Function to open a file picker on macOS
function Show-OpenFileDialog {   
    $script = @"
        set filePath to POSIX path of (choose file)
        return filePath
"@
    $result = osascript -e $script
    return $result
}

#Function to open a file picker on macOS
function Show-OpenFolderDialog {
    $script = @"
        set folderPath to POSIX path of (choose folder)
        return folderPath
"@
    $result = osascript -e $script
    return $result.Trim()
}

# Function to select a file (cross-platform)
function SelectFile {
    # Cross-platform file selection logic
    $filepath = $null

    if ($IsPlatformWindows) {
        Add-Type -AssemblyName System.Windows.Forms
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.InitialDirectory = [Environment]::GetFolderPath("MyDocuments")
        $OpenFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"
        if ($OpenFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $filepath = $OpenFileDialog.FileName
        }
    } elseif ($IsPlatformLinux -or $IsPlatformMacOS) {
        $filepath = Show-OpenFileDialog
        Write-Host "Selected file Mac: $filepath"
        
    } else {
        throw "Unsupported operating system. This script works on Windows, Linux, and macOS only."
    }
    return $filepath
}

#Select Export folder location
function SelectExportLocation
{
    $ExportFolder = $null   
    if ($IsPlatformLinux -or $IsPlatformMacOS)
    {
        $ExportFolder = Show-OpenFolderDialog
    } 
    elseif ($IsPlatformWindows)
    {
        Add-Type -AssemblyName System.Windows.Forms
        $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
        $FolderBrowser.Description = "Select a folder to export the file"
        if ($FolderBrowser.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $ExportFolder = $FolderBrowser.SelectedPath
        } else {
            Write-Host "No folder selected. Exiting."
            exit
        }
    }
    return $ExportFolder.Trim()
}

# Main menu
function Main-Menu {
    EnsureModule -ModuleName "ImportExcel"
    EnsureModule -ModuleName "Microsoft.Graph"

    #Select the excel file to upload
    #Read-Host "Please select the Excel file that you would like to update on Azure? Press enter to select" 

    Write-Host "`nPlease select the Excel file that you would like to update on Azure!" -ForegroundColor Yellow
    Read-Host "Press Enter to open the file selection dialog" 
    $env:AzureExcelPath = SelectFile
    if (-not $env:AzureExcelPath) {
        Write-Host "No Excel files found in $env:AzureExcelPath" -ForegroundColor Red
        return
    }
    Write-host "Selected file" $env:AzureExcelPath  -ForegroundColor Green

    #Select the export file location
    Write-Host "`nWhere would you like to save the CSV file that contains the update status?" -ForegroundColor Yellow
    Read-Host "Press Enter to open the file selection dialog"
    $env:AMSExportPath = SelectExportLocation
    if (-not $env:AMSExportPath) {
        Write-Host "No folder selected. Exiting." -ForegroundColor Red
        exit
    }
    Write-host "Selected export location " $env:AMSExportPath  -ForegroundColor Green

   
    Connect-MgGraph -Scopes "User.Read.All", "User.ReadWrite.All", "Directory.ReadWrite.All"
    Write-Warning "Ensure that you have activated your privileged role before proceeding."

    $data = Import-Excel -Path $env:AzureExcelPath
    $exitMenu = $false

    while (-not $exitMenu) {
        Write-Host "Menu:" -ForegroundColor Green
        Write-Host "1. Preview Changes"
        Write-Host "2. Apply Changes"
        Write-Host "3. Quit"
        $choice = Read-Host "Enter your choice"

        switch ($choice) {
            1 {
                $finalResults = Process-Users -data $data -previewOnly $true
                $PreviewValues = Display-Changes -changes $finalResults
                # Output the result
                Write-Host $PreviewValues
            }
            2 {
                $finalResults = Process-Users -data $data -previewOnly $false
                # Write-Host "Changes applied successfully." -ForegroundColor Green
            }
            3 {
                $exitMenu = $true
            }
            default {
                Write-Host "Invalid choice. Please try again." -ForegroundColor Red
            }
        }
    }
    
    Disconnect-MgGraph
}

# Run Main Menu
Main-Menu







