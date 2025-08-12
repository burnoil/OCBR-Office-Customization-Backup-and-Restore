<#
.SYNOPSIS
    A comprehensive tool for administrators to back up and restore a user's Microsoft Office settings.
    Includes full support for PowerPoint, Visio, and file-based VBA Add-ins.

.DESCRIPTION
    This script is designed to be run as SYSTEM or an Administrator. It automatically detects the
    currently active desktop user and targets their profile for backup and restore of key Office settings.

.VERSION
    1.2.0

.PARAMETER UserName
    Optional. Explicitly specifies the username to target (e.g., 'jdoe'), overriding auto-detection.

.PARAMETER Action
    For command-line use. Must be either 'backup' or 'restore'.

.PARAMETER Path
    For command-line use. The root folder for the backup/restore operation.

.PARAMETER Items
    For command-line use. An array of items to process: 'RibbonUI', 'Templates', 'Signatures', 'Dictionaries', 'AutoComplete', 'ExcelMacros', 'AutoCorrect', 'VisioContent', 'AddIns'.

.EXAMPLE
    # Back up ALL supported settings, including file-based Add-ins, for the active user.
    .\OfficeCustomizationTool.ps1 -Action backup -Path "C:\Backups" -Items "RibbonUI","Templates","Signatures","Dictionaries","AutoComplete","ExcelMacros","AutoCorrect","VisioContent","AddIns"
#>
param(
    [Parameter(HelpMessage = "Explicitly specify a user to target, overriding auto-detection.")] [string]$UserName,
    [Parameter(HelpMessage = "Specify the action: backup or restore.")] [ValidateSet('backup', 'restore', IgnoreCase = $true)] [string]$Action,
    [Parameter(HelpMessage = "The root directory for the backup/restore files.")] [string]$Path,
    [Parameter(HelpMessage = "The specific items to process.")]
    [ValidateSet('RibbonUI', 'Templates', 'Signatures', 'Dictionaries', 'AutoComplete', 'ExcelMacros', 'AutoCorrect', 'VisioContent', 'AddIns', IgnoreCase = $true)]
    [string[]]$Items
)
# --- APPLICATION VERSION ---
$version = "1.2.0"

# --- Add necessary assemblies for GUI ---
Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName System.Drawing
# --- Global Variables & Logging Setup ---
$logDirectory = "C:\Windows\MITLL\Logs"; $logFile = Join-Path -Path $logDirectory -ChildPath "OfficeCustomizationBackup.log"; if (-not (Test-Path -Path $logDirectory)) { try { New-Item -Path $logDirectory -ItemType Directory -Force -ErrorAction Stop | Out-Null } catch { Write-Warning "Could not create log directory. Logging will be disabled."; $logFile = $null } }
$officeApps = "WINWORD", "EXCEL", "POWERPNT", "OUTLOOK", "VISIO", "MSACCESS", "MSPUB", "WINPROJ"; $script:customizationPaths = $null; $script:activeUser = $null
# --- Core Functions ---
function Get-ActiveUser { try { $explorerProcess = Get-CimInstance -ClassName Win32_Process -Filter "Name = 'explorer.exe'" -ErrorAction Stop; if (!$explorerProcess) { return $null }; $ownerInfo = $explorerProcess | ForEach-Object { $owner = Invoke-CimMethod -InputObject $_ -MethodName GetOwner; if ($owner.ReturnValue -eq 0) { [PSCustomObject]@{ Domain = $owner.Domain; User = $owner.User; SessionId = $_.SessionId } } } | Sort-Object -Property SessionId | Select-Object -First 1; if (!$ownerInfo) { return $null }; $userObject = Get-CimInstance -ClassName Win32_UserAccount -Filter "Domain = '$($ownerInfo.Domain)' AND Name = '$($ownerInfo.User)'"; $profile = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($userObject.SID)"; return [PSCustomObject]@{ UserName = "$($ownerInfo.Domain)\$($ownerInfo.User)"; ProfilePath = $profile.ProfileImagePath; SID = $userObject.SID } } catch { Write-Warning "Could not determine active user: $_"; return $null } }
function Get-OfficeCustomizationPaths {
    param([string]$UserProfilePath, [string]$UserSID)
    $roamingAppData = Join-Path -Path $UserProfilePath -ChildPath "AppData\Roaming"
    $localAppData = Join-Path -Path $UserProfilePath -ChildPath "AppData\Local"
    $userShellFolders = "Registry::HKEY_USERS\$UserSID\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    $documentsPath = (Get-ItemProperty -Path $userShellFolders -ErrorAction SilentlyContinue).Personal
    $paths = @{
        RibbonUI     = @{ Name = "Ribbon/Toolbar"; Path = Join-Path $localAppData "Microsoft\Office"; Exists = $false; Filter = "*.officeUI"; Type = "File" }
        Templates    = @{ Name = "Office Templates"; Path = Join-Path $roamingAppData "Microsoft\Templates"; Exists = $false; Filter = "*"; Type = "Folder" }
        Signatures   = @{ Name = "Outlook Signatures"; Path = Join-Path $roamingAppData "Microsoft\Signatures"; Exists = $false; Filter = "*"; Type = "Folder" }
        Dictionaries = @{ Name = "Custom Dictionaries"; Path = Join-Path $roamingAppData "Microsoft\UProof"; Exists = $false; Filter = "*.dic"; Type = "File" }
        AutoComplete = @{ Name = "Outlook Auto-Complete"; Path = Join-Path $localAppData "Microsoft\Outlook\RoamCache"; Exists = $false; Filter = "Stream_Autocomplete_*.dat"; Type = "File" }
        ExcelMacros  = @{ Name = "Global Excel Macros"; Path = Join-Path $roamingAppData "Microsoft\Excel\XLSTART"; Exists = $false; Filter = "PERSONAL.XLSB"; Type = "File" }
        AutoCorrect  = @{ Name = "Office AutoCorrect"; Path = Join-Path $roamingAppData "Microsoft\Office"; Exists = $false; Filter = "*.acl"; Type = "File" }
        AddIns       = @{ Name = "VBA/File Add-ins"; Path = Join-Path $roamingAppData "Microsoft\AddIns"; Exists = $false; Filter = "*"; Type = "Folder" } # NEW ITEM
    }
    if ($documentsPath -and (Test-Path $documentsPath)) {
        $paths.Add("VisioContent", @{ Name = "Visio Stencils/Templates"; Path = Join-Path $documentsPath "My Shapes"; Exists = $false; Filter = "*"; Type = "Folder" })
    }
    foreach ($key in $paths.Keys) {
        if ($paths[$key].Path -and (Test-Path $paths[$key].Path)) {
            if (Get-ChildItem -Path $paths[$key].Path -Filter $paths[$key].Filter -ErrorAction SilentlyContinue | Select-Object -First 1) { $paths[$key].Exists = $true }
        }
    }
    return $paths
}
# (All other functions remain the same)
function Write-Log { param([string]$message); if ($logFile) { "$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss')) $message" | Add-Content -Path $logFile } }
function Update-UIAndLog { param([string]$message); Write-Log $message; if ($PSBoundParameters.ContainsKey('Action')) { Write-Host $message }; if ($script:logTextBox) { $script:logTextBox.AppendText("`r`n$message"); $script:logTextBox.SelectionStart = $script:logTextBox.Text.Length; $script:logTextBox.ScrollToCaret(); $script:mainForm.Update() } }
function Start-Backup { param([string]$BackupPath, [string[]]$ItemsToBackup); Update-UIAndLog "Starting backup for user '$($script:activeUser.UserName)'..."; try { foreach ($key in $ItemsToBackup) { if ($script:customizationPaths.ContainsKey($key) -and $script:customizationPaths[$key].Exists) { $sourceInfo = $script:customizationPaths[$key]; $destination = Join-Path $BackupPath $key; if (!(Test-Path $destination)) { New-Item -Path $destination -ItemType Directory -Force | Out-Null }; Update-UIAndLog "Backing up $($sourceInfo.Name)..."; if ($sourceInfo.Type -eq "File") { Get-ChildItem -Path $sourceInfo.Path -Filter $sourceInfo.Filter | ForEach-Object { Copy-Item -Path $_.FullName -Destination $destination -Force } } elseif ($sourceInfo.Type -eq "Folder") { Copy-Item -Path ($sourceInfo.Path + "\*") -Destination $destination -Recurse -Force } } else { Update-UIAndLog "Skipping ${key}: Not found." } }; Update-UIAndLog "Backup completed successfully!"; return $true } catch { Update-UIAndLog "ERROR during backup: $_"; return $false } }
function Start-Restore { param([string]$RestorePath, [string[]]$ItemsToRestore); $runningOfficeProcs = Get-Process -Name $officeApps -ErrorAction SilentlyContinue; if ($runningOfficeProcs) { $message = "Office apps must be closed. Close them now?"; $result = 'No'; if (!$PSBoundParameters.ContainsKey('Action')) { $result = [System.Windows.Forms.MessageBox]::Show($message, "Warning", "YesNo", "Warning") } else { Update-UIAndLog "WARNING: Office apps running. Aborting."; return $false }; if ($result -eq 'Yes') { Update-UIAndLog "Closing Office apps..."; $runningOfficeProcs | Stop-Process -Force; Start-Sleep -Seconds 2 } else { Update-UIAndLog "Restore cancelled."; return $false } }; Update-UIAndLog "Starting restore for user '$($script:activeUser.UserName)'..."; try { foreach ($key in $ItemsToRestore) { if ($script:customizationPaths.ContainsKey($key)) { $sourceInfo = $script:customizationPaths[$key]; $destinationInfo = $script:customizationPaths[$key]; $source = Join-Path $RestorePath $key; if (Test-Path $source) { Update-UIAndLog "Restoring $($sourceInfo.Name)..."; Copy-Item -Path ($source + "\*") -Destination $destinationInfo.Path -Recurse -Force } else { Update-UIAndLog "WARNING: Source for $key not found. Skipping." } } }; Update-UIAndLog "Restore completed successfully!"; return $true } catch { Update-UIAndLog "ERROR during restore: $_"; return $false } }
function Update-DetectedPathsUI { Update-UIAndLog "Refreshing detected paths..."; $script:customizationPaths = Get-OfficeCustomizationPaths -UserProfilePath $script:activeUser.ProfilePath -UserSID $script:activeUser.SID; $pathMessage = ""; foreach($key in $script:checkboxMap.Keys){ if ($script:customizationPaths.ContainsKey($key)) { $pathInfo = $script:customizationPaths[$key]; $status = if ($pathInfo.Exists) { "Detected" } else { "Not Found" }; $checkboxMap[$key].Enabled = $pathInfo.Exists; $checkboxMap[$key].Checked = $pathInfo.Exists; $pathMessage += "$($pathInfo.Name)`: ($status)`r`n  $($pathInfo.Path)`r`n`r`n" } else { $checkboxMap[$key].Enabled = $false; $checkboxMap[$key].Checked = $false } }; $pathDisplayTextBox.Text = $pathMessage.Trim(); Update-UIAndLog "Detection refresh complete." }

# --- SCRIPT EXECUTION LOGIC ---
Write-Log "--------------------------------"; Write-Log "Application v$version starting..."
if ($UserName) { try { $userObject = Get-CimInstance -ClassName Win32_UserAccount -Filter "Name = '$UserName'" -ErrorAction Stop; if ($userObject) { $profile = Get-ItemProperty -Path "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($userObject.SID)"; $script:activeUser = [PSCustomObject]@{ UserName = $userObject.Name; ProfilePath = $profile.ProfileImagePath; SID = $userObject.SID } } } catch { Write-Warning "Could not find specified user '$UserName'. Error: $_" } } else { $script:activeUser = Get-ActiveUser }
if (!$script:activeUser) { $errorMessage = "FATAL: Could not determine active user profile. Cannot continue."; Update-UIAndLog $errorMessage; if (!$PSBoundParameters.ContainsKey('Action')) { [System.Windows.Forms.MessageBox]::Show($errorMessage, "Error", "OK", "Error") }; Exit 1 }

if ($PSBoundParameters.ContainsKey('Action')) {
    $script:customizationPaths = Get-OfficeCustomizationPaths -UserProfilePath $script:activeUser.ProfilePath -UserSID $script:activeUser.SID
    Update-UIAndLog "Running in command-line mode for user '$($script:activeUser.UserName)'."; if (-not ($PSBoundParameters.ContainsKey('Path') -and $PSBoundParameters.ContainsKey('Items'))) { Update-UIAndLog "ERROR: -Action, -Path, and -Items are mandatory."; Exit 1 }; $success = $false; if ($Action -eq 'backup') { $success = Start-Backup -BackupPath $Path -ItemsToBackup $Items } elseif ($Action -eq 'restore') { $success = Start-Restore -RestorePath $Path -ItemsToRestore $Items }; Update-UIAndLog "Command-line operation finished."; if ($success) { Exit 0 } else { Exit 1 }
}

# --- GUI Window and Controls ---
$script:mainForm = New-Object System.Windows.Forms.Form; $script:mainForm.Text = "Office Customization Backup & Restore - v$version"; $script:mainForm.MinimumSize = '580, 720'; $script:mainForm.Size = '600, 780'; $script:mainForm.StartPosition = "CenterScreen"; $script:mainForm.FormBorderStyle = "Sizable"
$optionsGroupBox = New-Object System.Windows.Forms.GroupBox; $optionsGroupBox.Location = '20, 20'; $optionsGroupBox.Size = '540, 140'; $optionsGroupBox.Text = "1. Select Items to Process"; $optionsGroupBox.Anchor = "Top, Left, Right"
$pathsGroupBox = New-Object System.Windows.Forms.GroupBox; $pathsGroupBox.Location = '20, 170'; $pathsGroupBox.Size = '540, 200'; $pathsGroupBox.Text = "2. Detected Paths for: $($script:activeUser.UserName)"; $pathsGroupBox.Anchor = "Top, Left, Right"
$actionGroupBox = New-Object System.Windows.Forms.GroupBox; $actionGroupBox.Location = '20, 380'; $actionGroupBox.Size = '540, 150'; $actionGroupBox.Text = "3. Perform Action"; $actionGroupBox.Anchor = "Top, Left, Right"
$logGroupBox = New-Object System.Windows.Forms.GroupBox; $logGroupBox.Location = '20, 540'; $logGroupBox.Size = '540, 170'; $logGroupBox.Text = "Activity Log"; $logGroupBox.Anchor = "Top, Bottom, Left, Right"
# Checkboxes
$chkRibbon = New-Object System.Windows.Forms.CheckBox; $chkRibbon.Text = "Ribbon/Toolbar"; $chkRibbon.Location = '20, 30'; $chkRibbon.AutoSize = $true
$chkTemplates = New-Object System.Windows.Forms.CheckBox; $chkTemplates.Text = "Office Templates"; $chkTemplates.Location = '180, 30'; $chkTemplates.AutoSize = $true
$chkSignatures = New-Object System.Windows.Forms.CheckBox; $chkSignatures.Text = "Outlook Signatures"; $chkSignatures.Location = '360, 30'; $chkSignatures.AutoSize = $true
$chkDictionaries = New-Object System.Windows.Forms.CheckBox; $chkDictionaries.Text = "Custom Dictionaries"; $chkDictionaries.Location = '20, 55'; $chkDictionaries.AutoSize = $true
$chkAutoComplete = New-Object System.Windows.Forms.CheckBox; $chkAutoComplete.Text = "Outlook Auto-Complete"; $chkAutoComplete.Location = '200, 55'; $chkAutoComplete.AutoSize = $true
$chkExcelMacros = New-Object System.Windows.Forms.CheckBox; $chkExcelMacros.Text = "Global Excel Macros"; $chkExcelMacros.Location = '20, 80'; $chkExcelMacros.AutoSize = $true
$chkAutoCorrect = New-Object System.Windows.Forms.CheckBox; $chkAutoCorrect.Text = "Office AutoCorrect"; $chkAutoCorrect.Location = '200, 80'; $chkAutoCorrect.AutoSize = $true
$chkVisioContent = New-Object System.Windows.Forms.CheckBox; $chkVisioContent.Text = "Visio Stencils/Templates"; $chkVisioContent.Location = '20, 105'; $chkVisioContent.AutoSize = $true
$chkAddIns = New-Object System.Windows.Forms.CheckBox; $chkAddIns.Text = "VBA/File Add-ins"; $chkAddIns.Location = '200, 105'; $chkAddIns.AutoSize = $true
$helpButton = New-Object System.Windows.Forms.Button; $helpButton.Location = '360, 103'; $helpButton.Size = '160, 23'; $helpButton.Text = "What do these items mean?"; $helpButton.Anchor = "Top, Right"
$optionsGroupBox.Controls.AddRange(@($chkRibbon, $chkTemplates, $chkSignatures, $chkDictionaries, $chkAutoComplete, $chkExcelMacros, $chkAutoCorrect, $chkVisioContent, $chkAddIns, $helpButton))
# Paths display and other controls...
$pathDisplayTextBox = New-Object System.Windows.Forms.TextBox; $pathDisplayTextBox.Location = '15, 25'; $pathDisplayTextBox.Size = '510, 130'; $pathDisplayTextBox.Multiline = $true; $pathDisplayTextBox.ReadOnly = $true; $pathDisplayTextBox.Scrollbars = "Vertical"; $pathDisplayTextBox.Anchor = "Top, Bottom, Left, Right"; $pathDisplayTextBox.Font = "Consolas, 8.5"
$refreshButton = New-Object System.Windows.Forms.Button; $refreshButton.Location = '15, 160'; $refreshButton.Size = '510, 25'; $refreshButton.Text = "Refresh Detection"; $refreshButton.Anchor = "Bottom, Left, Right"
$pathsGroupBox.Controls.AddRange(@($pathDisplayTextBox, $refreshButton))
$pathLabel = New-Object System.Windows.Forms.Label; $pathLabel.Text = "Backup Location / Restore Source Path:"; $pathLabel.Location = '20, 30'; $pathLabel.AutoSize = $true
$pathTextBox = New-Object System.Windows.Forms.TextBox; $pathTextBox.Location = '20, 50'; $pathTextBox.Size = '420, 20'; $pathTextBox.Anchor = "Top, Left, Right"
$browseButton = New-Object System.Windows.Forms.Button; $browseButton.Location = '450, 48'; $browseButton.Size = '75, 23'; $browseButton.Text = "Browse..."; $browseButton.Anchor = "Top, Right"
$backupButton = New-Object System.Windows.Forms.Button; $backupButton.Location = '20, 90'; $backupButton.Size = '245, 40'; $backupButton.Text = "BACKUP"; $backupButton.Anchor = "Top, Left, Right"
$restoreButton = New-Object System.Windows.Forms.Button; $restoreButton.Location = '275, 90'; $restoreButton.Size = '250, 40'; $restoreButton.Text = "RESTORE"; $restoreButton.Anchor = "Top, Right"
$actionGroupBox.Controls.AddRange(@($pathLabel, $pathTextBox, $browseButton, $backupButton, $restoreButton))
$script:logTextBox = New-Object System.Windows.Forms.TextBox; $script:logTextBox.Location = '15, 25'; $script:logTextBox.Size = '510, 100'; $script:logTextBox.Multiline = $true; $script:logTextBox.Scrollbars = "Vertical"; $script:logTextBox.ReadOnly = $true; $script:logTextBox.Anchor = "Top, Bottom, Left, Right"
$logFileLabel = New-Object System.Windows.Forms.Label; $logFileLabel.Location = '15, 135'; $logFileLabel.Size = '510, 20'; $logFileLabel.Anchor = "Bottom, Left, Right";
$logGroupBox.Controls.AddRange(@($script:logTextBox, $logFileLabel))
$script:mainForm.Controls.AddRange(@($optionsGroupBox, $pathsGroupBox, $actionGroupBox, $logGroupBox))
$script:checkboxMap = @{ RibbonUI = $chkRibbon; Templates = $chkTemplates; Signatures = $chkSignatures; Dictionaries = $chkDictionaries; AutoComplete = $chkAutoComplete; ExcelMacros = $chkExcelMacros; AutoCorrect = $chkAutoCorrect; VisioContent = $chkVisioContent; AddIns = $chkAddIns }

# --- GUI Event Handlers ---
$script:mainForm.Add_Load({ Update-UIAndLog "GUI v$version started for user '$($script:activeUser.UserName)'."; if ($logFile) { $logFileLabel.Text = "Log File: $logFile" } else { $logFileLabel.Text = "Log File: Disabled (insufficient permissions)" }; Update-DetectedPathsUI })
$refreshButton.Add_Click({ Update-DetectedPathsUI })
$helpButton.Add_Click({
    $bullet = [char]0x2022
    $helpMessage = @"
ITEM DESCRIPTIONS:

$bullet Ribbon/Toolbar: Custom button layout for the top ribbon and Quick Access Toolbar in all Office apps.

$bullet Office Templates: Custom Word (.dotx), PowerPoint (.potx), and other templates you have created.

$bullet Outlook Signatures: All of your saved email signatures for Outlook.

$bullet Custom Dictionaries: Custom words added to the spelling dictionary.

$bullet Outlook Auto-Complete: The list of cached email addresses that appear when you start typing in the 'To:' field.

$bullet Global Excel Macros: Macros stored in your PERSONAL.XLSB file, available to all Excel workbooks.

$bullet Office AutoCorrect: Your list of custom text replacements (e.g., '(c)' to a copyright symbol).

$bullet Visio Stencils/Templates: Custom shape collections (.vssx) and templates (.vstx) from your 'My Shapes' folder.

$bullet VBA/File Add-ins: User-installed add-ins for Excel (.xlam) and PowerPoint (.ppam). This does not include installed programs (COM Add-ins).
"@
    [System.Windows.Forms.MessageBox]::Show($helpMessage, "Backup Item Help", "OK", "Information")
})
$browseButton.Add_Click({ $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog; if ($folderBrowser.ShowDialog() -eq "OK") { $pathTextBox.Text = $folderBrowser.SelectedPath } })
$backupButton.Add_Click({ if ([string]::IsNullOrWhiteSpace($pathTextBox.Text)) { [System.Windows.Forms.MessageBox]::Show("Select backup path.", "Error", "OK", "Error"); return }; $items = ($script:checkboxMap.GetEnumerator() | Where-Object { $_.Value.Checked } | ForEach-Object { $_.Key }); if ($items.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Select at least one item.", "Warning", "OK", "Warning"); return }; if (Start-Backup -BackupPath $pathTextBox.Text -ItemsToBackup $items) { [System.Windows.Forms.MessageBox]::Show("Backup complete.", "Success", "OK", "Information") } else { [System.Windows.Forms.MessageBox]::Show("Error during backup. Check log.", "Error", "OK", "Error") } })
$restoreButton.Add_Click({ if ([string]::IsNullOrWhiteSpace($pathTextBox.Text)) { [System.Windows.Forms.MessageBox]::Show("Select restore path.", "Error", "OK", "Error"); return }; $items = ($script:checkboxMap.GetEnumerator() | Where-Object { $_.Value.Checked } | ForEach-Object { $_.Key }); if ($items.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("Select at least one item.", "Warning", "OK", "Warning"); return }; if (Start-Restore -RestorePath $pathTextBox.Text -ItemsToRestore $items) { [System.Windows.Forms.MessageBox]::Show("Restore complete.", "Success", "OK", "Information") } else { [System.Windows.Forms.MessageBox]::Show("Error during restore. Check log.", "Error", "OK", "Error") } })

# --- Show the GUI ---
$script:mainForm.ShowDialog()
Write-Log "Application v$version closed."
# SIG # Begin signature block
# MIIMjgYJKoZIhvcNAQcCoIIMfzCCDHsCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUhe5zKd4ZrN/y4hGDSyJT0lYQ
# SmKgggnvMIIEwDCCA6igAwIBAgIBEzANBgkqhkiG9w0BAQsFADBWMQswCQYDVQQG
# EwJVUzEfMB0GA1UEChMWTUlUIExpbmNvbG4gTGFib3JhdG9yeTEMMAoGA1UECxMD
# UEtJMRgwFgYDVQQDEw9NSVRMTCBSb290IENBLTIwHhcNMTkwNzA4MTExMDAwWhcN
# MjkwNzA4MTExMDAwWjBRMQswCQYDVQQGEwJVUzEfMB0GA1UECgwWTUlUIExpbmNv
# bG4gTGFib3JhdG9yeTEMMAoGA1UECwwDUEtJMRMwEQYDVQQDDApNSVRMTCBDQS02
# MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAj2T0hoZXOA+UPr8SD/Re
# gKGDHDfz+8i1bm+cGV9V2Zxs1XxYrCBbnTB79AtuYR29HIf6HfsUrsqJH6gQtptF
# tux8QrWqx25iOE4tg2yeSVmrc/ZB4fRfufKi0idq2IA13kJgYQ8xCLpIiBEm8be7
# Lzlz9mGT0UVgRe3I5Jku935a7pOB2qHHH6OGWSs9AOPiJdo4oSWUbL5H3H5MmZCI
# 8T3Rj7dobmrRYOsUADI5kkqvOf7o1j09X7X2q4Q+ez4JHgGTLTxjvox7QEDYglZM
# Mh9qB2SGpvhCkKoZ3/05bT1oCt2Pb4iR7MlETNryi/mzZuOjf2gaYpuWweYVh2Ny
# 3wIDAQABo4IBnDCCAZgwEgYDVR0TAQH/BAgwBgEB/wIBADAdBgNVHQ4EFgQUk5BH
# A0LBTbQzHtRCl5+h4Ctwv4gwHwYDVR0jBBgwFoAU/8nJZUxTgPGpDDwhroIqx+74
# MvswDgYDVR0PAQH/BAQDAgGGMGcGCCsGAQUFBwEBBFswWTAuBggrBgEFBQcwAoYi
# aHR0cDovL2NybC5sbC5taXQuZWR1L2dldHRvL0xMUkNBMjAnBggrBgEFBQcwAYYb
# aHR0cDovL29jc3AubGwubWl0LmVkdS9vY3NwMDQGA1UdHwQtMCswKaAnoCWGI2h0
# dHA6Ly9jcmwubGwubWl0LmVkdS9nZXRjcmwvTExSQ0EyMIGSBgNVHSAEgYowgYcw
# DQYLKoZIhvcSAgEDAQYwDQYLKoZIhvcSAgEDAQgwDQYLKoZIhvcSAgEDAQcwDQYL
# KoZIhvcSAgEDAQkwDQYLKoZIhvcSAgEDAQowDQYLKoZIhvcSAgEDAQswDQYLKoZI
# hvcSAgEDAQ4wDQYLKoZIhvcSAgEDAQ8wDQYLKoZIhvcSAgEDARAwDQYJKoZIhvcN
# AQELBQADggEBALnwy+yzh/2SvpwC8q8EKdDQW8LxWnDM56DcHm5zgfi0WfEsQi8w
# xcV2Vb2eCNs6j0NofdgsSP7k9DJ6LmDs+dfZEmD23+r9zlMhI6QQcwlvq+cgTrOI
# oUcZd83oyTHr0ig5IFy1r9FpnG00/P5MV+zxmTbTDXJjC8VgxqWl2IhnPk8zr0Fc
# JK0BoYHtv7NHeC4WbNHQZCQf9UMSDALcVR23YZemWizmEK2Mclhjv0E+s7mLZn0A
# K03zCQSvwQrjt+2YzS7J8MxWlRA5cNj1bNbnTtIuEUPpLSYgsN8Q+Ks9ffk9D7yU
# t8No/ntuf6R38t/33c0LTCSJ9AIgjz7hUHMwggUnMIIED6ADAgECAhMwAAW/Xff+
# 6WMO1wIRAAAABb9dMA0GCSqGSIb3DQEBCwUAMFExCzAJBgNVBAYTAlVTMR8wHQYD
# VQQKDBZNSVQgTGluY29sbiBMYWJvcmF0b3J5MQwwCgYDVQQLDANQS0kxEzARBgNV
# BAMMCk1JVExMIENBLTYwHhcNMjQxMDI4MTgxMjU1WhcNMjcxMDI4MTgxMjU1WjBg
# MQswCQYDVQQGEwJVUzEfMB0GA1UEChMWTUlUIExpbmNvbG4gTGFib3JhdG9yeTEO
# MAwGA1UECxMFT3RoZXIxIDAeBgNVBAMTF0lTRCBEZXNrdG9wIEVuZ2luZWVyaW5n
# MIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEA0BQ5+bMtDvgRT7pCIgHp
# b0iuWsrGHTAKWvKo3T6uk/5r/Kp7VtqJFvcuwLqu0jm+As1kypxloyme0GAKCZcm
# nvyEtRIS5Vxn0FpPO1/y1Bm1JOZ30O7xoy3kimp/16jSmROMeCSdm9qPEmG60M5Y
# L12k7DOaU6/v+5MSZLQiDl20lf34u+Qt8SYNe/L4oA4kdsN3YMXuM6MVbbh6CJzb
# wBT3ceZNwRmkkqQOEQtA0Zr0n2UmoijuraIxU5DC+pISBJIcF3RbfFQNQMivR0lq
# rzQZDrKej/3D9FouGiBl8xZyVtJE0cNum6OE8b7nABtYwKP4jvz3ttxtIWVhoC/v
# WQIDAQABo4IB5zCCAeMwPQYJKwYBBAGCNxUHBDAwLgYmKwYBBAGCNxUIg4PlHYfs
# p2aGrYcVg+rwRYW2oR8dhuHfGoHsg1wCAWQCAQQwFgYDVR0lAQH/BAwwCgYIKwYB
# BQUHAwMwDgYDVR0PAQH/BAQDAgeAMBgGA1UdIAQRMA8wDQYLKoZIhvcSAgEDAQYw
# HQYDVR0OBBYEFLlL4q2UwnJN7ZTZ9W2D+7Y9a+tdMIGCBgNVHREEezB5pFswWTEY
# MBYGA1UEAwwPQW50aG9ueS5NYXNzYXJvMQ8wDQYDVQQLDAZQZW9wbGUxHzAdBgNV
# BAoMFk1JVCBMaW5jb2xuIExhYm9yYXRvcnkxCzAJBgNVBAYTAlVTgRpBbnRob255
# Lk1hc3Nhcm9AbGwubWl0LmVkdTAfBgNVHSMEGDAWgBSTkEcDQsFNtDMe1EKXn6Hg
# K3C/iDAzBgNVHR8ELDAqMCigJqAkhiJodHRwOi8vY3JsLmxsLm1pdC5lZHUvZ2V0
# Y3JsL2xsY2E2MGYGCCsGAQUFBwEBBFowWDAtBggrBgEFBQcwAoYhaHR0cDovL2Ny
# bC5sbC5taXQuZWR1L2dldHRvL2xsY2E2MCcGCCsGAQUFBzABhhtodHRwOi8vb2Nz
# cC5sbC5taXQuZWR1L29jc3AwDQYJKoZIhvcNAQELBQADggEBAFqyP/3MhIsDF2Qu
# ThdPiYz24768PIl64Tiaz8PjjxPnKTiayoOfnCG40wsZh+wlWvZZP5R/6FZab6ZC
# nkrI9IObUZdJeiN4UEypO1v5L6J1iXGq4Zc3QpkJUmjCIIYU0IPG9BPo0SX7mBiz
# DFafAGHReYkovs6vq035+4I6tsOQBpl+JfFPIT37Kpy+PlKz/OXzhVmQOa87mC1b
# YADxWAwwDJd1Mm1GFbXUHHBPkdusW+POqR7qh5WQf0dJpRTsMG/MzIqWiUZxDzkD
# lsqyRl4Y9nN9ii92PGpJF59AZAuEHDX0fqP6yeyMWYZGKpy7XqhQidW7nPxeqHl+
# EQW6EH0xggIJMIICBQIBATBoMFExCzAJBgNVBAYTAlVTMR8wHQYDVQQKDBZNSVQg
# TGluY29sbiBMYWJvcmF0b3J5MQwwCgYDVQQLDANQS0kxEzARBgNVBAMMCk1JVExM
# IENBLTYCEzAABb9d9/7pYw7XAhEAAAAFv10wCQYFKw4DAhoFAKB4MBgGCisGAQQB
# gjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYK
# KwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFDhkah8b
# NxZQ9a2hlz1//Euc9xNLMA0GCSqGSIb3DQEBAQUABIIBADr4i3dTDe9r7T/TB3ki
# ID0+bqyl2DWB6IWKnFGyrHcX2q+fX2o4Y+GmXIATPgzjfA5i3zAnCXAjI60snlB2
# pT+CLI8gwvLbrZQCipuObPJXVRVNpARPamxhYJjo+K3ZCWKi3wQJy9IIH7GQlJZH
# OHGAti/fzpClkt7Lu0XqidITFxB9cSzK8Rfd2B8hbxOVfZ6fdU24G701ot6k6v4D
# +Vqn8EcfcNZSyNIoczQMQj1k7OKvLKTIkWG6NdL1SDddrEpfwAxu8gyIlCh6Bd1L
# NNNpn1v3J1S0i9R1Ex8ZlMFlulQEZ9NPFFxsdax+yHYHJ4FCjnP5LIc4e07eOqCI
# IIE=
# SIG # End signature block
