# OCBR-Office-Customization-Backup-and-Restore
GUI tool to automate the backup and restoration of Microsoft Office customizations.

<img width="581" height="766" alt="image" src="https://github.com/user-attachments/assets/dedec946-f5c0-4b23-b84a-37f272eca6c9" />


.SYNOPSIS
    A comprehensive tool for administrators to back up and restore a user's Microsoft Office settings.
    Runs elevated but automatically targets the active desktop user.

.DESCRIPTION
    This script is designed to be run as SYSTEM or an Administrator. It automatically detects the
    currently active desktop user and targets their profile for backup and restore of key Office settings.

.VERSION
    1.1.0

.PARAMETER UserName
    Optional. Explicitly specifies the username to target (e.g., 'jdoe'), overriding auto-detection.

.PARAMETER Action
    For command-line use. Must be either 'backup' or 'restore'.

.PARAMETER Path
    For command-line use. The root folder for the backup/restore operation.

.PARAMETER Items
    For command-line use. An array of items to process: 'RibbonUI', 'Templates', 'Signatures', 'Dictionaries', 'AutoComplete', 'ExcelMacros', 'AutoCorrect'.

.EXAMPLE
    # Back up ALL supported settings for the active user via command line.
    .\OfficeCustomizationTool.ps1 -Action backup -Path "C:\Backups" -Items "RibbonUI","Templates","Signatures","Dictionaries","AutoComplete","ExcelMacros","AutoCorrect"
