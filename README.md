# OCBR---Office-Customization-Backup-and-Restore
GUI tool to automate the backup and restoration of Microsoft Office customizations.

<img width="581" height="748" alt="image" src="https://github.com/user-attachments/assets/65d41a2c-2182-43de-9ec0-f5375e2b0619" />


.SYNOPSIS
    A comprehensive tool for administrators to back up and restore a user's Microsoft Office settings.
    Runs elevated but automatically targets the active desktop user.

.DESCRIPTION
    This script is designed to be run as SYSTEM or an Administrator. It automatically detects the
    currently active desktop user and targets their profile for backup and restore of key Office settings.

.VERSION
    1.0.0

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
