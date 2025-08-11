# OCBR---Office-Customization-Backup-and-Restore
GUI tool to automate the backup and restoration of Microsoft Office customizations.

<img width="560" height="666" alt="image" src="https://github.com/user-attachments/assets/6faf4037-1bd2-4018-bbc8-f7181bf20907" />


Command line to backup:
powershell.exe -ExecutionPolicy Bypass -File "C:\Path\On\Client\To\OfficeCustomizationTool.ps1" -Action backup -Path $BackupPath -Items "RibbonUI","Templates","Signatures","Dictionaries"

Command line to restor:
powershell.exe -ExecutionPolicy Bypass -File "C:\Path\On\Client\To\OfficeCustomizationTool.ps1" -Action restore -Path $BackupPath -Items "RibbonUI","Templates","Signatures","Dictionaries"
