# OCBR---Office-Customization-Backup-and-Restore
GUI tool to automate the backup and restoration of Microsoft Office customizations.

<img width="581" height="748" alt="image" src="https://github.com/user-attachments/assets/fe09ce90-52ed-47e8-bee9-7ba4e489dfe9" />



Command line to backup:
powershell.exe -ExecutionPolicy Bypass -File "C:\Path\On\Client\To\OfficeCustomizationTool.ps1" -Action backup -Path $BackupPath -Items "RibbonUI","Templates","Signatures","Dictionaries"

Command line to restor:
powershell.exe -ExecutionPolicy Bypass -File "C:\Path\On\Client\To\OfficeCustomizationTool.ps1" -Action restore -Path $BackupPath -Items "RibbonUI","Templates","Signatures","Dictionaries"
