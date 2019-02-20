<#
  .SYNOPSIS
  Used to build Azure Backup reports from the tenant
  Inserts them into an Excel Workbook with individual worksheets.

  .DESCRIPTION
   Saves workbook is same directory as script. This script requires the ImportExcel module for PowerShell.
   https://github.com/dfinke/ImportExcel
   https://www.powershellgallery.com/packages/ImportExcel/4.0.11


  .OUTPUTS
  MARS_Backup_02-19-2019.xlsx

  .NOTES
#>

#variables - change per site
$client = 'ABIN'


#static variables
$arr01 = @()
$arr02 = @()
$date = Get-Date -Format MM-dd-yyyy
$path = "C:\scripts\" + $client + "_BackupReport_" + $date + ".xlsx"
$subjectpass = "[SUCCESS] - " + $client + " Recovery Point Report for " + $date
$subjectfail = "[FAILURE] - " + $client + " Recovery Point Report for " + $date
