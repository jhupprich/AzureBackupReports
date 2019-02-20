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
$admin = "jhupprich@dougrwgci.onmicrosoft.com" 
$securePass = "Prince0f$pace!" | ConvertTo-SecureString -AsPlainText -Force
$subID = "d6a3f26f-407a-49ce-b0ec-97f6b7109d6f"
$vault = "NSEDC-File"

#static variables
$cred = New-Object -TypeName System.Management.Automation.PSCredential `
-argumentlist $admin,$securePass
$arr01 = @()
$arr02 = @()
$date = Get-Date -Format MM-dd-yyyy
$path = "C:\scripts\" + $client + "_BackupReport_" + $date + ".xlsx"
$subjectPass = "[SUCCESS] - " + $client + " Recovery Point Report for " + $date
$subjectFail = "[FAILURE] - " + $client + " Recovery Point Report for " + $date


#connect and fetch data
Login-AzureRmAccount -Credential $cred -Subscription $subID

$vault = Get-AzureRmRecoveryServicesVault -name $vault
Set-AzureRmRecoveryServicesVaultContext -Vault $vault


