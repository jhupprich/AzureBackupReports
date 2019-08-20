<#
  .SYNOPSIS
  Used to build Azure Backup reports from DPM on premise
  Inserts them into an Excel Workbook with individual worksheets.

  .DESCRIPTION
   Saves workbook is same directory as script. This script requires the ImportExcel module for PowerShell.
   https://github.com/dfinke/ImportExcel
   https://www.powershellgallery.com/packages/ImportExcel/4.0.11


  .OUTPUTS
  MARS_Backup_02-19-2019.xlsx

  .NOTES
#>


## run as admin 
Import-Module 'C:\Program Files\Microsoft Azure Backup Server\DPM\DPM\bin\DpmCliInitScript.ps1'

#variables - change per site
$client = 'ABIN'


#static variables
$arr01 = @()
$arr02 = @()
$date = Get-Date -Format MM-dd-yyyy
$path = "C:\scripts\" + $client + "_BackupReport_" + $date + ".xlsx"
$subjectpass = "[SUCCESS] - " + $client + " Recovery Point Report for " + $date
$subjectfail = "[FAILURE] - " + $client + " Recovery Point Report for " + $date

$group = Get-DPMProtectionGroup
$data = Get-DPMDataSource -ProtectionGroup $group
foreach($d in $data)
    {
        $rp = Get-DPMRecoveryPoint -DataSource $d | select -Last 1 | `
        select Name, DataSource, BackupTime, Location, PhysicalPath
        $attr = New-Object System.Object
        $attr | Add-Member -Type NoteProperty -Name 'Name' -Value $d.Name
        $attr | Add-Member -Type NoteProperty -Name 'Computer' -Value $d.Computer
        $attr | Add-Member -Type NoteProperty -Name 'ObjectType' -Value $d.ObjectType
        $attr | Add-Member -Type NoteProperty -Name 'BackupTime' -Value $rp.BackupTime
        $attr | Add-Member -Type NoteProperty -Name 'Location'-Value $rp.Location
        $attr | Add-Member -Type NoteProperty -Name 'PhysicalPath'-Value $rp.PhysicalPath
        $arr01 += $attr
        if ($($rp.BackupTime) -lt ((Get-Date).AddHours( -24))) {$arr02 += 1}
    }


## Excel report
$txt01 = New-ConditionalText -Range "A1:F1" -ConditionalTextColor black -BackgroundColor silver 
$txt02 = New-ConditionalText -Range "D2:D26" -ConditionalType LessThan -Text "=Int(Now() - 1)" -ConditionalTextColor black `
-BackgroundColor red

$arr01 | Sort-Object -Property Computer | Export-Excel -Path $path -AutoSize `
-ConditionalText $txt01,$txt02

if($arr02.Count -gt 1){
    Send-MailMessage -To backup@youremail.com -From relayname@client.com -Subject $subjectfail `
    -SmtpServer relayserverFQDNorIP -Attachments $path
}

else{
    Send-MailMessage -To backup@youremail.com -From relayname@client.comm -Subject $subjectpass `
-SmtpServer relayserverFQDNorIP -Attachments $path
}
