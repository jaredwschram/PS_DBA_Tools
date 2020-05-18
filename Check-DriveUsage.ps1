#Author - Jared Schram

#Function to email report. Sends as HTML which makes it much easier to read and works on mobile
function mailMessage{
    $dba = 'DBA_Support@DOMAIN.com'
    #$dba = 'jared.schram@DOMAIN.com'
    $pass = Get-Content C:\PATH\TO\SCRIPTS\pass.txt | ConvertTo-SecureString
    $cred = new-object -typename System.Management.Automation.PSCredential -argumentlist ('DOMAIN\USER',$pass)
    Send-MailMessage -To $dba -from 'DBA_NoReply@DOMAIN.com' -subject 'UNIQUE daily disk space check' -BodyAsHtml "$tableOutput" -SmtpServer smtp.DOMAIN.com -Credential $cred
}

$servers = Import-csv 'C:\PATH\TO\SCRIPTS\UNIQUECheck.csv' 
#creating empty array to fill from server results
$Output = @()
foreach ($srv in $servers){
    #we get all the drives from the current server and then for each one grab specific properties and do some math to get the disk size into human readable format
    #Ideally i think this should be rewritten into a foreach loop instead as it creates cleaner code. Time necessitated using foreach-object
    $Output += Get-WmiObject Win32_Volume -Filter "DriveType='3'" -ComputerName $srv.Name | Where-Object Label -NotLike "System Reserved"| ForEach-Object {
        New-Object PSObject -Property @{
            Name = $_.Name
            Computer = $srv.Name
            #we're rounding to the nearest number of decimal places (2 in this case)
            FreeSpace_GB = ([Math]::Round($_.FreeSpace /1GB,2))
            TotalSize_GB = ([Math]::Round($_.Capacity /1GB,2))
            UsedSpace_GB = ([Math]::Round($_.Capacity /1GB,2)) - ([Math]::Round($_.FreeSpace /1GB,2))
        }
    }
}
#CSS applied to the HTML table so the report is legible can get robust in styling but not sure if we can highlight specific cells - probably need to use XLSX for that
$Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@
$tableOutput = $Output | ConvertTo-Html -Head $Header
mailMessage