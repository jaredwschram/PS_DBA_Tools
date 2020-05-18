#Author - Jared Schram
$sL = Import-Csv C:\PATH\TO\SCRIPTS\allDBAServerNames.csv
$date = (Get-Date).ToString("MM-dd-yyyy")
#CSS for the HTML table that gets sent in the event there is an error
$Header = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #cc0000;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@
#Variables for our incrementor and error array that are used in the catch block. 
$e = 0
$errOut = @()
function mailMessage{
    $dba = 'DBA_Support@DOMAIN.com'
    #$dba = 'jared.schram@DOMAIN.com'
    $pass = Get-Content C:\PATH\TO\SCRIPTS\pass.txt | ConvertTo-SecureString
    $cred = new-object -typename System.Management.Automation.PSCredential -argumentlist ('DOMAIN\USER',$pass)
    if($message -eq 1){
        $body = $tableOutput
    }else{
        $body = "Daily disk space check ran without errors. The full data dump can be found at C:\PATH\TO\REPORTS\dbaDriveCheckOutput.csv On the LOCATION"
    }
    Send-MailMessage -To $dba -from 'DBA_NoReply@DOMAIN.com' -subject 'Daily disk space check job' -BodyAsHtml "$body" -SmtpServer smtp.DOMAIN.com -Credential $cred
}
#needed to be resused since we have to account for Windows Server 2003 
function createRow{
    foreach($drive in $driveData){
        $obj = New-Object PSObject -Property @{
            Date = $date
            DriveLetter = $drive.Name
            HostName = $server
            #we're rounding to the nearest number of decimal places (2 in this case)
            FreeSpace_GB = ([Math]::Round($drive.FreeSpace /1GB,2))
            TotalSize_GB = ([Math]::Round($drive.Capacity /1GB,2))
            Percentage_Free = [Math]::Round($drive.FreeSpace / $drive.Capacity * 100,2)
        }
        $obj | Export-Csv -Append -Path C:\PATH\TO\REPORTS\dbaDriveCheckOutput.csv
    }
}
foreach($server in $sL.serverNames){
    try{
        #using ActiveDirectory to determine if a computer is Windows Server 2003 or not
        $ad = Get-ADComputer $server -Properties OperatingSystem
        if($ad.OperatingSystem -like "Windows Server 2003"){
            #For some reason Windows Server 2003 boxes weren't letting me query WMI directly but Invoke-Command works
            $driveData = Invoke-Command -ComputerName $server -ScriptBlock {Get-WmiObject Win32_Volume -Filter "DriveType='3'"}
            createRow
        }else{
            $driveData = Get-WmiObject Win32_Volume -Filter "DriveType='3'" -ComputerName $server | Where-Object Label -NotLike "System Reserved"
            createRow
        }
    }catch{
        #building array of errors
        $errOut += New-Object PSOBject -Property @{
            HostName = $server
            Error = $Error[$e].Exception.Message
        }
        $e++
    }

}
#if the error array contains values we convert it to HTML and email it if not we email a success message
if($errOut){
    $tableOutput = $errOut | ConvertTo-Html -Head $Header
    mailMessage 1
}else{
    mailMessage 2
}
