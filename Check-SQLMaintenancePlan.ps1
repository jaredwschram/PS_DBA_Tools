#Load SMO Extension
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.SqlServer.SMO') | Out-Null
#Defining path for export as its used in more than 1 function
$export = 'C:\PATH\TO\REPORTS\HealthCheck.xlsx'
#building a function for mail message so if a server has a failure we can change message priority - in Prod this loads all parameters from a file and has auth/encryption setup
function mailMessage{
    #$dba = 'DBA_Support@DOMAIN.com'
    $dba = 'jared.schram@DOMAIN.com'
    $pass = Get-Content C:\PATH\TO\SCRIPTS\script_Dependencies\pass.txt | ConvertTo-SecureString
    $cred = new-object -typename System.Management.Automation.PSCredential -argumentlist ('DOMAIN\USER',$pass)
    Send-MailMessage -To $dba -from 'DBA_NoReply@DOMAIN.com' -subject 'SQL Maintenance Plan Report' -Attachments $export -body 'This is an automated report to check the status of the overnight SQL Maintenace Plan job on every SQL Instance' -SmtpServer smtp.DOMAIN.com -Credential $cred
}
#Creates the Excel doc
$xl = New-Object -ComObject excel.application
#$true launches excel on your desktop where $false will run w/o launching excel, useful for testing
$xl.Visible = $false
#Adds a new sheet to the Excel doc
$wb = $xl.WorkBooks.Add()
#Selects sheet1 in document
$ws = $wb.WorkSheets.Item(1)
#Starting row/column values for working with Excel
$row = 3
$col = 2
#Color/Message Key
$cells = $ws.Cells
$cells.item(1, 3).Font.Bold = $true
$cells.item(1, 3).Font.Size = 18
$cells.item(1, 3) = "Maintenance Plan Check"
$cells.item(3, 9) = "Step successful"
$cells.item(3, 8).Interior.ColorIndex = 4
$cells.item(4, 9) = "Step failed"
$cells.item(4, 8).Interior.ColorIndex = 3
$cells.item(5, 9) = "Cannot Connect, No Job or Job Disabled"
$cells.item(5, 8).Interior.ColorIndex = 6
#Creating column headers
$cells.item($row, $col) = "Server"
$cells.item($row, $col).Font.Size = 16
$Cells.item($row, $col).Columnwidth = 30
$col++
$cells.item($row, $col) = "Step Name"
$cells.item($row, $col).Font.Size = 16
$Cells.item($row, $col).Columnwidth = 25
$col++
$cells.item($row, $col) = "Message"
$cells.item($row, $col).Font.Size = 16    
$Cells.item($row, $col).Columnwidth = 30
$col++
$cells.item($row, $col) = "Job Name"
$cells.item($row, $col).Font.Size = 16    
$Cells.item($row, $col).Columnwidth = 15
$col++
$cells.item($row, $col) = "Run Date"
$cells.item($row, $col).Font.Size = 16    
$Cells.item($row, $col).Columnwidth = 15
$col++
$row++
function failExcel{
    $col = 2
    $cells.item($row, $col) = $srv
    $col++
    $col++
    $cells.item($row, $col) = $msg
    $cells.item($row, $col).Interior.ColorIndex = $color
}

function getSQLJobInfo{
    #If we don't remove the existing one first it will prompt to overwrite it and unless you say yes it will hang forever
    if(Test-Path $export){
        Remove-Item $export
    }
    #we only check for jobs that ran greater than yesterday(i.e. Today)
    $date = (Get-Date).AddDays(-1)
    #our list of SQL instances to check against
    $srvList = Import-Csv -Path 'C:\PATH\TO\SCRIPTS\\script_Dependencies\allServerInstances.csv'
    foreach($srv in $srvList.serverInstances){
        #Create SMO object for current SQL instance in list
        $smo = New-Object ('Microsoft.SqlServer.Management.Smo.server') $srv
        try{       
            #ConnectionContext attempts a connection to the SQL instance using SMO
            $smo.ConnectionContext.Connect()
            #JobServer.Jobs class lists all details about all Jobs on the instance
            $allJobs = $smo.JobServer.Jobs
            if($allJobs.Name -like "Maintenance*"){
                foreach($job in $allJobs | Where-Object name -like "Maintenance*"){
                    if($job.isEnabled -eq $true){
                        #EnumHistory() method lists the entire history of a job and all its steps so we have to compare it against $date
                        $jobHist = $job.EnumHistory() | Where-Object RunDate -gt $date
                        #running through each step in the current job
                        foreach($step in $jobHist){
                            $stepName = $step.StepName
                            $eMessage = $step.Message
                            $jobName = $step.JobName
                            $runDate = $step.RunDate
                            #Every job has a final step called Job Outcome that we dont want because our job continues even on previous step failure
                            if($stepName -notlike "(Job Outcome)"){
                                #setting color of cell based on the job message whether it failed or not
                                if($eMessage -like "*failed*"){
                                    $msg = "Failed"
                                    $color = 3
                                }else{
                                    $msg = "Success"
                                    $color = 4
                                }
                                #incrementing row and building the excel data based on the step we're on
                                $row++
                                $col = 2
                                $i++
                                #$cells.Borders.Item($
                                $cells.item($row, $col) = $srv
                                $col++
                                $cells.item($row, $col) = $stepName
                                $col++
                                $cells.item($row, $col) = $msg
                                $cells.item($row, $col).Interior.ColorIndex = $color
                                $col++
                                $cells.item($row, $col) = $jobName
                                $col++
                                $cells.item($row, $col) = $runDate
                            }
                        }
                    }else{
                        $row++
                        $msg = "Job is disabled"
                        $color = 6
                        failExcel
                    }
                }
            }else{
                $row++
                $msg = "Job does not exist on server"
                $color = 6
                failExcel    
            }
        }catch{
            #Update excel w/ no job ran
            $row++
            $msg = "Could not connect to instance"
            $color = 6
            failExcel
        }finally{
            #While SMO by default will close the connection once no longer in use the ServerConnect class will not so we must call the Disconnect() method
            $smo.ConnectionContext.Disconnect()
        }
        #adding space between the instance
        $row++
    }
    #Saving excel doc and then forcing excel to stop - it won't ever stop unless we do this
    $wb.SaveAs($export)
    $xl.quit()
    Stop-Process -ProcessName EXCEL
}
getSQLJobInfo
Start-Sleep -Seconds 60
mailMessage