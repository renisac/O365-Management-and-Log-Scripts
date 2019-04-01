$AuthenticationConfig=(Get-Content $PSScriptRoot\Authentication.json|Out-String|ConvertFrom-Json)
$MessageTraceConfig=(Get-Content $PSScriptRoot\MessageTraceConfiguration.json|Out-String|ConvertFrom-Json)
$Hours=@('00','01','02','03','04','05','06','07','08','09','10','11','12','13','14','15','16','17','18','19','20','21','22','23')
$Minutes=@('00:00','15:00','30:00','45:00','59:59')
$TimesToCollect=@()
$Credential=New-Object System.Management.Automation.PSCredential ($AuthenticationConfig.Username, ($AuthenticationConfig.Password|ConvertTo-SecureString))
$O365ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $Credential -Authentication Basic -AllowRedirection
$OutputPath=Join-Path $PSScriptRoot $MessageTraceConfig.OutputFolder

Import-Module (Import-PSSession $O365ExchangeSession -AllowClobber ) -Global

$TimesToCollect=@()
for ($i=0;$i -lt $MessageTraceConfig.DaysToPull;$i++){
    foreach ($Hour in $Hours){
        foreach ($Minute in $Minutes){
            $time=$(get-datE -Format yyyy-MM-dd -Date (get-date).AddDays(-$i))+" "+$Hour+":"+$Minute
            $TimesToCollect+=$time
            #write-host $time
        }
    }
    

}
$localfiles = @()
    $localFiles = Get-ChildItem $MessageTraceConfig.OutputFolder -Recurse | Select-Object -Property Name
    $localFilesHash=@{}
    $localfiles|%{
        $localFilesHash.Add($_.name,"" )
    }

for ($t=0; $t -lt $TimesToCollect.Count -1; $t++){
    $FileName="$("MessageTraceLog-"+$TimesToCollect[$t]+"-"+$TimesToCollect[$t+1]+".json")" -replace ":",""
    if ($localfilesHash.ContainsKey($FileName)){
        Write-host "Already have file $FileName"
    }
    elseif ( (get-date).AddHours(-1).ToUniversalTime() -lt $TimesToCollect[$t]   ){
        write-host "Too soon to pull $filename"
    }
    else
    {
        $ThisRun=@()
        $page=1
        do{
            $ThisRun+=Get-MessageTrace -StartDate $TimesToCollect[$t] -EndDate $TimesToCollect[$t+1] -Page $page -PageSize 5000 
            Write-host "Got Page $page of $filename"
            $page ++
        }while (($ThisRun.Count % 5000 -eq 0) -and $ThisRun.Count -ne 0 ) 
        if($ThisRun.count -ne 0){
            $temp=New-Object System.Collections.ArrayList
            foreach ($item in $ThisRun){
                $temp.add(( New-Object PSObject -Property @{
                    SenderAddress=  $item.SenderAddress
                    RecipientAddress=  $item.recipientaddress
                    Received=  ($item.Received.GetDateTimeFormats("s"))[0]
                    Status=  $item.Status
                    ToIP=  $item.ToIP
                    FromIP=  $item.FromIP
                    Subject=  $item.Subject
                    Size= $item.Size
                    MessageTraceId=  $item.MessageTraceID.Guid
                    MessageId= $item.MessageID -replace "<","" -replace ">",""
                    RecipientDomain= ($item.RecipientAddress -split "@")[1]
                    SenderDomain= ($item.SenderAddress -split "@")[1]
                }))|out-null
            }
            
            $Temp|select -ExcludeProperty RunspaceId, Organization, StartDate,EndDate,Index,PSComputerName -Property *Address*,Received, Status, *ip, Subject, Size, MessageTraceID, MessageID, *domain* `
            |ConvertTo-Json -Compress |Out-File ($MessageTraceConfig.OutputFolder + $FileName) -Force -Encoding utf8
            Write-host "New file: $FileName $($temp.count) items"
            #break

        }
    }
}
get-pssession|remove-pssession