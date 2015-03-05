<#
Description:
Goals:
Parse email types for valuable alerts

Email types:
1) Monitoring:
Detect type and frequency of emails 




#>


$olFolderInbox = 6
$outlook = new-object -com outlook.application;
$ns = $outlook.GetNameSpace("MAPI");
$inbox = $ns.GetDefaultFolder($olFolderInbox)
#checks 10 newest messages
$inbox.items | select -first 10 | foreach {
    if($_.unread -eq $True) {
    $mBody = $_.body
    #Splits the line before any previous replies are loaded
    $mBodySplit = $mBody -split "From:"
    #Assigns only the first message in the chain
    $mBodyLeft = $mbodySplit[0]
    #build a string using the –f operator
    $q = "From: " + $_.SenderName + ("`n") + " Message: " + $mBodyLeft
    #create the COM object and invoke the Speak() method 
    (New-Object -ComObject SAPI.SPVoice).Speak($q) | Out-Null
    } 
}




