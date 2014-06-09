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




#Body
#TaskSubject                       : RE: BMC Monitoring for ORCA Virtuals
#SenderEmailAddress
#SenderName                        : Kato, Stephen N
#Sent                              : True
#SentOn                            : 2/14/2013 3:28:46 PM
#SentOnBehalfOfName                : Kato, Stephen N
#ReceivedByName                    : Park, Yong K
#ReceivedTime                      : 2/14/2013 3:28:47 PM
#HTMLBody
#Sensitivity                       : 0
#Size                              : 17502
#Subject                           : RE: BMC Monitoring for ORCA Virtuals 
#UnRead                            : False
#BCC                               : 
#CC                                : Civarra, Larry
#ConversationTopic                 : BMC Monitoring for ORCA Virtuals 
#CreationTime                      : 2/14/2013 3:28:47 PM