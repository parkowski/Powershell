FUNCTION CopyMove-Email {
    param
    (
    [string[]]$acronyms,
    [string[]]$folders,
    [int]$sent_max_keep
    )

    Add-type -assembly "Microsoft.Office.Interop.Outlook"
    $Outlook = New-Object -comobject Outlook.Application
    $namespace = $Outlook.GetNameSpace("MAPI")
    $olFolderInbox = 6
    $inbox = $ns.GetDefaultFolder($olFolderInbox)

    #construct a timespan to mark the absolute limit for keeping sent items in that folder, whether the item is marked for copying or not.
    $timespanSentMax = New-Object TimeSpan($sent_max_keep,0,0,0,0)

    #Using the timespan, we create the dates against which to test each e-mail item in the Sent Items folder.
    $Date = New-Object DateTime
    $SentMaxDate = New-Object DateTime
    $Date = Get-Date
    $SentMaxDate = $Date.Subtract($timespanSentMax)

    #We also need a string-formatted date to use in naming text files that will store a specific day's run of this application. The text files will store, respectively, mail that was copied and mail that was moved.

    $dt = $Date.ToString().Replace('/','')
    $dt = $dt.Replace(' ','')
    $dt = $dt.Replace(':','')

    #Then we need to accumulate data about moves and copies in temporary holding arrays:

    $array_of_move_results = @()
    $array_of_copy_results = @()

    #Experience shows that merely looping through all the messages in the Sent Items folder once isn’t enough; on a first loop, some items are handled but others aren’t touched. It may take three or four such loops to handle all the items that need handling. The number of loops probably depends on how many mail items are ready for copy-move operations and other factors in how Outlook interoperates with your computer. In any event, the solution is to use a Do-While structure to keep running the loop until all marked items have been properly managed.

    do {
        $hit = 0
        #The application is now ready to begin parsing through Sent Items:

        foreach($mail in $inbox) {
        #Note that joe.leibowitz@companyname.com is the author's Outlook name for the group of folders that includes Sent Items. 
        #You should, of course, use the name that shows up in your own Outlook client.
        #The first If test is against the established maximum-days window and for the existence of either the \\\ tag or the /// tag:

            if(
                $mail.SentOn -ge $SentMaxDate -and
                (
                ($mail.Subject.IndexOf('///') -gt - 1) -or
                ($mail.Subject.IndexOf('\\\') -gt - 1)
                )
              ) {
                $acronym_ctr = - 1
                #With a matching e-mail in hand, test it for each acronym in the $acronyms parameter:

                #test the mail for all acronyms
                foreach($acronym in $acronyms) {
                    $acronym_ctr += 1
                    #get the matching Folder for this acronym
                    $fldr = $folders[$acronym_ctr]
                    if($mail.Subject.ToUpper().IndexOf($acronym.ToUpper())  -gt - 1) {

                        #Copying Code Based on Acronyms
                        #For each matching acronym, the script makes a copy of the subject e-mail and moves the copy to the folder that matches the acronym. This part of the code is shown in Figure 2. Essentially, this copies the message to the new folder while also leaving it in the original Sent Items folder.

                        $MailCopy = $Outlook.CreateItem(0)
                            $MailCopy = $mail.Copy() 
                            $results =
                            $MailCopy.Move($namespace.Folders.Item('joe.leibowitz@abbott.com').Folders.Item($fldr))
                            write-host $results
                            $str = $mail.SenderName + ' ' +
                            $mail.SentOn + ' '  +
                            $mail.Subject
                            $array_of_copy_results += @($str,' to: ',$fldr,' ','from: Sent Items')
                        }
                }

                #Moving Items to Sent Items OLD
                $target_folder2 = 'C:\SentMailMoved_' + $dt + '_.txt'
                #having done the copies,move it to 'OLD' folder for storage
                $results2 =
                $mail.Move($namespace.Folders.Item('joe.leibowitz@abbott.com').Folders.Item('Sent
                Items OLD'))
                write-host $results2
                $str = $mail.SenderName + ' ' +
                $mail.SentOn + ' '  +
                $mail.Subject
                $array_of_move_results += @($str,' to: Sent Items OLD'' ','from: Sent Items')
                $item_handled = $true
                $hit = 1
            }

            #Managing Message Retention Limits
            elseif($mail.SentOn -lt $SentMaxDate) {
                $target_folder3 = 'C:\SentMailMoved_' + $dt + '_.txt'   #.txt'
                $results =
                    $mail.Move($namespace.Folders.Item('joe.leibowitz@abbott.com').Folders.Item('Sent Items OLD'))
                write-host $results
                $str = $mail.SenderName + ' ' +
                $mail.SentOn + ' '  +
                $mail.Subject
                $array_of_move_results += @($str,' to: Sent Items OLD',' ','from: Sent Items')
                $hit = 1

            }
#}
        }
    }
    while($hit -ne 0)

    $target_folder3 = 'C:\SentMailMoved_' + $dt + '_.txt'
    'Moved from Sent Items to Sent Items OLD: ',$array_of_move_results >> $target_folder3
    $target_folder = 'C:\SentMailCopied_' + $dt + '_.txt'
    'Copied items: ',$array_of_copy_results >> $target_folder
}

CopyMove-Email ('Admin','PRJ','FOR','DOM')('Administrative','Projects','Foreign','Domestic') 15
