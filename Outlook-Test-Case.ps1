Clear-Host
#$ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent
$ErrorActionPreference = "Stop"
$MyEmail = "muhammadsuhail.asrulsani@sc.com"
$DL = ""
$Invalid_SMTP_Domain = "njdnflsnfls@fdlksfjksd.com"
$SOB = "MohdHafizan.Ramlib@sc.com"

Function Outlook-Send-Email {
    try {
        Write-Host "Sending email with subject 'Test Email 1' : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        $mail = $ol.CreateItem(0)
        $mail.To = $MyEmail
        $mail.Subject = "Test Email 1"
        $mail.Body = "This is a test email 1"
        $mail.Send()
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
    }
}

Function Outlook-Send-Email-Attachment {
    try {
        Write-Host "Sending email with subject 'Test Email 1' and attachment : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        $mail = $ol.CreateItem(0)
        $mail.To = $MyEmail
        $mail.Subject = "Test Email 1 with attachment"
        $mail.Body = "This is a test email 1 with attachment"
        $Attachment = "C:\Users\2003686\OneDrive - Standard Chartered Bank\Desktop\BitBucket\Outlook-Test-Case\file1.txt"
        $Mail.Attachments.Add($Attachment) | Out-Null
        $mail.Send()
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
    }
}

Function Outlook-Send-Email-DL {
    try {
        Write-Host "Sending email with subject 'Test Email 1' : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        $mail = $ol.CreateItem(0)
        $mail.To = $DL
        $mail.Subject = "Test Email 1"
        $mail.Body = "This is a test email 1"
        $mail.Send()
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
    }
}

Function Outlook-Send-Email-InvalidSMTPDomain {
    try {
        Write-Host "Sending email with subject 'Test Email 1' : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        $mail = $ol.CreateItem(0)
        $mail.To = $Invalid_SMTP_Domain
        $mail.Subject = "Test Email 1"
        $mail.Body = "This is a test email 1"
        $mail.Send()
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
    }
}

Function Outlook-Send-Email-deliveryread-receipts {
    try {
        Write-Host "Sending email with subject 'Test Email 1' and read/delivery receipts : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        $mail = $ol.CreateItem(0)
        $mail.To = $MyEmail
        $mail.Subject = "Test Email 1 with read/delivery receipts"
        $mail.Body = "This is a test email 1 with read/delivery receipts"
        $mail.OriginatorDeliveryReportRequested = $true
        $mail.ReadReceiptRequested = $true
        $mail.Send()
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
    }
}

Function Outlook-Reply-Email {
    try {
        Write-Host "Replying email with subject 'Test Email 1' : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        $Namespace = $ol.GetNamespace("MAPI")
        $Mailbox = $Namespace.GetDefaultFolder(6)
        $Email = $Mailbox.Items.Find("[Subject]='Test Email 1'")
        $Reply = $Email.Reply()
        $Reply.Body = "This is a reply to the test email with subject 'Test Email 1'."
        $Reply.Send()
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
    }
}

Function Outlook-Forward-Email {
    try {
        Write-Host "Forwarding email with subject 'Test Email 1' : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        $Namespace = $ol.GetNamespace("MAPI")
        $Mailbox = $Namespace.GetDefaultFolder(6)
        $Email = $Mailbox.Items.Find("[Subject]='Test Email 1'")
        $Forward = $Email.Forward()
        $Forward.Body = "This is a forward to the test email with subject 'Test Email 1'."
        $Forward.To = $MyEmail
        $Forward.Send()
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
    }
}

Function Outlook-Enable-OOF {
    try {
        Write-Host "Enabling Out of office : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        $Namespace = $ol.GetNamespace("MAPI")
        $Mailbox = $Namespace.GetDefaultFolder(6)
        $ExchangeUser = $Mailbox.
        $ExchangeUser.SetOutOfOfficeAssistant($true, "I am currently out of the office.")
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
    }
}

Function Outlook-Send-Email-High {
    try {
        Write-Host "Sending email with subject 'Test Email 1 High Priority' with high priority : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        $mail = $ol.CreateItem(0)
        $mail.To = $MyEmail
        $mail.Subject = "Test Email 1 High Priority"
        $mail.Body = "Test Email 1 High Priority"
        $mail.Importance = 2
        $mail.Send()
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
    }
}

Function Outlook-Send-Email-OnBehalf {
    try {
        Write-Host "Sending on behalf email with subject 'Test Email 1' : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        $mail = $ol.CreateItem(0)
        $mail.To = $MyEmail
        $mail.Subject = "Test Email 1 onbehalf"
        $mail.Body = "Test Email 1 onbehalf"
        $mail.Importance = 1
        $mail.SentOnBehalfOfName = $SOB
        $mail.Send()
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
    }
}

Function Outlook-Set-OOF {
    try {
        Write-Host "Setting Out Of Office : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        $receipient = $ol.Session.CurrentUser
        $receipient.AutoReply.SetAutoReply("Currently Out of Office","Currently Out of Office", $true)
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
    }
}

Function Outlook-Move-Email-PST {
    try {
        Write-Host "Moving message with subject 'Test Email 1' to pst folder 'Exported' : " -NoNewline
        $exportFile = "C:\Users\2003686\OneDrive - Standard Chartered Bank\Documents\Outlook Files\MyEmail.pst" 
        $outlook = new-object -comobject outlook.application
        $namespace = $outlook.GetNameSpace("MAPI")
        $inbox = $namespace.GetDefaultFolder(6) 
        $namespace.AddStore($exportFile)
        $exportFolderID = ($namespace.folders | where{$_.FolderPath -eq "\\MyEmail"}).EntryID
        $exportPST = $namespace.GetFolderFromID($exportFolderID)
        $exportPSTFolder = $exportPST.Folders.Add("Exported") 
        $messages = $inbox.items | where{$_.Subject -like "*test email 1*"} 
        Foreach($Message in $Messages){$messages.Move($exportPSTFolder)}
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
    }
}

Function Outlook-SearchEmail {
    try {
        Write-Host "Searching email with subject 'Test Email 1' : " -NoNewline
        $outlook = new-object -comobject outlook.application
        $namespace = $outlook.GetNameSpace("MAPI")
        $inbox = $namespace.GetDefaultFolder(6) 
        $messages = $inbox.items | where{$_.Subject -like "*test email 1*"} 
        $found = Foreach($Message in $Messages) {
            $messages.Subject
        }

        if ($found) { Write-Host "Email exist" -ForegroundColor Green }
        if (!$found) { Write-Host "Email not exist" -ForegroundColor Yellow }
        #Write-Host "Success" -ForegroundColor Green
        }

    catch {
        Write-Warning ($_)
    }
}

Function Outlook-Access-OAB {
    try {
        Write-Host "Accessing OAB in online mode : " -NoNewline
        $ol = new-object -comobject outlook.application
        $namespace = $outlook.GetNameSpace("MAPI")
        $GAL = $namespace.GetGlobalAddressList()
        $found = $GAL.AddressEntries | Select-Object -First 1
        if ($found) { Write-Host "OAB can be access in Online mode" -ForegroundColor Green }
        if (!$found) { Write-Host "OAB cannot be access in Online mode" -ForegroundColor Yellow }

        Write-Host "Switching Outlook to Offline : " -NoNewline
        ($ol.ActiveExplorer()).CommandBars.ExecuteMso("ToggleOnline")
        Start-Sleep 5
        Write-Host "Success" -ForegroundColor Green
 
        Write-Host "Accessing OAB in offline mode : " -NoNewline
        $found = $GAL.AddressEntries | Select-Object -First 1
        if ($found) { Write-Host "OAB can be access in Offline mode" -ForegroundColor Green }
        if (!$found) { Write-Host "OAB cannot be access in Offline mode" -ForegroundColor Yellow }

        Write-Host "Switching Outlook back to Online : " -NoNewline
        ($ol.ActiveExplorer()).CommandBars.ExecuteMso("ToggleOnline")
        Start-Sleep 5
        Write-Host "Success" -ForegroundColor Green
        }

    catch {
        Write-Warning ($_)
    }
}

