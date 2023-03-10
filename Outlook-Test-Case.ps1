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
        $receipient.AutoReply.SetAutoReply("Currently Out of Office", "Currently Out of Office", $true)
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
    }
}

Function Outlook-Move-Email-PST {
    try {
        Write-Host "Moving message with subject 'Test Email 1' to pst folder 'Exported' : " -NoNewline
        $exportFile = "C:\Users\2003686\OneDrive - Standard Chartered Bank\Documents\Outlook Files\MyEmail.pst"??
        $outlook = new-object -comobject outlook.application
        $namespace = $outlook.GetNameSpace("MAPI")
        $inbox = $namespace.GetDefaultFolder(6)??
        $namespace.AddStore($exportFile)
        $exportFolderID = ($namespace.folders | Where-Object { $_.FolderPath -eq "\\MyEmail" }).EntryID
        $exportPST = $namespace.GetFolderFromID($exportFolderID)
        $exportPSTFolder = $exportPST.Folders.Add("Exported")??
        $messages = $inbox.items | Where-Object { $_.Subject -like "*test email 1*" }??
        Foreach ($Message in $Messages) { $messages.Move($exportPSTFolder) }
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
        $inbox = $namespace.GetDefaultFolder(6)??
        $messages = $inbox.items | Where-Object { $_.Subject -like "*test email 1*" }??
        $found = Foreach ($Message in $Messages) {
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

Function Outlook-Create-Rule {
    try {
        Write-Host "Creating Outlook Rule name Test Rule 1 : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        $namespace = $ol.GetNameSpace("MAPI")

        $rules = $namespace.DefaultStore.GetRules()

        $rule = $rules.Create("Test Rule 1", [Microsoft.Office.Interop.Outlook.OlRuleType]::olRuleReceive)


        $condition = $rule.Conditions.SenderAddress
        $condition.Enabled = $true
        $condition.Address = @($MyEmail)
        $Action = $rule.Actions.AssignToCategory
        $Action.Enabled = $true
        $Action.Categories = @("Red Category")

        $rules.Save()
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
    }
}

Function Outlook-Verify-FreeBusy {
    try {
        Write-Host "Verifying free/busy information : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        #$namespace = $ol.GetNameSpace("MAPI")
        #$calendar = ($ol.GetNamespace("MAPI")).GetDefaultFolder(9)
        $receipient = $ol.Session.CreateRecipient($MyEmail)
        $freebusy = $receipient.FreeBusy("3/6/2023 08:00", $true)
        if ($freebusy) { Write-Host "Success" -ForegroundColor Green }
    }

    catch {
        Write-Warning ($_)
    }
}

Function Calendar-Create-Meeting {
    try {
        Write-Host "Creating Test Meeting 1 : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        $meeting = $ol.CreateItem('olAppointmentItem')
        $meeting.Subject = 'Test Meeting 1'
        $meeting.Body = 'Test Meeting 1'
        $meeting.Location = 'Virtual'
        $meeting.ReminderSet = $true
        $meeting.Importance = 1
        $meeting.MeetingStatus = [Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeeting
        $meeting.Recipients.Add("$email")
        $meeting.ReminderMinutesBeforeStart = 15
        $meeting.Start = (Get-Date -Hour 8 -Minute 0 -Second 0).AddDays(1)
        $meeting.Duration = 30
        $meeting.Send()
        Write-Host "Done" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
        Continue
    }
}

Function Calendar-Forward-Meeting {

    try {
        Write-Host "Forwarding Test Meeting 1 : " -NoNewline
        $outlook = New-Object -ComObject Outlook.Application??

        $appointment = $outlook.Session.GetDefaultFolder(9).Items | Where-Object { $_.Subject -eq "Test Meeting 1" }??

        $mail = $appointment.ForwardAsVcal()

        $mail.Recipients.Add("$email")??

        $mail.Subject = "Forwarded: " + $appointment.Subject
        $mail.Body = "The following appointment has been forwarded to you: " + $appointment.Subject
        $mail.Send()??

        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
        Continue
    }
}

Function Calendar-Create-Meeting-Attachment {

    try {
        Write-Host "Creating Test Meeting 2 with attachment : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        #Create a new appointment item
        $appt = $ol.CreateItem(1)
        #Set the properties of the appointment item
        $appt.Subject = "Test Meeting 2 with attachment"
        $appt.Location = "Virtual"
        $appt.Start = (Get-Date -Hour 8 -Minute 0 -Second 0).AddDays(1)
        $appt.End = (Get-Date -Hour 8 -Minute 30 -Second 0).AddDays(1)
        $appt.RequiredAttendees = "$email"
        $appt.Body = "Test Meeting 2 with attachment"
        $appt.ReminderMinutesBeforeStart = 30

        <#$attachment = #>$appt.Attachments.Add("C:\Users\2003686\OneDrive - Standard Chartered Bank\Desktop\Outlook Test Case\file1.txt")

        #Save the appointment and send the meeting request
        $appt.Save()
        $appt.Send()
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
        Continue
    }
}

Function Calendar-Create-Meeting-DL {

    try {
        Write-Host "Sending Meeting to DL : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        $meeting = $ol.CreateItem('olAppointmentItem')
        $meeting.Subject = 'Test Meeting 2'
        $meeting.Body = 'Test Meeting 2+
        '
        $meeting.Location = 'Virtual'
        $meeting.ReminderSet = $true
        $meeting.Importance = 1
        $meeting.MeetingStatus = [Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeeting
        $meeting.Recipients.Add("$dl")
        $meeting.ReminderMinutesBeforeStart = 15
        $meeting.Start = (Get-Date -Hour 8 -Minute 0 -Second 0).AddDays(1)
        $meeting.Duration = 30
        $meeting.Send()
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
        Continue
    }
}

Function Calendar-Modify-Meeting {

    try {
        Write-Host "Modifying Meeting Body : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        $calendar = ($ol.GetNamespace("MAPI")).GetDefaultFolder(9)
        $meeting = $calendar.Items
        $meeting = $meeting | Where-Object { $_.Subject -eq "Test Meeting 1" }
        $meeting.Body = "Meeting has been modified"
        $meeting.Send()
        $meeting.Save()
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
        Continue
    }
}

Function Calendar-Delete-Meeting {

    try {
        Write-Host "Deleting Meeting With Subject Test Meeting 1 : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        $calendar = ($ol.GetNamespace("MAPI")).GetDefaultFolder(9)
        $meeting = $calendar.Items
        $meeting = $meeting | Where-Object { $_.Subject -eq "Test Meeting 1" }
        $meeting.MeetingStatus = [Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeetingCanceled
        $meeting.Delete()
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
        Continue
    }
}

Function Calendar-Create-Recurring-Meeting {

    try {
        Write-Host "Creating Recurring Meeting With Subject Test Meeting 1 : " -NoNewline
        $ol = New-Object -ComObject Outlook.Application
        $calendar = ($ol.GetNamespace("MAPI")).GetDefaultFolder(9)
        $meeting = $calendar.Items.Add(1)
        $meeting.Subject = "Test Recurring Meeting 1"
        $meeting.Start = (Get-Date -Hour 8 -Minute 0 -Second 0).AddDays(1)
        $meeting.Duration = 30
        $meeting.Location = "Virtual"
        $meeting.Body = "This is a Test Recurring Metting 1"
        $meeting_recurring = ($meeting.GetRecurrencePattern())
        $meeting_recurring.RecurrenceType = 2
        $meeting_recurring.Interval = 1
        $meeting_recurring.StartTime = (Get-Date -Hour 8 -Minute 0 -Second 0).AddDays(1)
        $meeting_recurring.Occurrences = 10
        $meeting.Save()
        Write-Host "Success" -ForegroundColor Green
    }

    catch {
        Write-Warning ($_)
        Continue
    }
}

Function Calendar-Check-Protocol {
    Write-Host "Checking calendar properties protocol : " -NoNewline
    $ol = New-Object -ComObject Outlook.Application
    $calendar = ($ol.GetNamespace("MAPI")).GetDefaultFolder(9)
    $calendar.WebViewURL
    if ($null -eq $calendar.WebViewURL) {
        Write-Host "MAPI" -ForegroundColor Green
    }

    elseif ($calendar.WebViewURL) {
        Write-Host "REST" -ForegroundColor Red
    }

}

Function Calendar-Check-Sharedimprovement {
    try {
        Write-Host "Checking calendar shared calendar improvements option : " -NoNewline
        $check = (Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\Calendar").ShowAsOutlookAB
        if ($null -eq $check) {
            Write-Host "Disabled" -ForegroundColor Green
        }

        if ($check -eq "0") {
            Write-Host "Disabled" -ForegroundColor Green
        }

        if ($check -eq "1") {
            Write-Host "Enabled" -ForegroundColor Red
        }
    }

    catch {
        Write-Warning ($_)
        Continue
    }
}
