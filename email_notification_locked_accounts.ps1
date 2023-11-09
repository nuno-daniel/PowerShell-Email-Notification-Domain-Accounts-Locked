<#
For use in a scheduled task on an Active Directory Domain Controller
Name: Lockout Email
Trigger: On event - Log: Security, Source: Microsoft-Windows-Security-Auditing, Event ID: 4740
#>

# Set script encoding to UTF-8
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

try {
    # Get information about the locked account and other information like the desktop name, events, and so on
    $AccountLockOutEvent = Get-EventLog -LogName "Security" -InstanceID 4740 -Newest 1
    $LockedAccount = $($AccountLockOutEvent.ReplacementStrings[0])
    $AccountLockedAt = $($AccountLockOutEvent.ReplacementStrings[1])
    $AccountLockOutEventTime = $AccountLockOutEvent.TimeGenerated
    $AccountLockOutEventMessage = $AccountLockOutEvent.Message

    # Get the user details from Active Directory
    $UserDetails = Get-AdUser -Identity $AccountLockOutEvent.ReplacementStrings[0] -Properties EmailAddress

    # Extract email address from user details
    $UserEmail = $UserDetails.EmailAddress

    # Define helpdesk / support email
    $HelpdeskEmail = "helpdesk@domain.com"

    # Define CcRecipient
    $CcRecipient = $HelpdeskEmail

    # Read HTML content from file
    $EmailBodyFilePath = "C:\scripts\locked_accounts\email_template.html"
    $EmailBodyTemplate = Get-Content -Path $EmailBodyFilePath -Raw
    # Replace variables in the HTML template
    $EmailBody = $ExecutionContext.InvokeCommand.ExpandString($EmailBodyTemplate)
	
    # Check if the email address is available
    if ($UserEmail -ne $null -and $UserEmail -ne "") {
        # Email parameters
        $To = $UserEmail
		
    } else {
        # If no user email is available, still send to helpdesk, but don't include user in CC
        $To = $HelpdeskEmail
    }

    # Create a MailMessage object
    $MailMessage = New-Object System.Net.Mail.MailMessage
    $MailMessage.From = "Access Control System <noreply@domain.com>"
    $MailMessage.To.Add($To)
    $MailMessage.CC.Add($CcRecipient)
    $MailMessage.Subject = "[NOTIFICATION] Domain account blocked: $($AccountLockOutEvent.ReplacementStrings[0])"
    $MailMessage.Body = $EmailBody
    $MailMessage.IsBodyHtml = $true

    # Retrieve the password securely
    $Username = "noreply@domain.com"
    $PasswordFile = "C:\scripts\locked_accounts\pwd.txt"
    $EncryptedPassword = Get-Content -Path $PasswordFile | ConvertTo-SecureString
    
    # Create the PSCredential object
    $Credential = New-Object System.Management.Automation.PSCredential($Username, $EncryptedPassword)

    # Create an SMTP client and send the email
    $SMTPClient = New-Object Net.Mail.SmtpClient("domain.com")
    $SMTPClient.Port = 587
    $SMTPClient.EnableSsl = $true
    $SMTPClient.Credentials = $Credential

    # Explicitly set the content type to UTF-8
    $ContentType = New-Object System.Net.Mime.ContentType
    $ContentType.CharSet = "UTF-8"
    $MailMessage.BodyEncoding = [System.Text.Encoding]::GetEncoding($ContentType.CharSet)
    
    # Send email
    $SMTPClient.Send($MailMessage)
    
    # Log successful email delivery
    $SuccessLogEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Email sent successfully to $($To)"
    Add-Content -Path "C:\scripts\locked_accounts\log\email_success_log.txt" -Value $SuccessLogEntry

} catch {
    $ErrorMessage = $_.Exception.Message
    Write-Host "Failed to send email. Error: $ErrorMessage"

    # Log the error
    $ErrorLogEntry = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - Failed to send email. Error: $ErrorMessage"
    Add-Content -Path "C:\scripts\locked_accounts\log\email_error_log.txt" -Value $ErrorLogEntry
}

#$EmailBody = $EmailBody | Out-File -Encoding UTF8 -FilePath "C:\scripts\locked_accounts\output.html"
$EmailBody | Out-File -Encoding UTF8 -FilePath "C:\scripts\locked_accounts\output.html"
