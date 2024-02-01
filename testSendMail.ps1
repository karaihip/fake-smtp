# SMTPサーバーの詳細を設定します
$SmtpServer = "localhost"   #"smtp.example.com"
$SmtpPort = 8587

# メールの詳細を設定します
$EmailFrom = "sender@example.com"
$EmailTo = "receiver@example.com"
$EmailSubject = "Test email subject"
$EmailBody = "This is a test email."

# メールメッセージを作成します
$Message = New-Object System.Net.Mail.MailMessage($EmailFrom, $EmailTo, $EmailSubject, $EmailBody)
$Message.CC.Add("ccEmail@example2.com")

# SMTPクライアントを作成し、メールを送信します
$SmtpClient = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
$SmtpClient.Send($Message)

Write-Host "Email sent successfully."
