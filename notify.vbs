'==== CONFIGURE EMAIL and MESSAGE SETTINGS ==================================================================

Const name      = "your name"
Const fromEmail	= "email@gmail.com"
Const password	= "password"
Const toEmail   = "email@gmail.com"
Const subject   = "[Message from your computer] Someone is using me"

Dim   textBody, messTitle, messBody

      textBody  = "YO! Waddup!" & Vbcrlf & "Your computer at work here. Some alien with unhygenic fingers is tippy-tapping on me." & Vbcrlf & "Sort it out."

      messTitle = "Warning: You can't handle the truth!"
      messBody  = "You are using " & name & "'s computer." & Vbcrlf & "A log will be made of your activity and sent to " & name & "'s email."

'==== END OF CONFIGURATION ==================================================================================

Dim emailObj, emailConfig
Set emailObj      = CreateObject("CDO.Message")
emailObj.From     = fromEmail
emailObj.To       = toEmail
emailObj.Subject  = subject
emailObj.TextBody = textBody

Set emailConfig = emailObj.Configuration
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver")       = "smtp.gmail.com"
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport")   = 465
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing")        = 2
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpusessl")       = true
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusername")     = fromEmail
emailConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendpassword")     = password
emailConfig.Fields.Update

emailObj.Send

Set emailObj	= nothing
Set emailConfig	= nothing

'==== DISPLAY MESSAGE =======================================================================================

msgbox messBody, 48, messTitle