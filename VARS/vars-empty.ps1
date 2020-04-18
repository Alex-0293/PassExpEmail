[string]$Global:BodyFile               = "" #"C:\DATA\Projects\PassExpEmails\template.html"
[string]$Global:LogoFile               = "" #"C:\DATA\Projects\PassExpEmails\logo.gif"
[string]$global:GlobalKey1             = "" #AESKey
[string]$global:APP_SCRIPT_ADMIN_Login = "" # Enc data
[string]$global:APP_SCRIPT_ADMIN_Pass  = "" # Enc data
[string]$Global:OpLog                  = "" #"C:\DATA\Projects\PassExpEmails\op.log"
[string]$Global:Subject                = "" #"Уведомление о скором истечении срока действия пароля пользователей"
[string]$Global:SmtpServer             = "" #"mail.example.com"
[string]$Global:From                   = "" #"user@example.com"
[string]$Global:DC                     = "" #"Server"
[string]$Global:OU                     = "" #"OU=DEPARTMENTS, OU=COMPANY, DC=COMPANY, DC=local"
[string]$Global:AdminEmail             = "" #"admin@example.com"
[array] $Global:Days                   = @(1,2,3,7)
[bool]  $global:UseMailAuth            = $false