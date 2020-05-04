# Rename this file to Settings.ps1
#### Script params
    [string]$Global:LogoFile               = ""          # Attachment logo file path.
    [string]$global:APP_SCRIPT_ADMIN_Login = ""          # AES Login value path.
    [string]$global:APP_SCRIPT_ADMIN_Pass  = ""          # AES Password value path.
    [string]$Global:SmtpServer             = ""          # SMTP server FQDN.
    [string]$Global:From                   = ""          # From email.
    [string]$Global:DC                     = ""          # Domain controller name.
    [string]$Global:AdminEmail             = ""          # Admin email.

 ######################### no replacement ########################   

    [string]$Global:BodyFile               = "$ProjectRoot\$DATAFolder\template.html"                                     # Body template file path.
    [string]$Global:Subject                = "Уведомление о скором истечении срока действия пароля пользователей"  # Email subject.
    [array] $Global:Days                   = @(1,2,3,7)                                                            # Days remain to send email to user.
    [bool]  $global:UseMailAuth            = $false                                                                # Use SMTP authorization.
    
[bool] $Global:LocalSettingsSuccessfullyLoaded = $true

# Error trap
trap {
   $Global:LocalSettingsSuccessfullyLoaded = $False
   exit 1
}
