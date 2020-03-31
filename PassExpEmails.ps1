<#
    Name:       Список пользователей домена с указанием IP при удаленном доступе, праве удаленного доступа и состояния пароля
    Ver:           1.0
    Date:         25.10.2017
    Platform:  Windows server 2012
    PSVer:       4.0
    Author:     AlexK
#>
Import-Module AlexkUtils
#requires -version 3

#########################################################################
function Get-Workdir () {
    if ($PSScriptRoot -eq "") {
        if ($PWD -ne "") {
            $MyScriptRoot = $PWD
        }        
        else {
            Write-Host "Where i am? What is my work dir?"
        }
    }
    else {
        $MyScriptRoot = $PSScriptRoot
    }
    return $MyScriptRoot
}
# Error trap
trap {
    Get-ErrorReporting $_    
    exit 1
}
#########################################################################
Clear-Host

[string]$MyScriptRoot = Get-Workdir

Get-Vars    "$MyScriptRoot\Vars.ps1"
InitLogging $MyScriptRoot "Latest"



$Login          = Get-VarFromFile $global:GlobalKey1 $global:APP_SCRIPT_ADMIN_Login
$Pass           = ConvertTo-SecureString -String (Get-VarFromFile $global:GlobalKey1 $global:APP_SCRIPT_ADMIN_Pass) -AsPlainText -Force
$UserCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Login, $Pass

$Session = New-PSSession -ComputerName $Global:dc -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -AllowClobber -Module ActiveDirectory

$ADAccounts = Get-ADUser -filter * -searchbase  $Global:OU -properties PasswordExpired, PasswordNeverExpires, PasswordLastSet, Mail, Enabled | Where-object { $_.Enabled -eq $true -and $_.PasswordNeverExpires -eq $false }
[array]$output = @()


Foreach ($ADAccount in $ADAccounts) {
    $accountFGPP = Get-ADUserResultantPasswordPolicy $ADAccount  -ErrorAction SilentlyContinue 
    if ($null -ne $accountFGPP) {
        $maxPasswordAgeTimeSpan = $accountFGPP.MaxPasswordAge
    }
    else {
        $maxPasswordAgeTimeSpan = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
    }

    #Fill in the user variables
    $samAccountName    = $ADAccount.samAccountName
    $userEmailAddress  = $ADAccount.mail
    $userPrincipalName = $ADAccount.UserPrincipalName

    if ($ADAccount.PasswordExpired) {
        #Write-host "The password for account $samAccountName has expired!"
        $res = [pscustomobject]@{
            Sam             = $samAccountName
            Email           = $userEmailAddress
            UPN             = $userPrincipalName
            PasswordExpired = $ADAccount.PasswordExpired
            ExpiredOn       = 0 
            Enabled         = $ADAccount.Enabled
        }
        $output += $res
    }
    else {
        $ExpiryDate     = $ADAccount.PasswordLastSet + $maxPasswordAgeTimeSpan
        $TodaysDate     = Get-Date
        $DaysToExpire   = $ExpiryDate - $TodaysDate
        $DaysToExpireDD = @($DaysToExpire.ToString() -Split ("\S{17}$"))
        $expiryDate     = $expiryDate.ToString("d", $ci)
        $res            = [pscustomobject]@{
            Sam             = $samAccountName
            Email           = $userEmailAddress
            UPN             = $userPrincipalName
            PasswordExpired = $ADAccount.PasswordExpired
            ExpiredOn       = [int]$DaysToExpireDD[0] 
            Enabled         = $ADAccount.Enabled
        }
        $output += $res
    }
}
write-debug "$($output | Sort-Object ExpiredOn | Format-Table -AutoSize)"
$Data = $output | Where-Object { ($Global:Days -contains $_.ExpiredOn) -and ($_.PasswordExpired -ne $True) } | Sort-Object ExpiredOn 

write-debug "$($Data | Format-Table -AutoSize)"

#$Data = $output | Where-Object {($_.sam -eq "alex") } 

foreach ($Item in $Data) {
    $Body    = Get-Content  $Global:BodyFile
    $NewBody = $Body.Replace("<!--Data-->", "$($Item.ExpiredOn)")
    #$NewBody | Out-File "$MyScriptRoot/email.html"

    $params = @{
        SmtpServer          = $Global:SmtpServer
        Subject             = $Global:Subject
        Body                = $NewBody
        HtmlBody            = $True
        From                = $Global:From
        To                  = $Item.Email
        SSL                 = $true
        Attachment          = $Global:LogoFile
        AttachmentContentId = "Logo"
        Cntr                = 0
    }
    if ($global:UseMailAuth) {
        $params.Add("User", (Get-VarFromFile $Global:GlobalKey1 $Global:MailUser))
        $params.Add("Pass", (Get-VarFromFile $Global:GlobalKey1 $Global:MailPass))
    }

    Send-Email @params

    $Message = "Send Email to $($Item.Email)"
    Add-ToLog $Message $Global:OpLog 
    Write-Debug $Message
}

#Send email to administrator
if(@($Data).count -gt 0){
    $Body = $Data | Select-Object  sam, PasswordExpired, ExpiredOn, Enabled | Format-Table -AutoSize | out-String    
    
    $params = @{
        SmtpServer          = $Global:SmtpServer
        Subject             = $Global:Subject
        Body                = $Body
        HtmlBody            = $False 
        From                = $Global:From
        To                  = $Global:AdminEmail
        SSL                 = $true
        Attachment          = ""
        AttachmentContentId = ""
        Cntr                = 0
    }
    if ($global:UseMailAuth) {
        $params.Add("User", (Get-VarFromFile $Global:GlobalKey1 $Global:MailUser))
        $params.Add("Pass", (Get-VarFromFile $Global:GlobalKey1 $Global:MailPass))
    }

    Send-Email @params

    $Message = "Send Email to administrator $($Global:AdminEmail)"
    Add-ToLog   $Message $Global:OpLog 
    Write-Debug $Message
}