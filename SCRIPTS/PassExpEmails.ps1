<#
    .SYNOPSIS 
        Alexk
        12.01.2020
        1
    .DESCRIPTION
    .PARAMETER
    .EXAMPLE
#>
$ImportResult = Import-Module AlexkUtils  -PassThru
if ($null -eq $ImportResult) {
    Write-Host "Module 'AlexkUtils' does not loaded!"
    exit 1
}
else {
    $ImportResult = $null
}
#requires -version 3

#########################################################################
function Get-WorkDir () {
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

[string]$MyScriptRoot       = Get-WorkDir
[string]$Global:ProjectRoot = Split-Path $MyScriptRoot -parent

Get-VarsFromFile    "$ProjectRoot\VARS\Vars.ps1"
Initialize-Logging   $ProjectRoot "Latest"



$Login          = Get-VarToString(Get-VarFromAESFile $global:GlobalKey1 $global:APP_SCRIPT_ADMIN_Login)
$Pass           = Get-VarFromAESFile $global:GlobalKey1 $global:APP_SCRIPT_ADMIN_Pass
$UserCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Login, $Pass

$Session = New-PSSession -ComputerName $Global:dc -Authentication Kerberos -Credential $UserCredential

$OU = $global:OU
[array]$output = @()

$output = Invoke-Command -Session $Session -ScriptBlock {`
    Import-Module ActiveDirectory
    $ADAccounts = Get-ADUser -filter * -searchbase  $Using:OU -properties PasswordExpired, PasswordNeverExpires, PasswordLastSet, Mail, Enabled | Where-object { $_.Enabled -eq $true -and $_.PasswordNeverExpires -eq $false }
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
        [string] $samAccountName    = $ADAccount.samAccountName
        [string] $userEmailAddress  = $ADAccount.mail
        [string] $userPrincipalName = $ADAccount.UserPrincipalName

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
            #$DaysToExpireDD = @($DaysToExpire.ToString() -Split ("\S{17}$"))
            $expiryDate     = $expiryDate.ToString("d", $ci)
            $res            = [pscustomobject]@{
                Sam             = $samAccountName
                Email           = $userEmailAddress
                UPN             = $userPrincipalName
                PasswordExpired = $ADAccount.PasswordExpired
                ExpiredOn       = [int]$DaysToExpire.Days
                Enabled         = $ADAccount.Enabled
            }
            $output += $res
        }
    }
    return $output
}
write-debug "$($output | Sort-Object ExpiredOn | Format-Table -AutoSize)"
[array]$Data = $output | Where-Object { ($Global:Days -contains $_.ExpiredOn) -and ($_.PasswordExpired -ne $True) } | Sort-Object ExpiredOn 

write-debug "$($Data | Format-Table -AutoSize)"

#$Data = $output | Where-Object { ($_.sam -eq "Admin1") } 

foreach ($Item in $Data) {
    [string] $Body    = Get-Content  $Global:BodyFile
    [string] $NewBody = $Body.Replace("<!--Data-->", "$($Item.ExpiredOn)")
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
    }
    if ($global:UseMailAuth) {
        $params.Add("User", (Get-VarFromAESFile $Global:GlobalKey1 $Global:MailUser))
        $params.Add("Pass", (Get-VarFromAESFile $Global:GlobalKey1 $Global:MailPass))
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
    }
    if ($global:UseMailAuth) {
        $params.Add("User", (Get-VarFromAESFile $Global:GlobalKey1 $Global:MailUser))
        $params.Add("Pass", (Get-VarFromAESFile $Global:GlobalKey1 $Global:MailPass))
    }

    Send-Email @params

    $Message = "Send Email to administrator $($Global:AdminEmail)"
    Add-ToLog   $Message $Global:OpLog 
    Write-Debug $Message
}