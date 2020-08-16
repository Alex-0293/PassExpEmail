<#
    .SYNOPSIS 
        .AUTOR
        .DATE
        .VER
    .DESCRIPTION
    .PARAMETER
    .EXAMPLE
#>
Param (
    [Parameter( Mandatory = $false, Position = 0, HelpMessage = "Initialize global settings." )]
    [bool] $InitGlobal = $true,
    [Parameter( Mandatory = $false, Position = 1, HelpMessage = "Initialize local settings." )]
    [bool] $InitLocal = $true   
)

$Global:ScriptInvocation = $MyInvocation
if ($env:AlexKFrameworkInitScript){. "$env:AlexKFrameworkInitScript" -MyScriptRoot (Split-Path $PSCommandPath -Parent) -InitGlobal $InitGlobal -InitLocal $InitLocal} Else {Write-host "Environmental variable [AlexKFrameworkInitScript] does not exist!" -ForegroundColor Red; exit 1}
# Error trap
trap {
    if (get-module -FullyQualifiedName AlexkUtils) {
       Get-ErrorReporting $_

        . "$GlobalSettingsPath\$SCRIPTSFolder\Finish.ps1"  
    }
    Else {
        Write-Host "[$($MyInvocation.MyCommand.path)] There is error before logging initialized. Error: $_" -ForegroundColor Red
    }   
    exit 1
}
################################# Script start here #################################
clear-host

$Login          = Get-VarToString(Get-VarFromAESFile $global:GlobalKey1 $Global:APP_SCRIPT_ADMIN_LoginFilePath)
$Pass           = Get-VarFromAESFile $global:GlobalKey1 $Global:APP_SCRIPT_ADMIN_PassFilePath
$UserCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Login, $Pass

$Session = New-PSSession -ComputerName $Global:dc -Authentication Kerberos -Credential $UserCredential

$OU = $global:OU
[array]$output = @()

$output = Invoke-Command -Session $Session -ScriptBlock {
    Import-Module ActiveDirectory
    $ADAccounts = Get-ADUser -filter * -SearchBase  $Using:OU -properties PasswordExpired, PasswordNeverExpires, PasswordLastSet, Mail, Enabled | Where-object { $_.Enabled -eq $true -and $_.PasswordNeverExpires -eq $false }
    [array]$output = @()

    Foreach ($ADAccount in $ADAccounts) {
        $accountFGPP = Get-ADUserResultantPasswordPolicy $ADAccount  -ErrorAction SilentlyContinue 
        if ($null -ne $accountFGPP) {
            $maxPasswordAgeTimeSpan = $AccountFGPP.MaxPasswordAge
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
            $res = [PSCustomObject]@{
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
    if ($Item.Email) {
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
            $params.Add("User", (Get-VarFromAESFile $Global:GlobalKey1 $Global:MailUserFile))
            $params.Add("Pass", (Get-VarFromAESFile $Global:GlobalKey1 $Global:MailPassFile))
        }

        Send-Email @params -Verbose
        #$params | format-table -AutoSize

        Add-ToLog -Message "Send Email to [$($Item.Email)]." -logFilePath $ScriptLogFilePath -display -status "Info" -level ($ParentLevel + 1)
    }
    Else {
        Add-ToLog -Message "Email for user [$($Item.Sam)] is incorrect." -logFilePath $ScriptLogFilePath -display -status "Error" -level ($ParentLevel + 1)
    }
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
        $params.Add("User", (Get-VarFromAESFile $Global:GlobalKey1 $Global:MailUserFile))
        $params.Add("Pass", (Get-VarFromAESFile $Global:GlobalKey1 $Global:MailPassFile))
    }

    Send-Email @params
    #$params | format-table -AutoSize
    
    Add-ToLog -Message "Send Email to administrator [$($Global:AdminEmail)]." -logFilePath $ScriptLogFilePath -display -status "Info" -level ($ParentLevel + 1)
}

################################# Script end here ###################################
. "$GlobalSettingsPath\$SCRIPTSFolder\Finish.ps1"