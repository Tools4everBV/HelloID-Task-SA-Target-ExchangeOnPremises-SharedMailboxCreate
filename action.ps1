# HelloID-Task-SA-Target-ExchangeOnPremises-SharedMailboxCreate
###############################################################
# Form mapping
$formObject = @{
    Name               = $form.CommonName
    Alias              = $form.Alias
    UserPrincipalName  = $form.UserPrincipalName
    OrganizationalUnit = $form.OrganizationalUnit
    Password           = (ConvertTo-SecureString -AsPlainText $form.Password -Force)
}

[bool]$IsConnected = $false
try {
    Write-Information "Executing ExchangeOnPremises action: [SharedMailboxCreate] for: [$($formObject.Name)]"
    $adminSecurePassword = ConvertTo-SecureString -String $ExchangeAdminPassword -AsPlainText -Force
    $adminCredential = [System.Management.Automation.PSCredential]::new($ExchangeAdminUsername, $adminSecurePassword)
    $sessionOption = New-PSSessionOption -SkipCACheck -SkipCNCheck
    $exchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ExchangeConnectionUri -Credential $adminCredential -SessionOption $sessionOption -Authentication Kerberos  -ErrorAction Stop
    $null = Import-PSSession $exchangeSession -DisableNameChecking -AllowClobber -CommandName 'New-Mailbox'
    $IsConnected = $true

    $mailbox = New-Mailbox @formObject -Shared -ErrorAction Stop

    $auditLog = @{
        Action            = 'CreateResource'
        System            = 'ExchangeOnPremises'
        TargetIdentifier  = $mailbox.ExchangeGuid.Guid
        TargetDisplayName = $formObject.name
        Message           = "ExchangeOnPremises action: [SharedMailboxCreate] for: [$($formObject.Name)] executed successfully"
        IsError           = $false
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Information "ExchangeOnPremises action: [SharedMailboxCreate] for: [$($formObject.Name)] executed successfully"
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'CreateResource'
        System            = 'ExchangeOnPremises'
        TargetIdentifier  = $formObject.name
        TargetDisplayName = $formObject.name
        Message           = "Could not execute ExchangeOnPremises action: [SharedMailboxCreate] for: [$($formObject.Name)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnPremises action: [SharedMailboxCreate] for: [$($formObject.Name)], error: $($ex.Exception.Message)"
} finally {
    if ($IsConnected) {
        Remove-PSSession -Session $exchangeSession -Confirm:$false  -ErrorAction Stop
    }
}
###############################################################
