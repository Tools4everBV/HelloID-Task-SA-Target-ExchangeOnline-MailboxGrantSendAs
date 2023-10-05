# HelloID-Task-SA-Target-ExchangeOnline-MailboxGrantSendAs
##########################################################
# Form mapping
$formObject = @{
    MailboxIdentity = $form.MailboxIdentity
    UsersToAdd      = $form.UsersToAdd.id
}
[bool]$IsConnected = $false
try {
    Write-Information "Executing ExchangeOnline action: [MailboxGrantSendAs] for: [$($formObject.MailboxIdentity)]"

    $null = Import-Module ExchangeOnlineManagement

    $securePassword = ConvertTo-SecureString $ExchangeOnlineAdminPassword -AsPlainText -Force
    $credential = [System.Management.Automation.PSCredential]::new($ExchangeOnlineAdminUsername, $securePassword)
    $null = Connect-ExchangeOnline -Credential $credential -ShowBanner:$false -ShowProgress:$false -ErrorAction Stop -Verbose:$false -CommandName 'Add-RecipientPermission','Disconnect-ExchangeOnline'
    $IsConnected = $true

    foreach ($user in $formObject.UsersToAdd) {
        $recipient = Add-RecipientPermission -Identity $formObject.MailboxIdentity  -AccessRights SendAs -Confirm:$false -Trustee $user -ErrorAction Stop

        $auditLog = @{
            Action            = 'UpdateResource'
            System            = 'ExchangeOnline'
            TargetIdentifier  = $recipient.Identity
            TargetDisplayName = $formObject.MailboxIdentity
            Message           = "ExchangeOnline action: [MailboxGrantSendAs] Added [$($user)] to mailbox [$($formObject.MailboxIdentity)] executed successfully"
            IsError           = $false
        }
        Write-Information -Tags 'Audit' -MessageData $auditLog
        Write-Information "ExchangeOnline action: [MailboxGrantSendAs] Added [$($user)] to mailbox [$($formObject.MailboxIdentity)] executed successfully"
    }
} catch {
    $ex = $_
    $auditLog = @{
        Action            = 'UpdateResource'
        System            = 'ExchangeOnline'
        TargetIdentifier  = $formObject.MailboxIdentity
        TargetDisplayName = $formObject.MailboxIdentity
        Message           = "Could not execute ExchangeOnline action: [MailboxGrantSendAs] for: [$($formObject.MailboxIdentity)], error: $($ex.Exception.Message)"
        IsError           = $true
    }
    Write-Information -Tags 'Audit' -MessageData $auditLog
    Write-Error "Could not execute ExchangeOnline action: [MailboxGrantSendAs] for: [$($formObject.MailboxIdentity)], error: $($ex.Exception.Message)"
} finally {
    if ($IsConnected) {
        $null = Disconnect-ExchangeOnline -Confirm:$false -Verbose:$false
    }
}
##########################################################
