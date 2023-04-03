# HelloID-Task-SA-Target-ExchangeOnline-MailboxGrantSendAs

## Prerequisites
Before using this snippet, verify you've met with the following requirements:
- [ ] The PowerShell EXO v3 module must be installed on the server running the Agent. [See instructions](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps)

- [ ] User defined variables: `ExchangeOnlineAdminUsername` and `$ExchangeOnlineAdminPassword` created in your HelloID portal. [See also](https://docs.helloid.com/en/variables/custom-variables.html)


## Description

This code snippet executes the following tasks:

1. Define a hash table `$formObject`. The keys of the hash table represent the properties of the `Add-RecipientPermission` cmdlet, while the values represent the values entered in the form.

> To view an example of the form output, please refer to the JSON code pasted below.

```json
{
    "MailboxDistinguishedName": "JohnD01@myenvironment.onmicrosoft.com",
    "UsersToAdd": [
        {
            "id": "JaneD01"
        },
        {
            "id": "JohnD02@myenvironment"
        },
        {
            "id": "d55a09fe-b231-XxXx-XxXx-9c369364e991"
        }
    ]
}
```
> :exclamation: It is important to note that the names of your form fields might differ. Ensure that the `$formObject` hashtable is appropriately adjusted to match your form fields.
> The **ID** can hold different values [See the Microsoft Docs page](https://learn.microsoft.com/en-us/powershell/module/exchange/add-recipientpermission?view=exchange-ps#-trustee)

2. Constructs a powershell credential object from the supplied administrative username and password

3. Connects with the credentials to the Exchange online environment by means of the `Connect-ExchangeOnline` cmdlet

4. Calls the `Add-RecipientPermission` cmdlet to add Send-As permissions to a recipient

5. Disconnects from the Exchange environment by means of the `Disconnect-ExchangeOnline` cmdlet
