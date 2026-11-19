# Copy-ReceiveConnector.ps1

Copy a selected receive connector and its configuration and permissions to other Exchange Servers

## Description

This script copies a receive connector from a source Exchange Server to a single target Exchange server or to all Exchange servers.

Configured permissions are copied as well, if required.

All values read from the source connector are converted explicitly before they are handed over to New-ReceiveConnector, Set-ReceiveConnector, and Add-ADPermission. This is required because the Exchange Management Shell returns deserialized objects, and with PowerShell serialization payload signing enabled, Exchange no longer converts deserialized Exchange types like IPRange automatically (see issue #7).

## Requirements

- Windows Server 2016, 2019, 2022, or 2025
- Exchange Server 2016, Exchange Server 2019, or Exchange Server Subscription Edition (SE)
- Exchange Management Shell

## Parameters

### SourceServer

Source Exchange server hosting the receive connector to copy

### ConnectorName

Name of the receive connector to copy

### TargetServer

Target Exchange server to copy the selected receive connector to

### DomainController

Domain Controller name. All read and write operations use this domain controller.

### CopyToAllOther

Switch to copy to all other Exchange servers with the Mailbox role. Edge Transport servers are excluded.

### CopyPermissions

Copy non inherited source receive AD permissions to target receive connector. Inherited permissions will not be copied. Permissions that already exist on the target connector are skipped.

### MoveToFrontend

Change source connector transport role to FrontendTransport. This is required when you copy a receive connector from Exchange 2010 to Exchange 2013+

### ResetBindings

Do not copy bindings but reset receive connector network bindings to 0.0.0.0:25. The script warns if the source connector uses server specific bindings.

### UpdateExistingConnector

Update an existing receive connector without confirmation prompt.

### ViewEntireForest

View entire Active Directory forest

### SecondsToWait

Seconds to wait for Active Directory before AD permissions or permission groups are updated on a newly created connector. Default: 60

### WhatIf

Show what would happen without creating or changing anything.

## Examples

``` PowerShell
.\Copy-ReceiveConnector.ps1 -SourceServer MBX01 -ConnectorName nikos-one-RC2 -TargetServer MBX2 -DomainController MYDC1.mcsmemail.de
```

Copy receive connector nikos-one-RC2 from server MBX01 to server MBX2

``` PowerShell
.\Copy-ReceiveConnector.ps1 -SourceServer MBX01 -ConnectorName nikos-one-RC1 -CopyToAllOther -DomainController MYDC1.mcsmemail.de
```

Copy receive connector nikos-one-RC1 from server MBX01 to all other Exchange servers

``` PowerShell
.\Copy-ReceiveConnector.ps1 -SourceServer MBX2010 -ConnectorName "varunagroup relay" -TargetServer MBX01 -MoveToFrontend -ResetBindings -DomainController MYDC1.mcsmemail.de
```

Copy receive connector "varunagroup relay" from Exchange 2010 server MBX2010 to Exchange Server MBX01, make it a FrontEnd-Connector, and reset network binding

``` PowerShell
.\Copy-ReceiveConnector.ps1 -SourceServer MBX01 -ConnectorName MYRECEIVECONNECTOR -CopyToAllOther -DomainController MYDC1.mcsmemail.de -CopyPermissions -UpdateExistingConnector
```

Copy receive connector MYRECEIVECONNECTOR including permissions from server MBX01 to all other Exchange servers without confirmation prompt if connectors already exist

``` PowerShell
.\Copy-ReceiveConnector.ps1 -SourceServer MBX01 -ConnectorName MYRECEIVECONNECTOR -TargetServer MBX2 -DomainController MYDC1.mcsmemail.de -CopyPermissions -WhatIf
```

Show what would happen without changing anything

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki

### Stay connected

* Blog: [https://blog.granikos.eu](https://blog.granikos.eu)
* Bluesky: [https://bsky.app/profile/stensitzki.eu](https://bsky.app/profile/stensitzki.eu)
* LinkedIn: [https://www.linkedin.com/in/thomasstensitzki](https://www.linkedin.com/in/thomasstensitzki)
* YouTube: [https://www.youtube.com/@ThomasStensitzki](https://www.youtube.com/@ThomasStensitzki)
* LinkTree: [https://linktr.ee/stensitzki](https://linktr.ee/stensitzki)

For more Microsoft 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

* Website: [https://granikos.eu](https://granikos.eu)
* Bluesky: [https://bsky.app/profile/granikos.eu](https://bsky.app/profile/granikos.eu)