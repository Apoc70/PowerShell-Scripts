# Export-MessageQueue.ps1

Export messages from a transport queue to file system for manual replay

Source: [Export-MessageQueue in Exchange4ITPros](https://github.com/Apoc70/Exchange4ITPros/tree/main/Exchange%20Server/Maintenance/Export-MessageQueue)

> **Maintenance note:** This script is now maintained in the [Exchange for IT Pros repository](https://github.com/Apoc70/Exchange4ITPros).

## Description

This script suspends messages in a transport queue and exports them to the local file system. After a successful export, the messages can optionally be deleted from the queue.

This script utilizes the [GlobalFunctions module](https://github.com/Apoc70/Exchange4ITPros/tree/main/Modules/GlobalFunctions).

## Parameters

### Queue

Full name of the transport queue, e.g. SERVER\354
Use Get-Queue -Server SERVERNAME to identify message queue

### Path

Path to folder for exported messages

### DeleteAfterExport

After a successful export, prompts to delete suspended messages from the queue without sending an NDR.

## Examples

``` PowerShell
.\Export-MessageQueue -Queue MCMEP01\45534 -Path D:\ExportedMessages
```

Export messages from queue MCMEP01\45534 to D:\ExportedMessages and do not delete messages after export

``` PowerShell
.\Export-MessageQueue -Queue MCMEP01\45534 -Path D:\ExportedMessages -DeleteAfterExport
```

Export messages from queue MCMEP01\45534 to D:\ExportedMessages and delete messages after export

## Disclaimer

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
