# Export-MessageQueue.ps1

Export messages from a transport queue to file system for manual replay

## Description

This script suspends a transport queue, exports the messages to the local file system. After successful export the messages are optionally deleted from the queue.

This script utilizes the GlobalFunctions library [https://github.com/Apoc70/GlobalFunctions](https://github.com/Apoc70/GlobalFunctions)

## Parameters

### Queue

Full name of the transport queue, e.g. SERVER\354
Use Get-Queue -Server SERVERNAME to identify message queue

### Path

Path to folder for exported messages

### DeleteAfterExport

Switch to delete per Exchange Server subfolders and creating new folders

## Examples

``` PowerShell
.\Export-MessageQueue -Queue MCMEP01\45534 -Path D:\ExportedMessages
```

Export messages from queue MCMEP01\45534 to D:\ExportedMessages and do not delete messages after export

``` PowerShell
.\Export-MessageQueue -Queue MCMEP01\45534 -Path D:\ExportedMessages -DeleteAfterExport
```

Export messages from queue MCMEP01\45534 to D:\ExportedMessages and delete messages after export

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
