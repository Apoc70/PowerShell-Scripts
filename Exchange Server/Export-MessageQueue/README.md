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

## Stay connected

- My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)
- MVP Blog: [https://blogs.msmvps.com/thomastechtalk/](https://blogs.msmvps.com/thomastechtalk/)
- Tech Talk YouTube Channel (DE): [http://techtalk.granikos.eu](http://techtalk.granikos.eu)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)