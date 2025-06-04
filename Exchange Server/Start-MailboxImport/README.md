# Start-MailboxImport.ps1

Import one or more PST files into an exisiting mailbox or a archive.

## Description

This script imports one or more PST files into a user mailbox or a user archive as batch.

PST file names can used as target folder names for import. PST files are renamed to support file name limitations by New-MailboxImportRequest cmdlet.

All files of a given folder will be imported into the user's mailbox.

## Requirements

- Windows Server 2016+
- Exchange Server 2016+
- GlobalFunctions PowerShell Module, [https://www.powershellgallery.com/packages/GlobalFunctions](https://www.powershellgallery.com/packages/GlobalFunctions)

## Parameters

### Identity

Mailbox identity in which the PST files get imported

### Archive

Import PST files into the online archive.

### FilePath

Folder which contains the PST files. Has to be an UNC path.

### FilenameAsTargetFolder

Import the PST files into dedicated target folders. The folder name will equal the file name.

### BadItemLimit

Standard is set to 0. Don't max it out because the script doesn't add "AcceptLargeDatalost".

### ContinueOnError

If set the script continue with the next PST file if a import request failed.

### SecondsToWait

Timespan to wait between import request staus checks in seconds. Default:

### IncludeFolders

If set the import would only import the given folder + subfolders. Note: If you want to import subfolders you have to use /* at the end of the folder. (Test/*).

### TargetFolder

Import the files in to definied target folder. Can't be used together with FilenameAsTargetFolder

### Recurse

If this parameter is set all PST files in subfolders will be also imported

### RenameFileAfterImport

Rename successfully imported PST files to simplify a re-run of the script. A .PST file will be renamed to .imported

## Examples

``` PowerShell
.\Start-MailboxImport.ps1 -Identity testuser -Filepath "\\testserver\share"
```

Import all PST files into the mailbox "testuser"

``` PowerShell
.\Start-MailboxImport.ps1 -Identity testuser -Filepath "\\testserver\share\*" -FilenameAsTargetFolder -SecondsToWait 90
```

Import all PST files into the mailbox "testuser". Use PST file name as target folder name. Wait 90 seconds between each status check

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki

Related blog post: [https://granikos.eu/simple-import-of-multiple-pst-files-for-a-single-user/](https://granikos.eu/simple-import-of-multiple-pst-files-for-a-single-user/)

### Stay connected

* Blog: [https://blog.granikos.eu](https://blog.granikos.eu)
* Bluesky: [https://bsky.app/profile/stensitzki.eu](https://bsky.app/profile/stensitzki.eu)
* LinkedIn: [https://www.linkedin.com/in/thomasstensitzki](https://www.linkedin.com/in/thomasstensitzki)
* YouTube: [https://www.youtube.com/@ThomasStensitzki](https://www.youtube.com/@ThomasStensitzki)
* LinkTree: [https://linktr.ee/stensitzki](https://linktr.ee/stensitzki)

For more Microsoft 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

* Website: [https://granikos.eu](https://granikos.eu)
* Bluesky: [https://bsky.app/profile/granikos.eu](https://bsky.app/profile/granikos.eu)
