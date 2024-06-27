# Set-UserPicture

This script fetches images from a source folder, resizes images for use with Exchange, Active Directory and Intranet. Resized images are are written to Exchange and Active Directory.

## Description

The script parses all images provided in a dedicated source folder. The images are resized and cropped for use with the following targets:

* Exchange Server user photo
* Active Directory thumbnailPhoto attribute
* Local intranet use in your local infrastructure

Source images must be named using the respective user logon name.

Example: MYDOMAIN\JohnDoe --> JohnDoe.jpg

Preferably, the images are stored in jpg format.

Optionally, processed image files are moved from the respective folder (Exchange, AD, Intranet) to a processed folder.

## Requirements

* GlobalFunctions PowerShell module, described here: [https://scripts.granikos.eu](https://scripts.granikos.eu)
* ResizeImage.exe executable, described here: [https://granikos.eu/add-resized-user-photos-automatically/](https://granikos.eu/add-resized-user-photos-automatically/)
* Exchange Server 2016+ Management Shell (EMS) for storing user photos in on-premises mailboxes
* Exchange Online Management Shell for storing user photos in cloud mailboxes
* Write access to thumbnailPhoto attribute in Active Directory

## Parameters

### PictureSource

Absolute path to source images
Filenames must match the logon name of the user

### TargetPathAD

Absolute path to store images resized for Active Directory

### TargetPathExchange

Absolute path to store images resized for Exchange

### TargetPathIntranet

Absolute path to store images resized for Intranet

### ExchangeOnPrem

Switch to create resized images for Exchange On-Premesis and store images in users mailbox
Requires the image tool to be available in TargetPathExchange

### ExchangeOnline

Switch to create resized images for Exchange Online and store images in users mailbox
Requires the image tool to be available in TargetPathExchange

### ExchangeOnPremisesFqdn

Name the on-premises Exchange Server Fqdn to connect to using remote PowerShell

### ActiveDirectory

Switch to create resized images for Active Directory and store images in users thumbnailPhoto attribute
Requires the image tool to be available in TargetPathAD

### Intranet

Switch to create resized images for Intranet
Requires the image tool to be available in TargetPathIntranet

### SaveUserStatus

Switch to save a last modified status in a local Xml file. Currently in development.

### MoveAction

Optional action to move processed images to a dedicated sub folder.

Possible values:

* MoveTargetToProcessed = Move Exchange, AD or Intranet pictures to a subfolder
* MoveSourceToProcessed = Move image source to a subfolder

## Examples

``` PowerShell
.\Set-UserPicture.ps1 -ExchangeOnPrem
```

Resize photos stored in the default PictureSource folder for Exchange (648x648) and write images to user mailboxes

``` PowerShell
.\Set-UserPicture.ps1 -ExchangeOnline -PictureSource '\\SRV01\HRShare\Photos' -TargetPathExchange '\\SRV02\ExScripts\Photos'
```

Resize photos stored on a SRV01 share for Exchange Online and save resized photos on a SRV02 share

``` PowerShell
.\Set-UserPicture.ps1 -ActiveDirectory
```

Resize photos stored in the default PictureSource folder for Active Directory (96x96) and write images to user thumbnailPhoto attribute

``` PowerShell
.\Set-UserPicture.ps1 -Intranet
```

Resize photos stored in the default PictureSource folder for Intranet (150x150)

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki

## Stay connected

- My Blog: [https://blog.granikos.eu](https://blog.granikos.eu)
- Bluesky: [https://bsky.app/profile/stensitzki.bsky.social](https://bsky.app/profile/stensitzki.bsky.social)
- LinkedIn: [https://www.linkedin.com/in/thomasstensitzki](https://www.linkedin.com/in/thomasstensitzki)
- YouTube: [https://www.youtube.com/@ThomasStensitzki](https://www.youtube.com/@ThomasStensitzki)
- LinkTree: [https://linktr.ee/stensitzki](https://linktr.ee/stensitzki)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Website: [https://www.granikos.eu](https://www.granikos.eu)
- Bluesky: [https://bsky.app/profile/granikos.bsky.social](https://bsky.app/profile/granikos.bsky.social)