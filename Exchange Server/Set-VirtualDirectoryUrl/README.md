# Set-VirtualDirectoryUrl

## SYNOPSIS

Configure Exchange Server 2013 Virtual Directory Url Settings

Thomas Stensitzki

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

Version 2.1, 2021-11-23

Use GitHub for Please send ideas, comments and suggestions.

## LINK

[https://scripts.granikos.eu](https://scripts.granikos.eu)

## DESCRIPTION

Exchange Server virtual directories (vDirs) require a proper configuration of
internal and external Urls. This is even more important in a co-existence
scenario with legacy Exchange Server versions.

Read more about Exchange Server 2013+ vDirs [here](https://techcommunity.microsoft.com/t5/exchange-team-blog/configuring-multiple-owa-ecp-virtual-directories-on-the-exchange/ba-p/611217?WT.mc_id=M365-MVP-5003086).


## NOTES

Requirements
- Windows Server 2016+
- Exchange Server 2016+

Revision History
--------------------------------------------------------------------------------
1.0     Initial community release
2.0     Updated for Exchange Server 2016, 2019, vNEXT
2.1     PowerShell Hygiene

## PARAMETERS

### PARAMETER InternalUrl

The internal url FQDN with leading protocol definition, ie. https://mobile.mcsmemail.de

### PARAMETER ExternalUrl

The internal url FQDN with leading protocol definition, ie. https://mobile.mcsmemail.de

## EXAMPLE

Configure internal and external url for different host headers
.\Set-VirtualDirectoryUrl -InternalUrl https://internal.mcsmemail.de -ExternalUrl https://mobile.mcsmemail.de

# Stay connected

* Blog: [https://blog.granikos.eu](https://blog.granikos.eu)
* Bluesky: [https://bsky.app/profile/stensitzki.eu](https://bsky.app/profile/stensitzki.eu)
* LinkedIn: [https://www.linkedin.com/in/thomasstensitzki](https://www.linkedin.com/in/thomasstensitzki)
* YouTube: [https://www.youtube.com/@ThomasStensitzki](https://www.youtube.com/@ThomasStensitzki)
* LinkTree: [https://linktr.ee/stensitzki](https://linktr.ee/stensitzki)

For more Microsoft 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

* Website: [https://granikos.eu](https://granikos.eu)
* Bluesky: [https://bsky.app/profile/granikos.eu](https://bsky.app/profile/granikos.eu)
