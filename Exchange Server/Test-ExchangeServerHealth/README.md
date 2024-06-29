# Test-ExchangeServerHealth.ps1



## Description


## Requirements

- Windows Server 2016+
- Exchange Server 2016+

## Parameters

## Examples

### Example 1

Checks all servers in the organization and outputs the results to the shell window.

``` PowerShell
.\Test-ExchangeServerHealth.ps1 -Server HO-EX2010-MB1
```

### Example 2

Checks the server HO-EX2010-MB1 and outputs the results to the shell window.

``` PowerShell
.\Test-ExchangeServerHealth.ps1 -Server HO-EX2010-MB1
```

### Example 3

Checks all servers in the organization, outputs the results to the shell window, a HTML report, and emails the HTML report to the address configured in the script.

``` PowerShell
.\Test-ExchangeServerHealth.ps1 -ReportMode -SendEmail
```

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Original script by Paul Cunningham: [https://github.com/cunninghamp/Test-ExchangeServerHealth.ps1/tree/master](https://github.com/cunninghamp/Test-ExchangeServerHealth.ps1/tree/master)

Updated for Exchange Server 2019 by: Thomas Stensitzki

Related blog post:

### Stay connected

- My Blog: [https://blog.granikos.eu](https://blog.granikos.eu)
- Bluesky: [https://bsky.app/profile/stensitzki.bsky.social](https://bsky.app/profile/stensitzki.bsky.social)
- LinkedIn: [https://www.linkedin.com/in/thomasstensitzki](https://www.linkedin.com/in/thomasstensitzki)
- YouTube: [https://www.youtube.com/@ThomasStensitzki](https://www.youtube.com/@ThomasStensitzki)
- LinkTree: [https://linktr.ee/stensitzki](https://linktr.ee/stensitzki)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Website: [https://granikos.eu](https://www.granikos)
- Bluesky: [https://bsky.app/profile/granikos.bsky.social](https://bsky.app/profile/granikos.bsky.social)