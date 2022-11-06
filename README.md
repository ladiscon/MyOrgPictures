# MyOrgPictures
A tool for generating a slide with pictures of everyone who reports to me

## Pre-requisites:

### Office 365 desktop apps

You need to have Office 365 desktop apps installed in the PC so that script can interact with PowerPoint APIs.

### PowerShell module for AzureAD

The script will use commands in the AzureAD module to retrieve people information and the reports for the given manager.
In order to install a module you must PowerShell in administrator mode.

Start -> type PowerShell -> select "Run as administrator"

Type the following command and click yes to confirm 2 prompts:
```
Install-Module AzureAD
```

### PowerShell module for Microsoft Graph

The script will use Microsoft Graph to retrieve the user picture. Note that AzureAD module can also retrieve a user picture, but only that stored in AzureAD which is a lower resolution version of the user's Microsoft 365 profile. Only thry Microsoft Graph is possible to retrieve the user's high resolution picture.

Type the following command and click yes to confirm:
```
Install-Module Microsoft.Graph
```

## Running the script:

1. Start PowerShell:

Start -> type PowerShell -> select "Run"

2. Connect to Azure AD:

```
Connect-AzureAD -AccountId $myLoginAccount
```
Where **$myLoginAccount** is your the login account you use to login to Azure AD. For example:

```
$myLoginAccount = "ladiscon@microsoft.com"
```

3. Run the script:

```
.\Get-MyOrgPictures.ps1 -Includes $listOfPeople
```
Where **$listOfPeople** is a quoted, comma-separated list of aliases or email addresses of the manager and other people to retrieve the picture for. For example:

```
$listOfPeople = "ladiscon","nirobson","tydakuja"
```

This will enumerate all direct reports recursively, download all pictures, and start PowerPoint and create slide with those pictures and names.

## Examples:

In this example I have 3 aliases in the -Includes parameter, one is my own alias, and the two other are embedded engineers that work in my org.

```
.\Get-MyOrgPictures.ps1 -Includes "ladiscon","nirobson","tydakuja"
```
