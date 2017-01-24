# Get-ExchangeOrganizationDetails.ps1
This script fetches Exchange organization configuration data and stores the data in plain text or CSV files.

##Description
The script gathers a lot of Exchange organizational configuration data. The data is stored in separate log files.

The log files are stored in a separate subfolder located under the script directory. An exisiting subfolder will be deleted automatically.

Optionally, the log files can automatically be zipped. The zipped archive can be sent by email as an attachment.    

##Inputs
#Prefix
Prefix to be used with log files and zip archive

#FolderName
Folder name of sub folder where all log files will be stored (default: ExchangeOrgInfo)

#Zip
Switch to optonally send all log files to a zip archive

#SendMail
Switch to send the zipped archive via email

#MailFrom
Sender email address

#MailTo
Recipient(s) email address(es)

#MailServer
FQDN of SMTP mail server to be used


##Outputs
All outputs are written as separate files to an automatically created subfolder (relative to the script folder).

When the Zip switch is used, the zipped archive will be created in the same folder where the script is located.

##Examples
```
.\Get-ExchangeOrganizationDetails.ps1 -Prefix MYCOMPANY
```
Gather all data using MYCOMPANY as a prefix

```
.\Get-ExchangeOrganizationDetails.ps1 -Prefix MYCOMPANY -Zip
```
Gather all data using MYCOMPANY as a prefix and save all files as a compressed archive

## Note
THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

##TechNet Gallery
Find the script at TechNet Gallery
* https://gallery.technet.microsoft.com/Script-to-remove-unwanted-9d119c6b


##Credits
Written by: Thomas Stensitzki

Stay connected:

* My Blog: http://justcantgetenough.granikos.eu
* Twitter:	https://twitter.com/stensitzki
* LinkedIn:	http://de.linkedin.com/in/thomasstensitzki
* Github:	https://github.com/Apoc70

For more Office 365, Cloud Security and Exchange Server stuff checkout services provided by Granikos

* Blog:     http://blog.granikos.eu/
* Website:	https://www.granikos.eu/en/
* Twitter:	https://twitter.com/granikos_de