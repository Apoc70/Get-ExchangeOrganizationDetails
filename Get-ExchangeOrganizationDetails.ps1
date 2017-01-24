<# 

.SYNOPSIS 
This script fetches Exchange organization configuration data and stores the data in plain text or CSV files.

Thomas Stensitzki 

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE  
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER. 

Version 1.1, 2017-01-24

Please send ideas, comments and suggestions to support@granikos.eu 

.LINK 
More information can be found at http://scripts.granikos.eu

.DESCRIPTION 
The script gathers a lot of Exchange organizational configuration data. The data is stored in separate log files.

The log files are stored in a separate subfolder located under the script directory. An exisiting subfolder will be deleted automatically.

Optionally, the log files can automatically be zipped. The zipped archive can be sent by email as an attachment.    

     
.NOTES 
Requirements 
- Windows Server 2012 R2  
- .NET 4.5
- Exchange Server Management Shell
    
Revision History 
-------------------------------------------------------------------------------- 
1.0      Initial community release 
1.1      Some PowerShell hygiene

.PARAMETER Prefix
Prefix to be used with log files and zip archive

.PARAMETER FolderName
Folder name of sub folder where all log files will be stored (default: ExchangeOrgInfo)

.PARAMETER Zip
Switch to optionally send all log files to a zip archive

.PARAMETER SendMail
Switch to send the zipped archive via email

.PARAMETER MailFrom
Sender email address

.PARAMETER MailTo
Recipient(s) email address(es)

.PARAMETER MailServer
FQDN of SMTP mail server to be used

.EXAMPLE 
Gather all data using MYCOMPANY as a prefix
    
.\Get-ExchangeOrganizationDetails.ps1 -Prefix MYCOMPANY

.EXAMPLE
Gather all data using MYCOMPANY as a prefix and save all files as a compressed archive
    
.\Get-ExchangeOrganizationDetails.ps1 -Prefix MYCOMPANY -Zip

#>

param(
        [string] $Prefix = '',
        [string] $FolderName = 'ExchangeOrgInfo',
        [switch] $Zip,
        [switch] $SendMail,
        [string] $MailFrom = '',
        [string] $MailTo = '',
        [string] $MailServer = ''
)

# Enable during development only
# Set-StrictMode -Version Latest 

<#
    Hash table defining Exchange information sources

    Primary hash key defines the Exchange data source of information. Fetching data using Get-Mailbox, requires "Mailbox" as the primary key.
    The primary hash value is a secondary hash table defining the overall behavior for gathering Exchange configuration data.

    As a default each log files contains an overview section (FT) first, followed by a detailed section (FL). This does
    not make sense for all data gathered from the Exchange Organization.

    Output = [OverviewDetailed|DetailedOnly|SendConnector|VirtualDirectory|DistributionGroup]

        OverviewDetailed = default
        DetailedOnly = Only detailed output (FL)
        SendConnector = export as CSV while expanding the AddressSpaces property
        VirtualDirectory = Special requirement to filter certain vDirs that are only available on Exchange 2013+ servers
        DistributionGroup = At least exporting default distribution lists and room lists

    Command = Optional string value
    
        Command attributes for an Exchange cmdlet. E.g. "-IncludePreExchange2013" to gather legacy mailbox databases as well

    Unlimited = [true|false]

        Whether -ResultSize Unlimited should be used

    Sort = Optional string value

        Object attribute name used for sorting, e.g. "Name" or "Identity"
#>
$infoSources = [ordered]@{
    "MailboxDatabase" = @{Output = "OverviewDetailed"; 
                        Command = "-IncludePreExchange2013";
                        Unlimited = "false";
                        Sort = "Name"}
    "DatabaseAvailabilityGroup" = @{Output = "OverviewDetailed";Command = "-Status";Unlimited = "false";Sort = "Name"}
    "OrganizationConfig" = @{Output = "DetailedOnly";Command = "";Unlimited = "false";Sort = ""}
    "TransportConfig" = @{Output = "DetailedOnly";Command = "";Unlimited = "false";Sort = ""}
    "ReceiveConnector" = @{Output = "OverviewDetailed";Command = "";Unlimited = "false";Sort = "Identity"}
    "SendConnector" = @{Output = "SendConnector";Command = "";Unlimited = "false";Sort = "Identity"}
    "RemoteDomain" = @{Output = "OverviewDetailed";Command = "";Unlimited = "false";Sort = "Name"}
    "AcceptedDomain" = @{Output = "OverviewDetailed";Command = "";Unlimited = "false";Sort = "Name"}
    "EmailAddressPolicy" = @{Output = "OverviewDetailed";Command = "";Unlimited = "false";Sort = "Name"}
    "AddressList" = @{Output = "OverviewDetailed";Command = "";Unlimited = "false";Sort = "Name"}
    "TransportRule" = @{Output = "OverviewDetailed";Command = "";Unlimited = "false";Sort = "Name"}
    "SharingPolicy" = @{Output = "OverviewDetailed";Command = "";Unlimited = "false";Sort = "Name"}
    "RetentionPolicy" = @{Output = "OverviewDetailed";Command = "";Unlimited = "false";Sort = "Name"}
    "RetentionPolicyTag" = @{Output = "OverviewDetailed";Command = "";Unlimited = "false";Sort = "Name"}
    "DlpPolicy" = @{Output = "OverviewDetailed";Command = "";Unlimited = "false";Sort = "Name"}
    "SiteMailboxProvisioningPolicy" = @{Output = "OverviewDetailed";Command = "";Unlimited = "false";Sort = "Name"}
    "OwaMailboxPolicy" = @{Output = "OverviewDetailed";Command = "";Unlimited = "false";Sort = "Name"}

    # recipients
    "DistributionGroup" = @{Output = "DistributionGroup";Command = "";Unlimited = "true";Sort = "Name"}
    "Mailbox" = @{Output = "OverviewDetailed";Command = "";Unlimited = "true";Sort = "Name"}
    "RemoteMailbox" = @{Output = "OverviewDetailed";Command = "";Unlimited = "true";Sort = "Name"}
    "MailContact" = @{Output = "OverviewDetailed";Command = "";Unlimited = "true";Sort = "Name"}
    "UMMailbox" = @{Output = "OverviewDetailed";Command = "";Unlimited = "true";Sort = "Name"}

    # virtual directories
    "MapiVirtualDirectory" = @{Output = "VirtualDirectory";Command = '?{$_.AdminDisplayVersion -ilike ''*15''}';Unlimited = "false";Sort = "Name"}
    "AutodiscoverVirtualDirectory" = @{Output = "VirtualDirectory";Command = "";Unlimited = "false";Sort = "Name"}
    "EcpVirtualDirectory" = @{Output = "VirtualDirectory";Command = "";Unlimited = "false";Sort = "Name"}
    "WebServicesVirtualDirectory" = @{Output = "VirtualDirectory";Command = "";Unlimited = "false";Sort = "Name"}
    "ActiveSyncVirtualDirectory" = @{Output = "VirtualDirectory";Command = "";Unlimited = "false";Sort = "Name"}
    "OabVirtualDirectory" = @{Output = "VirtualDirectory";Command = "";Unlimited = "false";Sort = "Name"}
    "OwaVirtualDirectory" = @{Output = "VirtualDirectory";Command = "";Unlimited = "false";Sort = "Name"}
    "PowerShellVirtualDirectory" = @{Output = "VirtualDirectory";Command = "";Unlimited = "false";Sort = "Name"}
    }

# Declare some variables
$ScriptDir = Split-Path $script:MyInvocation.MyCommand.Path
$date = Get-Date -Format 'yyyyMMdd'
$FormatEnumerationLimit = -1
$global:TargetFolder = ''
$ZipFileName = "ExchangeOrganizationInfos-$($date).zip"

Set-ADServerSettings -ViewEntireForest $true

<#
    Check, if output folder exists, if so delete first, then create new folder
#>
function Check-OutputFolder {
    $global:TargetFolder = Join-Path -Path $ScriptDir -ChildPath $FolderName
    if(Test-Path -Path $global:TargetFolder) {
        # remove output folder, if exists
        Remove-Item $global:TargetFolder -Force -Confirm:$false -Recurse
    }
    # create output folder
    New-Item -ItemType Directory -Path $global:TargetFolder | Out-Null    
}

<#
    Ensure that there is no existing zip archive before creating the archive
#>
function Zip-OutputFolder {
    if($Prefix -eq '') {
         $target = Join-Path -Path $ScriptDir -ChildPath $ZipFileName
    }
    else {
        $target = Join-Path -Path $ScriptDir -ChildPath "$($prefix)-$($ZipFileName)"
    }
    if(Test-Path $target) {Remove-Item $target -Force -Confirm:$false | Out-Null}
    Add-Type -AssemblyName 'System.IO.Compression.Filesystem'
    [IO.Compression.Zipfile]::CreateFromDirectory($global:TargetFolder, $target)
    $target
}

<#
    Check params
#>
function Check-SendMail {
     if( ($SendMail) -and ($MailFrom -ne '') -and ($MailTo -ne '') -and ($MailServer -ne '') ) {
        return $true
     }
     else {
        return $false
     }
}

<#
    Main function to gather data from Exchange organization
#>
function Get-Data {
param (
    [string]$Source,
    [string]$Output,
    [string]$Command,
    [string]$Unlimited,
    [string]$Sort
)
    # fetch Exchange data
    Write-Verbose "Fetching $($Source), Command = $($Command), Unlimited = $($Unlimited)"
    $file = Join-Path "$($Prefix)$($Source)-$($date).txt" -Path $global:TargetFolder

    # prepare result size settings for PowerShell expression
    switch($Unlimited) {
        'false' {$resultSize = ''}
        default {$resultSize = '-ResultSize Unlimited'}
    }

    # prepare sorting for PowerShell expression
    switch($Sort) {
        '' { $sorting = '' }
        default { $sorting = "| Sort-Object $($Sort)" }
    }

    # gather the data
    switch($Output) {
        'SendConnector' {
            # Special send connector 
            $file = Join-Path "$($Prefix)$($Source)-$($date).csv" -Path $global:TargetFolder
            
            # get send connectors and expand all address spaces
            $expr = "Get-$($Source) $($Command) $($resultSize)$($sorting)| " + 'Select-Object Identity,@{N=''AddressSpaces'';E={[string]::join('';'',($_.AddressSpaces))}},Enabled' + " | Export-Csv $($file) -Encoding utf8 -NoTypeInformation -Force"
            
            # Write-Verbose $expr
            Invoke-Expression $expr
        }
        'VirtualDirectory' {
            if($Command -eq '') {
                # vDirs from all Exchange servers
                $expr = "Get-ExchangeServer | Sort-Object | Get-$($Source) | FT -AutoSize | Out-File $($file) -Encoding UTF8 -Force"
                Invoke-Expression $expr
                
                $expr = "Get-ExchangeServer | Sort-Object | Get-$($Source) | FL | Out-File $($file) -Encoding UTF8 -Append"
                Invoke-Expression $expr
            }
            else {
                # vDirs from Exchange 2013+ servers
                $expr = "Get-ExchangeServer | $($Command) | Sort-Object | Get-$($Source) | FT -AutoSize | Out-File $($file) -Encoding UTF8 -Force"
                Invoke-Expression $expr
               
                $expr = "Get-ExchangeServer | $($Command) | Sort-Object | Get-$($Source) | FL | Out-File $($file) -Encoding UTF8 -Append"
                Invoke-Expression $expr
            }
        }
        'DistributionGroup' {
            # fetch all distributions groups
            $expr = "Get-$($Source) $($Command) $($resultSize)$($sorting) | FT -AutoSize | Out-File $($file) -Encoding UTF8 -Force"
            Invoke-Expression $expr
            
            $expr = "Get-$($Source) $($Command) $($resultSize)$($sorting) | FL | Out-File $($file) -Encoding UTF8 -Append"
            Invoke-Expression $expr

            # fetch all room lists
            $Command = '-RecipientTypeDetails RoomList'
            $file = Join-Path "$($Prefix)$($Source)-RoomList-$($date).txt" -Path $global:TargetFolder
            
            $expr = "Get-$($Source) $($Command) $($resultSize)$($sorting) | FT -AutoSize | Out-File $($file) -Encoding UTF8 -Force"
            Invoke-Expression $expr
            
            $expr = "Get-$($Source) $($Command) $($resultSize)$($sorting) | FL | Out-File $($file) -Encoding UTF8 -Append"
            Invoke-Expression $expr
        }
        'DetailedOnly' {
            # just the details, please
            $expr = "Get-$($Source) $($Command) $($resultSize)$($sorting)| FL | Out-File $($file) -Encoding UTF8 -Force"
            Invoke-Expression $expr
            }
        default {
            # OverviewDetailed
            $expr = "Get-$($Source) $($Command) $($resultSize)$($sorting) | FT -AutoSize | Out-File $($file) -Encoding UTF8 -Force"
            Invoke-Expression $expr
            
            $expr = "Get-$($Source) $($Command) $($resultSize)$($sorting) | FL | Out-File $($file) -Encoding UTF8 -Append"
            Invoke-Expression $expr
        }
    }
}

## MAIN #######################################

# Prepare Out put folder
Check-OutputFolder

$step = 1
$infoSources.GetEnumerator() | ForEach-Object {
    # some nice progress bar, as the script will take some time
    Write-Progress -Activity 'Fetching data' -Status "Working on Get-$($_.Key) [Step $($step) / $($infoSources.Count)]" -PercentComplete (($step/$infoSources.Count)*100)
    
    # do it
    Get-Data -Source $_.Key -Output $infoSources[$_.Key].Output -Command $infoSources[$_.Key].Command -Unlimited $infoSources[$_.Key].Unlimited

    $step++
}

if($Zip) {
    # zip?
    $zipArchive = Zip-OutputFolder

    if($SendMail) {
        $Subject = 'Exchange Organizational Info'
        $Body = 'The attached zip file contains Exchange configuration data.'
        if(Test-Path $zipArchive) {
            Send-MailMessage -From $MailFrom -To $Mailto -SmtpServer $MailServer -Body $Body -BodyAsHtml -Subject $Subject -Attachments $zipArchive
        }
    }
}