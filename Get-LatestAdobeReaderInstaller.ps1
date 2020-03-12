<#PSScriptInfo
.VERSION 1.0.0
.GUID e4d9eb84-bf65-4985-a5b4-9bcbe20afb05
.AUTHOR NickolajA
.DESCRIPTION Get the latest Adobe Reader DC setup installation details from the official Adobe FTP server
.COMPANYNAME SCConfigMgr
.COPYRIGHT 
.TAGS AdobeReader Intune ConfigMgr PowerShell FTP
.LICENSEURI 
.PROJECTURI https://github.com/SCConfigMgr/Other/blob/master/Get-LatestAdobeReaderInstaller.ps1
.ICONURI 
.EXTERNALMODULEDEPENDENCIES 
.REQUIREDSCRIPTS 
.EXTERNALSCRIPTDEPENDENCIES 
.RELEASENOTES
#>
<#
.SYNOPSIS
    Get the latest Adobe Reader DC setup installation details from the official Adobe FTP server.

.DESCRIPTION
    Get the latest Adobe Reader DC setup installation details from the official Adobe FTP server.

.PARAMETER Type
    Specify the installer type, either EXE or MSP.

.PARAMETER Language
    Specify the desired language of the installer, e.g. 'en_US'.

.EXAMPLE
    # Retrieve the latest available Adobe Reader DC setup installer of type 'EXE' from the official Adobe FTP server:
    .\Get-LatestAdobeReaderInstaller.ps1 -Type EXE -Language en_US

    # Retrieve the latest available Adobe Reader DC patch installer of type 'MSP' from the official Adobe FTP server:
    .\Get-LatestAdobeReaderInstaller.ps1 -Type MSP

.NOTES
    FileName:    Get-LatestAdobeReaderInstaller.ps1
    Author:      Nickolaj Andersen
    Contact:     @NickolajA
    Created:     2020-03-12
    Updated:     2020-03-12
    
    Version history:
    1.0.0 - (2020-03-12) Script created.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [parameter(Mandatory = $false, HelpMessage = "Specify the installer type, either EXE or MSP.")]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("EXE", "MSP")]
    [string]$Type = "EXE",

    [parameter(Mandatory = $false, HelpMessage = "Specify the desired language of the installer, e.g. 'en_US'.")]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("en_US", "de_DE", "es_ES", "fr_FR", "ja_JP")]
    [string]$Language = "en_US"
)
Process {
    # Set script error action preference
    $ErrorActionPreference = "Stop"

    # Functions
    function Get-AdobeReaderFTPItem {
        [CmdletBinding(SupportsShouldProcess = $true)]
        param(
            [parameter(Mandatory = $false, HelpMessage = "Specify the directory path, e.g. 'ftp://ftp.adobe.com/pub/adobe/reader/win/AcrobatDC'.")]
            [ValidateNotNullOrEmpty()]
            [string]$Path = "ftp://ftp.adobe.com/pub/adobe/reader/win/AcrobatDC"
        )
        Process {
            # Construct anonymous credentials to use when connecting to Adobe's FTP
            $FTPCredentials = ([System.Management.Automation.PSCredential]::new("anonymous", ("password" | ConvertTo-SecureString -AsPlainText -Force)))
       
            # Construct WebRequest object for recieving FTP data stream
            [System.Net.FtpWebRequest]$WebRequest = [System.Net.WebRequest]::Create($Path)
            $WebRequest.Method = [System.Net.WebRequestMethods+FTP]::ListDirectoryDetails
            $WebRequest.Credentials = $FTPCredentials
            $WebRequest.Timeout = 90000
            $WebRequest.KeepAlive = $false
            $WebRequest.UseBinary = $false
            $WebRequest.UsePassive = $true
    
            try {
                # Get FTP response data stream
                $FTPResponse = $WebRequest.GetResponse()
                $FTPResponseStream = $FTPResponse.GetResponseStream()
                $FTPStreamReader = New-Object -TypeName System.IO.StreamReader -ArgumentList $FTPResponseStream

                # Read each line of the stream and add it a list
                $StreamList = New-Object -TypeName System.Collections.ArrayList
                while ($ListItem = $FTPStreamReader.ReadLine()) {
                    # Split directory listing string into objects (borrowed from PSFTP module from PSGallery: https://www.powershellgallery.com/packages/PSFTP)
                    $null, $null, $null, $null, $null, $null, $null, [string]$Date, [string]$Name = [regex]::Split($ListItem, '^([d-])([rwxt-]{9})\s+(\d{1,})\s+([.@A-Za-z0-9-]+)\s+([A-Za-z0-9-]+)\s+(\d{1,})\s+(\w+\s+\d{1,2}\s+\d{1,2}:?\d{2})\s+(.+?)\s?$', "SingleLine,IgnoreCase,IgnorePatternWhitespace")
                    
                    # Parse date string into date object (borrowed from PSFTP module from PSGallery: https://www.powershellgallery.com/packages/PSFTP)
                    $DatePart = $Date -split "\s+"
                    $NewDateString = "$($DatePart[0]) $('{0:D2}' -f [int]$DatePart[1]) $($DatePart[2])"
                    if($DatePart[2] -match ":") {
                        $Month = ([DateTime]::ParseExact($DatePart[0],"MMM" ,[System.Globalization.CultureInfo]::InvariantCulture)).Month
                        if((Get-Date).Month -ge $Month) {
                            $NewDate = [DateTime]::ParseExact($NewDateString, "MMM dd HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)
                        }
                        else {
                            $NewDate = ([DateTime]::ParseExact($NewDateString, "MMM dd HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)).AddYears(-1)
                        }
                    } 
                    else {
                        $NewDate = [DateTime]::ParseExact($NewDateString, "MMM dd yyyy", [System.Globalization.CultureInfo]::InvariantCulture)
                    }

                    # Construct custom object to be added to array list
                    $PSObject = [PSCustomObject]@{
                        Path = -join@($Path, "/", $Name.Trim())
                        Date = $NewDate
                        Name = $Name.Trim()
                    }

                    # Filter out unwanted objects and add everything else to array list
                    if ($Name -notlike "misc") {
                        $StreamList.Add($PSObject) | Out-Null
                    }
                }

                # Handle return value from function
                Write-Output -InputObject ($StreamList | Sort-Object -Property Date)
            }
            catch [System.Exception] {
                throw $_.Exception.Message; break
            }
        }
        End {
            # Perform cleanup and disconnect FTP connection
            $FTPResponse.Close()
            $FTPResponse.Dispose()
        }
    }

    function Get-LatestAdobeReaderInstallerItem {
        $FTPDirectoryItem = Get-AdobeReaderFTPItem | Select-Object -Skip $LatestCount -Last 1
        if ($FTPDirectoryItem -ne $null) {
            $FTPDirectoryItems = Get-AdobeReaderFTPItem -Path $FTPDirectoryItem.Path
            if ($FTPDirectoryItems -ne $null) {
                switch ($Type) {
                    "EXE" {
                        $FTPSetupInstaller = $FTPDirectoryItems | Where-Object { ($_.Name -match $FTPDirectoryItem.Name) -and ($_.Name -match $Language) -and ($_.Name -match $Type.ToLower()) }
                    }
                    "MSP" {
                        $FTPSetupInstaller = $FTPDirectoryItems | Where-Object { ($_.Name -match $FTPDirectoryItem.Name) -and ($_.Name -match $Type.ToLower()) }
                    }
                }
                
                if ($FTPSetupInstaller -ne $null) {
                    foreach ($FTPSetupInstallerItem in $FTPSetupInstaller) {
                        $PSObject = [PSCustomObject]@{
                            FileName = $FTPSetupInstallerItem.Name
                            SetupVersion = -join@($FTPDirectoryItem.Name.SubString(0, 2), ".", $FTPDirectoryItem.Name.SubString(2, 3), ".", $FTPDirectoryItem.Name.SubString(4, 5))
                            URL = $FTPSetupInstallerItem.Path
                            Date = $FTPSetupInstallerItem.Date
                        }
                        Write-Output -InputObject $PSObject
                    }
                }
                else {
                    $LatestCount++
                    Get-LatestAdobeReaderInstallerItem
                }
            }
        }
    }

    # Retrieve the latest setup installer based on parameter input
    $LatestCount = 0
    Get-LatestAdobeReaderInstallerItem
}