<#
.SYNOPSIS
Compiles a user audit report for AD users and optionally Entra ID users

.DESCRIPTION
Checks if the ImportExcel module is installed for automatic formatting of the report and saving to an xlsx file. It will prompt the user if they want to install if not found.
Checks if the Microsoft.Graph module is installed for compiling an audit report of both on-prem and cloud accounts .It will prompt the user if they want to install if not found.
	Prompts with an interactive logon to the tenant with only required permissions. Pulls all Entra ID users with relevant properties and GlobalAdmins.
Gets a list of all AD user accounts and all of their properties.
Processes user accounts
	Processes the AD user accounts into a single collection
	if Entra connection was successful, it processes Entra ID users and merges into a single report. Including merging timestamps for lastlogon datetime.
Generate and save a report
	if ImportExcel installed the report will be saved as an xlsx file with automatic sizing and conditional formatting
	if ImportExcel is not installed, report will be saved to a csv file. Formatting will have to be performed manually.

.INPUTS
None

.OUTPUTS
[file]"C:\Temp\$($domainName)_Users_Report_$TimeStamp.csv"
OR
[file]"C:\Temp\$($domainName)_Users_Report_$TimeStamp.xlsx"

.NOTES
Version:        1.0
Author:         Mark Newton
Creation Date:  06/23/2024
Purpose/Change: Initial script development
#>

##############################################################################################################
#                                            Globals (WRITABLE)                                              #
##############################################################################################################

# Dictionary of known service accounts and their descriptions. Descriptions are just for internal reference only.
$KnownServiceAccounts = @{
	"minable" = "RMM Admin Service Account"
	"rmm-service" = "RMM Admin Service Account"
	"svc-rmm" = "RMM Admin Service Account"
	"Sophos" = "Sophos Service Account"
	"SQLAdmin" = "SQL Admin Service Account"
	"VeeamAdmin" = "Veeam Admin Service Account"
	"asa2ldap" = "ASA LDAP Lookup Service Account"
	"axcientbackupadmin" = "Retired Axcient Backup Service Account"
	"mwservice" = "Retired Level Platforms Managed Workplace Service Account"
	"SBSMonAcct" = "Default Small Business Service Monitoring Service Account"
	"untangle2ldap" = "Untangle NG Firewall LDAP Lookup Service Account"
	"untangle" = "Untangle NG Firewall Service Account"
	"xerox2ldap" = "Xerox LDAP Lookup Service Account"
	"RoarAgent" = "Lionguard Roar Agent Service Account"
	"wlc2ldap" = "Wireless LAN Client LDAP Lookup Service Account"
	"icims" = "Internet Collaborative Information Management Systems Service Account"
	"svc-duo" = "DUO Auth Proxy Service Account"
	"svc-entra" = "Entra Connect Sync Account"
	"svc-ldap-duo" = "DUO Auth Proxy Service Account"
	"svc-knowbe4-adisync" = "KnowBe4 ADI Sync Service Account"
	"svc-ldap-ADSync" = "DUO Auth Proxy Service Account"
	"svc-liongard" = "Lionguard Service Account"
	"OpenDns_Connector" = "OpenDNS Connector Service Account"
	"svc-crestron" = "Crestron Flex Scheduling Service Account"
	"svc-ldap-zix" = "ZIX Service Account"
	"svc-scanner" = "Scanner Service Account"
	"svc_sophos" = "Sophos Service Account"
	"svc_sophosxg" = "Sophos Service Account"
	"svc-maillist" = "Email List Service Account"
	"ldap" = "LDAP Lookup Service Account"
	"krbtgt" = "Kerberos AD Service Account"
	"Guest" = "Guest AD Service Account"
}

##############################################################################################################
#                                                Functions                                                   #
##############################################################################################################

function Write-Color {
	<#
    .SYNOPSIS
    Write-Color is a wrapper around Write-Host delivering a lot of additional features for easier color options.

    .DESCRIPTION
    Write-Color is a wrapper around Write-Host delivering a lot of additional features for easier color options.

    It provides:
    - Easy manipulation of colors,
    - Logging output to file (log)
    - Nice formatting options out of the box.
    - Ability to use aliases for parameters

    .PARAMETER Text
    Text to display on screen and write to log file if specified.
    Accepts an array of strings.

    .PARAMETER Color
    Color of the text. Accepts an array of colors. if more than one color is specified it will loop through colors for each string.
    if there are more strings than colors it will start from the beginning.
    Available colors are: Black, DarkBlue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow, Gray, DarkGray, DarkBlue, Green, Cyan, Red, Magenta, Yellow, White

    .PARAMETER BackGroundColor
    Color of the background. Accepts an array of colors. if more than one color is specified it will loop through colors for each string.
    if there are more strings than colors it will start from the beginning.
    Available colors are: Black, DarkBlue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow, Gray, DarkGray, DarkBlue, Green, Cyan, Red, Magenta, Yellow, White

    .PARAMETER Center
    Calculates the window width and inserts spaces to make the text center according to the present width of the powershell window. Default is false.

    .PARAMETER StartTab
    Number of tabs to add before text. Default is 0.

    .PARAMETER LinesBefore
    Number of empty lines before text. Default is 0.

    .PARAMETER LinesAfter
    Number of empty lines after text. Default is 0.

    .PARAMETER StartSpaces
    Number of spaces to add before text. Default is 0.

    .PARAMETER LogFile
    Path to log file. if not specified no log file will be created.

    .PARAMETER DateTimeFormat
    Custom date and time format string. Default is yyyy-MM-dd HH:mm:ss

    .PARAMETER LogTime
    if set to $true it will add time to log file. Default is $true.

    .PARAMETER LogRetry
    Number of retries to write to log file, in case it can't write to it for some reason, before skipping. Default is 2.

    .PARAMETER Encoding
    Encoding of the log file. Default is Unicode.

    .PARAMETER ShowTime
    Switch to add time to console output. Default is not set.

    .PARAMETER NoNewLine
    Switch to not add new line at the end of the output. Default is not set.

    .PARAMETER NoConsoleOutput
    Switch to not output to console. Default all output goes to console.

    .EXAMPLE
    Write-Color -Text "Red ", "Green ", "Yellow " -Color Red,Green,Yellow

    .EXAMPLE
    Write-Color -Text "This is text in Green ",
                      "followed by red ",
                      "and then we have Magenta... ",
                      "isn't it fun? ",
                      "Here goes DarkCyan" -Color Green,Red,Magenta,White,DarkCyan

    .EXAMPLE
    Write-Color -Text "This is text in Green ",
                      "followed by red ",
                      "and then we have Magenta... ",
                      "isn't it fun? ",
                      "Here goes DarkCyan" -Color Green,Red,Magenta,White,DarkCyan -StartTab 3 -LinesBefore 1 -LinesAfter 1

    .EXAMPLE
    Write-Color "1. ", "Option 1" -Color Yellow, Green
    Write-Color "2. ", "Option 2" -Color Yellow, Green
    Write-Color "3. ", "Option 3" -Color Yellow, Green
    Write-Color "4. ", "Option 4" -Color Yellow, Green
    Write-Color "9. ", "Press 9 to exit" -Color Yellow, Gray -LinesBefore 1

    .EXAMPLE
    Write-Color -LinesBefore 2 -Text "This little ","message is ", "written to log ", "file as well." `
                -Color Yellow, White, Green, Red, Red -LogFile "C:\testing.txt" -TimeFormat "yyyy-MM-dd HH:mm:ss"
    Write-Color -Text "This can get ","handy if ", "want to display things, and log actions to file ", "at the same time." `
                -Color Yellow, White, Green, Red, Red -LogFile "C:\testing.txt"

    .EXAMPLE
    Write-Color -T "My text", " is ", "all colorful" -C Yellow, Red, Green -B Green, Green, Yellow
    Write-Color -t "my text" -c yellow -b green
    Write-Color -text "my text" -c red

    .EXAMPLE
    Write-Color -Text "Testuję czy się ładnie zapisze, czy będą problemy" -Encoding unicode -LogFile 'C:\temp\testinggg.txt' -Color Red -NoConsoleOutput

    .NOTES
    Understanding Custom date and time format strings: https://learn.microsoft.com/en-us/dotnet/standard/base-types/custom-date-and-time-format-strings
    Project support: https://github.com/EvotecIT/PSWriteColor
    Original idea: Josh (https://stackoverflow.com/users/81769/josh)

    #>
	[Alias('Write-Colour')]
	[CmdletBinding()]
	param(
		[Alias('T')] [String[]]$Text,
		[Alias('C','ForegroundColor','FGC')] [ConsoleColor[]]$Color = [ConsoleColor]::White,
		[Alias('B','BGC')] [ConsoleColor[]]$BackGroundColor = $null,
		[bool]$VerticalCenter = $False,
		[bool]$HorizontalCenter = $False,
		[Alias('Indent')] [int]$StartTab = 0,
		[int]$LinesBefore = 0,
		[int]$LinesAfter = 0,
		[int]$StartSpaces = 0,
		[Alias('L')] [string]$LogFile = '',
		[Alias('DateFormat','TimeFormat')] [string]$DateTimeFormat = 'yyyy-MM-dd HH:mm:ss',
		[Alias('LogTimeStamp')] [bool]$LogTime = $true,
		[int]$LogRetry = 2,
		[ValidateSet('unknown','string','unicode','bigendianunicode','utf8','utf7','utf32','ascii','default','oem')] [string]$Encoding = 'Unicode',
		[switch]$ShowTime,
		[switch]$NoNewLine,
		[Alias('HideConsole')] [switch]$NoConsoleOutput
	)
	if (-not $NoConsoleOutput) {
		$DefaultColor = $Color[0]
		if ($null -ne $BackGroundColor -and $BackGroundColor.Count -ne $Color.Count) {
			Write-Error "Colors, BackGroundColors parameters count doesn't match. Terminated."
			return
		}
		if ($VerticalCenter) {
			for ($i = 0; $i -lt ([math]::Max(0,$Host.UI.RawUI.BufferSize.Height / 4)); $i++) {
				Write-Host -Object "`n" -NoNewline
			}
		} # Center the output vertically according to the powershell window size
		if ($LinesBefore -ne 0) {
			for ($i = 0; $i -lt $LinesBefore; $i++) {
				Write-Host -Object "`n" -NoNewline
			}
		} # Add empty line before
		if ($HorizontalCenter) {
			$MessageLength = 0
			foreach ($Value in $Text) {
				$MessageLength += $Value.Length
			}
			Write-Host ("{0}" -f (' ' * ([math]::Max(0,$Host.UI.RawUI.BufferSize.Width / 2) - [math]::Floor($MessageLength / 2)))) -NoNewline
		} # Center the line horizontally according to the powershell window size
		if ($StartTab -ne 0) {
			for ($i = 0; $i -lt $StartTab; $i++) {
				Write-Host -Object "`t" -NoNewline
			}
		} # Add TABS before text

		if ($StartSpaces -ne 0) {
			for ($i = 0; $i -lt $StartSpaces; $i++) {
				Write-Host -Object ' ' -NoNewline
			}
		} # Add SPACES before text
		if ($ShowTime) {
			Write-Host -Object "[$([datetime]::Now.ToString($DateTimeFormat))] " -NoNewline -ForegroundColor DarkGray
		} # Add Time before output
		if ($Text.Count -ne 0) {
			if ($Color.Count -ge $Text.Count) {
				# the real deal coloring
				if ($null -eq $BackGroundColor) {
					for ($i = 0; $i -lt $Text.Length; $i++) {
						Write-Host -Object $Text[$i] -ForegroundColor $Color[$i] -NoNewline

					}
				} else {
					for ($i = 0; $i -lt $Text.Length; $i++) {
						Write-Host -Object $Text[$i] -ForegroundColor $Color[$i] -BackgroundColor $BackGroundColor[$i] -NoNewline

					}
				}
			} else {
				if ($null -eq $BackGroundColor) {
					for ($i = 0; $i -lt $Color.Length; $i++) {
						Write-Host -Object $Text[$i] -ForegroundColor $Color[$i] -NoNewline

					}
					for ($i = $Color.Length; $i -lt $Text.Length; $i++) {
						Write-Host -Object $Text[$i] -ForegroundColor $DefaultColor -NoNewline

					}
				}
				else {
					for ($i = 0; $i -lt $Color.Length; $i++) {
						Write-Host -Object $Text[$i] -ForegroundColor $Color[$i] -BackgroundColor $BackGroundColor[$i] -NoNewline

					}
					for ($i = $Color.Length; $i -lt $Text.Length; $i++) {
						Write-Host -Object $Text[$i] -ForegroundColor $DefaultColor -BackgroundColor $BackGroundColor[0] -NoNewline

					}
				}
			}
		}
		if ($NoNewLine -eq $true) {
			Write-Host -NoNewline
		}
		else {
			Write-Host
		} # Support for no new line
		if ($LinesAfter -ne 0) {
			for ($i = 0; $i -lt $LinesAfter; $i++) {
				Write-Host -Object "`n" -NoNewline
			}
		} # Add empty line after
	}
	if ($Text.Count -and $LogFile) {
		# Save to file
		$TextToFile = ""
		for ($i = 0; $i -lt $Text.Length; $i++) {
			$TextToFile += $Text[$i]
		}
		$Saved = $false
		$Retry = 0
		do {
			$Retry++
			try {
				if ($LogTime) {
					"[$([datetime]::Now.ToString($DateTimeFormat))] $TextToFile" | Out-File -FilePath $LogFile -Encoding $Encoding -Append -ErrorAction Stop -Whatif:$false
				}
				else {
					"$TextToFile" | Out-File -FilePath $LogFile -Encoding $Encoding -Append -ErrorAction Stop -Whatif:$false
				}
				$Saved = $true
			}
			catch {
				if ($Saved -eq $false -and $Retry -eq $LogRetry) {
					Write-Warning "Write-Color - Couldn't write to log file $($_.Exception.Message). Tried ($Retry/$LogRetry))"
				}
				else {
					Write-Warning "Write-Color - Couldn't write to log file $($_.Exception.Message). Retrying... ($Retry/$LogRetry)"
				}
			}
		} until ($Saved -eq $true -or $Retry -ge $LogRetry)
	}
}

function Initialize-ImportExcel {
	# Check if ImportExcel module is installed
	if (Get-Module -ListAvailable -Name 'ImportExcel') {
		Write-Color -Text "ImportExcel module detected. Will save directly to XLSX with automated formatting..." -Color Green -ShowTime

		# Import the ImportExcel module and set the $ImportExcel variable to True
		Import-Module ImportExcel
		$ImportExcel = $True
		$RemoveImportExcel = $False
	} else {
		# Check if we are running in an admin session. Otherwise skip trying to install the module and throw a warning to console.
		if ($AdminSession) {
			# ImportExcel module is not installed. Ask if allowed to install and user wants to install it.
			Write-Color -Text 'WARNING: ImportExcel module is not installed. Without it the report will output in CSV and you will have to format it manually.' -Color Yellow -ShowTime
			Write-Color -Text "if authorized to install modules on this system,"," would you like to install it for this script? ","(Y/N)" -Color Red,White,Yellow -NoNewline -ShowTime; $InstallImportExcel = Read-Host ' '

			switch ($InstallImportExcel) {
				"Y" {
					try {
						if ((Get-PSRepository).Name -contains "PSGallery") {
							if ((Get-PSRepository | Where-Object { $_.Name -eq 'PSGallery' }).InstallationPolicy -eq 'Untrusted') {
								Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
							} 
						} else {
							Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
						}

						if ((Get-PackageProvider).Name -notcontains 'NuGet') {
							[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
							Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
							Import-PackageProvider -Name NuGet -Force
						} 
						Write-Color -Text "Installing the ImportExcel module." -ShowTime
						Install-Module -Name 'ImportExcel' -Force
						Import-Module -Name 'ImportExcel'
						Write-Color -Text "ImportExcel module installed successfully." -Color Green -ShowTime
						$ImportExcel = $True
					} catch {
						Write-Color -Text "ERROR: ImportExcel module failed to install. See the error below. The report will output to CSV only until the error is corrected." -Color Red -ShowTime
						Write-Color -Text "Err Line: ","$($_.InvocationInfo.ScriptLineNumber)"," Err Name: ","$($_.Exception.GetType().FullName) "," Err Msg: ","$($_.Exception.Message)" -Color Red,Magenta,Red,Magenta,Red,Magenta -ShowTime
						$ImportExcel = $False
					}
				}
				"N" {
					Write-Color -Text "ImportExcel module will not be installed. ","Proceeding to save to CSV format." -Color White,Yellow -ShowTime
					$ImportExcel = $False
				}
				Default {
					Write-Color -Text "No option was selected. ","Proceeding to save to CSV format." -Color White,Yellow -ShowTime
					$ImportExcel = $False
				}
			}
		} else {
			Write-Color -Text "NOTICE: if authorized to install PowerShell modules on this system you can run this script in an admin session to install the ImportExcel module and save directly to xlsx with automated formatting" -Color Yellow -ShowTime
		}
	}

	return $ImportExcel
}

function Initialize-Entra {
	<#
    .DESCRIPTION
    Check if the user wants to connect to Entra ID and process cloud users.
    if the Microsoft.Graph module is not installed it will prompt the user if they want to install it.
    if the PSRepository or PackageProvider are modified or the module is installed, it will be removed at the end of the script.
    if a connection to Entra ID is successful then it grabs all Entra ID users and their relevant properties and a list of all global admins.

    .PARAMETER RemoveGraphAPI
    Optional parameter to allow the function to be run in a loop until successful connection or the user cancels
    Configures the script to remove the Microsoft.Graph module upon exit

    .PARAMETER [boolean]UntrustPSGallery
    Optional parameter to allow the function to be run in a loop until successful connection or the user cancels
    Configures the script to untrust the PSGallery upon exit

    .PARAMETER [boolean]RemovePSGallery
    Optional parameter to allow the function to be run in a loop until successful connection or the user cancels
    Configures the script to remove the PSGallery upon exit

    .PARAMETER [boolean]RemoveNuGet
    Optional parameter to allow the function to be run in a loop until successful connection or the user cancels
    Configures the script to remove the NuGet package manager upon exit

    .EXAMPLE
    [returning all the output variables without inputting any variables for initial first run of the function]
    $Entra, $PremiumEntraLicense, $AzUsers, $GlobalAdminMembers, $RemoveGraphAPI, $UntrustPSGallery, $RemovePSGallery, $RemoveNuGet = Initialize-Entra
    
	.EXAMPLE
    [To repeat the function until graph API connection or user cancels]
    Initialize-Entra -RemoveGraphAPI $RemoveGraphAPI -UntrustPSGallery $UntrustPSGallery -RemovePSGallery $RemovePSGallery -RemoveNuGet $RemoveNuGet
    #>

	Write-Color -Text "Would you like to connect to Entra ID? ","(Y/N)" -Color White,Yellow -NoNewline -ShowTime; $EntraID = Read-Host ' '
	
	switch ($EntraID) {
		'Y' {
			$Modules = Get-Module -ListAvailable | Select-Object -ExpandProperty Name
			if ($Modules -Contains "Microsoft.Graph.Authentication" -and $Modules -Contains "Microsoft.Graph.Users" -and $Modules -Contains "Microsoft.Graph.DirectoryObjects" -and $Modules -Contains "Microsoft.Graph.Identity.DirectoryManagement") {
				Write-Color -Text "Required Microsoft.Graph modules detected. Connecting to Graph API..." -Color Green -ShowTime

				# Import the ImportExcel module and set the $ImportExcel variable to True
				Import-Module Microsoft.Graph.Authentication
				Import-Module Microsoft.Graph.Users
				Import-Module Microsoft.Graph.DirectoryObjects
				Import-Module Microsoft.Graph.Identity.DirectoryManagement

				$GraphAPI = $True
			} else {
				# Check if we are running in an admin session. Otherwise skip trying to install the module and throw a warning to console.
				if ($AdminSession) {
					# Graph API module is not installed. Ask if allowed to install and user wants to install it.
					Write-Color -Text 'WARNING: Graph API modules required for this report are not installed. The report will display on-premises AD Users only.' -Color Yellow -ShowTime
					Write-Color -Text "if authorized to install modules on this system,"," would you like to install the required Graph API modules for this script? ","(Y/N)" -Color Red,White,Yellow -NoNewline -ShowTime; $InstallGraph = Read-Host ' '

					switch ($InstallGraph) {
						"Y" {
							try {
								Write-Color -Text "Installing the required Graph API modules." -ShowTime

								if ((Get-PSRepository).Name -contains "PSGallery") {
									if ((Get-PSRepository | Where-Object { $_.Name -eq 'PSGallery' }).InstallationPolicy -eq 'Untrusted') {
										Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
									} 
								} else {
									Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
								}

								if ((Get-PackageProvider).Name -notcontains "NuGet") {
									[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
									Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force
									Import-PackageProvider -Name NuGet -Force
								} 
								if ($Modules -NotContains 'Microsoft.Graph.Authentication') {
									Install-Module -Name 'Microsoft.Graph.Authentication' -Force
								}
								if ($Modules -NotContains 'Microsoft.Graph.Users') {
									Install-Module -Name 'Microsoft.Graph.Users' -Force
								}
								if ($Modules -NotContains 'Microsoft.Graph.DirectoryObjects') {
									Install-Module -Name 'Microsoft.Graph.DirectoryObjects' -Force
								}
								if ($Modules -NotContains 'Microsoft.Graph.Identity.DirectoryManagement') {
									Install-Module -Name 'Microsoft.Graph.Identity.DirectoryManagement' -Force
								}
								Write-Color -Text "Microsoft.Graph modules installed successfully" -Color Green -ShowTime
								Import-Module Microsoft.Graph.Authentication
								Import-Module Microsoft.Graph.Users
								Import-Module Microsoft.Graph.DirectoryObjects
								Import-Module Microsoft.Graph.Identity.DirectoryManagement

								$GraphAPI = $True
							} catch {
								Write-Color -Text "ERROR: Microsoft.Graph module failed to install. See the error below. The report will output to CSV only until the error is corrected." -Color Red -ShowTime
								Write-Color -Text "Err Line: ","$($_.InvocationInfo.ScriptLineNumber)","Err Name: ","$($_.Exception.GetType().FullName) ","Err Msg: ","$($_.Exception.Message)" -Color Red,Magenta,Red,Magenta,Red,Magenta -ShowTime
								$GraphAPI = $False
							}
						}
						"N" {
							Write-Color -Text "Graph API module will not be installed. ","Report will show on-premises AD users only." -Color White,Yellow -ShowTime
							$GraphAPI = $False
						}
						Default {
							Write-Color -Text "No option was selected. ","Graph API module will not be installed. Report will show on-premises AD users only." -Color White,Yellow -ShowTime
							$GraphAPI = $False
						}
					}
				} else {
					Write-Color -Text "NOTICE: if authorized to install PowerShell modules on this system you can run this script in an admin session to install the Graph API module and run the audit against cloud users and combine cloud properties with on-prem properties." -Color Yellow -ShowTime
				}
			}

			# if Microsoft.Graph modules were successfully installed
			if ($GraphAPI) {
				try {
					# Interactive login to the tenant requesting the required permissions only
					Connect-MgGraph -Scopes 'Directory.Read.All, User.Read.All, AuditLog.Read.All' -NoWelcome -ErrorAction Stop
					try {
						# Try to get all users including SignInActivity which is only available with a premium license
						$AzUsers = Get-MgUser -All -Property Id,UserPrincipalName,SignInActivity,OnPremisesSyncEnabled,displayName,samAccountName,AccountEnabled,mail,lastPasswordChangeDateTime,PasswordPolicies,CreatedDateTime,OnPremisesSyncEnabled,OnPremisesUserPrincipalName,OnPremisesSamAccountName,OnPremisesDomainName,OnPremisesSecurityIdentifier -ErrorAction Stop
						$PremiumEntraLicense = $True
					} catch {
						# if the tenant doesnt have a premium license get all users without including SignInActivity
						if ($_.Exception.Message -like "*Neither tenant is B2C or tenant doesn't have premium license*") {
							Write-Color -Text "WARNING: This tenant does not have a premium license. LastLogonDate will show on-premises AD datetimes only!" -Color Yellow -ShowTime
							$AzUsers = Get-MgUser -All -Property Id,UserPrincipalName,OnPremisesSyncEnabled,displayName,samAccountName,AccountEnabled,mail,lastPasswordChangeDateTime,PasswordPolicies,CreatedDateTime,OnPremisesSyncEnabled,OnPremisesUserPrincipalName,OnPremisesSamAccountName,OnPremisesDomainName,OnPremisesSecurityIdentifier -ErrorAction Stop
							$PremiumEntraLicense = $False
						}
					}

					$GlobalAdminRoleId = Get-MgDirectoryRole | Where-Object { $_.DisplayName -eq "Global Administrator" } | Select-Object -ExpandProperty ID
					$GlobalAdminMembers = Get-MgDirectoryRoleMemberAsUser -DirectoryRoleId $GlobalAdminRoleId
					$Entra = $True
				} catch {
					Write-Color -Text "ERROR: Connection to Graph API failed!" -Color Red -ShowTime
					Write-Color -Text "Err Line: ","$($_.InvocationInfo.ScriptLineNumber)","Err Name: ","$($_.Exception.GetType().FullName) ","Err Msg: ","$($_.Exception.Message)" -Color Red,Magenta,Red,Magenta,Red,Magenta -ShowTime
					Write-Color -Text "Would you like to try connecting to the Graph API again? ","(Y/N)" -Color White,Yellow -NoNewline -ShowTime; $TryAgain = Read-Host ' '
					switch ($TryAgain) {
						"Y" {
							Initialize-Entra -RemoveGraphAPI $RemoveGraphAPI -UntrustPSGallery $UntrustPSGallery -RemovePSGallery $RemovePSGallery -RemoveNuGet $RemoveNuGet
						}
						"N" {
							Write-Color -Text "Graph API modules will not be used. ","Report will show on-premises AD users only." -Color White,Yellow -ShowTime
							$Entra = $False
							$PremiumEntraLicense = $False
							$AzUsers = $Null
							$GlobalAdminMembers = $Null
						}
					}
				}
			} else {
				Write-Color -Text "WARNING: Connection to Graph API failed. Report will show on-premises AD users only." -Color Yellow -ShowTime
				$Entra = $False
				$PremiumEntraLicense = $False
				$AzUsers = $Null
				$GlobalAdminMembers = $Null
			}
		}
		'N' {
			$Entra = $False
			$PremiumEntraLicense = $False
			$AzUsers = $Null
			$GlobalAdminMembers = $Null
		}
		Default {
			$Entra = $False
			$PremiumEntraLicense = $False
			$AzUsers = $Null
			$GlobalAdminMembers = $Null
		}
	}

	return $Entra,$PremiumEntraLicense,$AzUsers,$GlobalAdminMembers
}

function Get-ADServiceAccounts {
	param (
		[Parameter(Mandatory = $True)] $ADUsers
	)

	$FoundServiceAccounts = @{}

	# Get list of managed or group managed service accounts from AD
	$ADServiceAccounts = Get-ADServiceAccount -Filter *

	# Iterate through MSA or gMSA accounts if any were found
	if ($ADServiceAccounts.Count -gt 0) {
		foreach ($ServiceAccount in $ADServiceAccounts) {
			$FoundServiceAccounts[$ServiceAccount.samAccountName] = "MSA or gMSA Account"
		}
	}

	# Iterate through the AD accounts to check if they are known service accounts
	foreach ($User in $ADUsers) {
		# Iterate through the keys of the Known Service Accounts dictionary which contains the usernames of the service accounts
		foreach ($key in $KnownServiceAccounts.Keys) {
			# if the user is a known service account add it to the Found Service Accounts dictionary
			if ($KnownServiceAccounts.Keys -contains $User.samAccountName) {
				$FoundServiceAccounts[$key] = $KnownServiceAccounts[$key]
			} else {
				# Generic service account capture based on wildcard name comparison
				if ($User.samAccountName -like "*svc*") {
					$FoundServiceAccounts[$User.samAccountName] = "Unknown Generic Service Account"
				} elseif ($User.samAccountName -like "*MSOL_*") {
					$FoundServiceAccounts[$User.samAccountName] = "Microsoft Entra ID Connect Service Account"
				} elseif ($User.samAccountName -like "*AAD_*") {
					$FoundServiceAccounts[$User.samAccountName] = "Microsoft Entra ID Connect Service Account"
				} 
			}
		}
	}

	return $FoundServiceAccounts
}

function Get-RecommendedActions {
    <#
    .DESCRIPTION
    Processes user properties and provides recommendations for the account

    .PARAMETER [string]UserType
    Defines the type of user account (Cloud, On-Prem, Hybrid).

    .PARAMETER [datetime]AccountExpired
    Represents the datetime set to the AccountExpired property in AD.

    .PARAMETER [boolean]EnterpriseAdmin
    Indicates if the account is an enterprise admin (True/False).

    .PARAMETER [boolean]DomainAdmin
    Indicates if the account is a domain admin (True/False).

    .PARAMETER [boolean]GlobalAdmin
    Indicates if the account is a global admin (True/False).

    .PARAMETER [datetime]PasswordLastSet
    Represents the datetime when the password was last set.

    .PARAMETER [datetime]LastLogonDate
    Represents the datetime when the account was last logged on.

    .PARAMETER [boolean]PasswordNeverExpires
    Indicates if the password is set to never expire (True/False).

    .PARAMETER [boolean]PasswordExpired
    Indicates if the password is expired (True/False).

    .PARAMETER [boolean]LockedOut
    Indicates if the account is locked out (True/False).

    .PARAMETER [boolean]CannotChangePassword
    Indicates if the account cannot change its password (True/False).

    .PARAMETER [boolean]ServiceAccount
    Indicates if the account is a service account (True/False).

    .PARAMETER [string]SamAccountName
    Represents the SAM account name of the user.

    .PARAMETER [string]DistinguishedName
    Represents the distinguished name of the user in AD.
    #>

    param(
        [Parameter(Mandatory = $True)] $UserType,
        $Enabled = $Null,
        $AccountExpired = $Null, 
        $EnterpriseAdmin = $Null, 
        $DomainAdmin = $Null, 
        $GlobalAdmin = $Null, 
        $PasswordLastSet = $Null, 
        $LastLogonDate = $Null, 
        $PasswordNeverExpires = $Null, 
        $PasswordExpired = $Null, 
        $LockedOut = $Null, 
        $CannotChangePassword = $Null,
        $ServiceAccount = $Null,
        $SamAccountName = $Null,
        $DistinguishedName = $Null
    )

    $RecommendedActions = [System.Collections.Generic.List[string]]::new()

    # If account is known service account, skip the rest of the validation checks
    if ($ServiceAccount) {
        if ($Enabled -eq $False) {
            $RecommendedActions.Add("Verify if this service account is needed since it is disabled")
        } else {
            # Check for duplicate Entra connect accounts to recommend cleaning them up
            if ($SamAccountName -like "*MSOL_*" -or $SamAccountName -like "*AAD_*") {
                $DupSyncUser = [System.Collections.Generic.List[string]]::new()
                foreach ($User in $ADUsers) {
                    if ($SamAccountName -like "*MSOL_*" -or $SamAccountName -like "*AAD_*") {
                        $DupSyncUser.Add($SamAccountName)
                    }
                }
                if ($DupSyncUser.Count -gt 1) {
                    $RecommendedActions.Add("Multiple Entra Connect Sync accounts found. Verify the current active sync account and old accounts should be disabled and a delta sync ran to verify if the account is still needed for the database or can be removed")
                    return $RecommendedActions -join "; "
                }
            }
            # If nothing found, then return an empty string
            return " "
        }
    } else {
        # Account has an expiration date
        # Account is not enabled. Nothing to recommend.
        if ($Null -ne $AccountExpired -and $Enabled -eq $False) {
            # If account is expired and not enabled, then return an empty string
            return " "
        # Account is enabled.
        } elseif ($Null -ne $AccountExpired -and $Enabled -eq $True) {
            # Expiration date is in the past. Recommend disabling the account.
            if ($AccountExpired -lt (Get-Date)) {
                $RecommendedActions.Add("Disable account. Account has already expired")
                $Recommended = $RecommendedActions -join "; "
    			return $Recommended
            # Expiration date is in the future. Recommend verifying the account should expire.
            } else {
                $RecommendedActions.Add("Verify if this account should expire")
            }
        # Account not set to expire
        } elseif ($Null -eq $AccountExpired) {
            # AD User account
            if ($UserType -eq "On-Prem" -or $UserType -eq "Hybrid") {
                if ($Enabled -eq $False -and $DisabledOU -eq $True -and $DistinguishedName -notlike "*Disabled User*") {
                    $RecommendedActions.Add("Create a Disabled Users OU in an Entra Connect synced OU in AD and move this account into it")
                }
                if ($EnterpriseAdmin -eq $True -or $DomainAdmin -eq $True -or $GlobalAdmin -eq $True -and $Enabled -eq $False) {
                    if ($EnterpriseAdmin) {
                        $RecommendedActions.Add("Check group memberships, account is disabled and a member of Enterprise Admins group")
                    }

                    if ($DomainAdmin) {
                        $RecommendedActions.Add("Check group memberships, account is disabled and a member of Domain Admins group")
                    }

                    if ($GlobalAdmin) {
                        $RecommendedActions.Add("Check group memberships, account is disabled and a member of Global Admins group")
                    }
                } elseif ($EnterpriseAdmin -eq $True -or $DomainAdmin -eq $True -or $GlobalAdmin -eq $True -and $Enabled -eq $True) {
                    if ($EnterpriseAdmin) {
                        $RecommendedActions.Add("Verify group memberships, account is member of Enterprise Admins group")
                    }

                    if ($DomainAdmin) {
                        $RecommendedActions.Add("Verify group memberships, account is member of Domain Admins group")
                    }

                    if ($GlobalAdmin) {
                        $RecommendedActions.Add("Verify group memberships, account is member of Global Admins group")
                    }
                }

                if ($PasswordLastSet) {
                    if ($PasswordLastSet -lt (Get-Date).AddDays(-90)) {
                        if ($NoExpiry -eq $False) {
                            $RecommendedActions.Add("Password has not been changed for over 90 days")
                        }
                    }
                }

                if ($LastLogonDate) {
                    if ($LastLogonDate -lt (Get-Date).AddDays(-90) -and $LastLogonDate -ge (Get-Date).AddDays(-180)) {
                        $RecommendedActions.Add("Verify if account still needed. Account has not been logged in for over 90 days")
                    } elseif ($LastLogonDate -lt (Get-Date).AddDays(-180)) {
                        $RecommendedActions.Add("Verify if account still needed. Account has not been logged in for over 180 days")
                    }
                }

                if ($PasswordNeverExpires) {
                    if ($NoExpiry -eq $False) {
                        $RecommendedActions.Add("Verify if the password should never expire for this account")
                    }
                }

                if ($PasswordExpired -eq $True) {
                    $RecommendedActions.Add("Verify if this is an active user. The accounts password is expired")
                }

                if ($LockedOut -eq $True) {
                    $RecommendedActions.Add("Verify if this is an active user. This account is locked out")
                }

                if ($CannotChangePassword -eq $True) {
                    $RecommendedActions.Add("Verify if account should not be able to change their password")
                }
            # Cloud only user
            } else {
                if ($GlobalAdmin -eq $True -and $Enabled -eq $False) {
                    if ($GlobalAdmin) {
                        $RecommendedActions.Add("This user is disabled and a member of Global Admins group, verify if they can be removed from the Global Admins group")
                    }
                } elseif ($GlobalAdmin -eq $True -and $Enabled -eq $True) {
                    if ($GlobalAdmin) {
                        $RecommendedActions.Add("Verify group memberships, account is member of Global Admins group")
                    }
                }
            }
        }
    }

    return $RecommendedActions -join "; "
}

function Measure-ADUsers {
	<#
    .DESCRIPTION
    Processes the active directory domain user accounts.
    if a connection to Entra was established it will only process on on-premises AD user accounts only. Hybrid and cloud users will be processed later.

    .PARAMETER [object]ADUsers
    Collection of all AD Users and all of their properties

	.PARAMETER [boolean]Entra
    Boolean for if Graph API was used and connected to Entra

	.PARAMETER [object]ServiceAccounts
    Collection of all found service accounts in AD

    .PARAMETER [object]AzUsers
    Collection of all Entra ID Users and their relevent properties

    .EXAMPLE
    Measure-ADUsers -ADUsers $ADUsers -AzUsers $AzUsers -Entra $True
    Measure-ADUsers -ADUsers $ADUsers -AzUsers $AzUsers -Entra $False
    Measure-ADUsers -ADUsers $ADUsers -Entra $False
    #>

	param(
		[Parameter(Mandatory = $True)] $ADUsers,
		[Parameter(Mandatory = $True)] [boolean]$Entra,
		[Parameter(Mandatory = $True)] [object]$ServiceAccounts,
		$AzUsers
	)

	# Initialize arrays for UserCollection and AzUsersToProcess
	$UserCollection = @()
	#$AzUsersToProcess = @()

	# Initialize user counter for progress bar
	$Count = 1

	if ($ADUsers.Count -gt 0) {
		Write-Color -Text "Processing Active Directory Users" -ShowTime

		# Process each user account found in active directory
		foreach ($User in $ADUsers) {
			Write-Progress -Id 1 -Activity "Processing AD Users" -Status "Current Count: ($Count/$($ADUsers.Count))" -PercentComplete (($Count / $ADUsers.Count) * 100) -CurrentOperation "Processing... $($User.DisplayName)"

			# Check the users samAccountName against the list of Admin Users to verify if they are a domain admin
			if (($EnterpriseAdmins.SamAccountName) -contains $User.SamAccountName) {
				$EnterpriseAdmin = $True
			} else {
				$EnterpriseAdmin = $False
			}

			# Check the users samAccountName against the list of Admin Users to verify if they are a domain admin
			if (($DomainAdmins.SamAccountName) -contains $User.SamAccountName) {
				$DomainAdmin = $True
			} else {
				$DomainAdmin = $False
			}

			# if the account has an account expiration date then consider it expired.
			if ($Null -ne $User.AccountExpirationDate) {
				$AccountExpired = $User.AccountExpirationDate
			} else {
				$AccountExpired = $Null
			}

			# if email property is blank then set to a blank space for formatting the spreadsheet. This stops a previous column from displaying over it.
			if ($Null -eq $User.mail) {
				$Mail = " "
			} else {
				$Mail = $User.mail
			}

			# Check if SamAccountName corresponds to a known service account
			if ($ServiceAccounts.Keys -contains $User.SamAccountName) {
				$KnownServiceAccount = $True
			} else {
				$KnownServiceAccount = $False
			}

			if ($Entra) {

				# On-prem user without synced cloud user
				if (($AzUsers).OnPremisesSecurityIdentifier -notcontains $User.SID) {
					# Get recommended actions for the user
					$Recommended = Get-RecommendedActions -UserType "On-Prem" -Enabled $User.Enabled -AccountExpired $AccountExpired -EnterpriseAdmin $EnterpriseAdmin -DomainAdmin $DomainAdmin -GlobalAdmin $False -PasswordLastSet $User.PasswordLastSet -LastLogonDate $LastLogonDate -PasswordNeverExpires $User.PasswordNeverExpires -PasswordExpired $User.PasswordExpired -LockedOut $User.lockedOut -CannotChangePassword $User.CannotChangePassword -ServiceAccount $KnownServiceAccount -SamAccountName $User.samAccountName -DistinguishedName $User.DistinguishedName

					# Add the user to the UserCollection
					$UserCollection += [pscustomobject]@{
						"Name" = $User.DisplayName
						SamAccountName = $User.SamAccountName
						"On-Prem UserPrincipalName" = $User.UserPrincipalName
						"Cloud UserPrincipalName" = "N/A"
						"Email Address" = $Mail
						"User Type" = "On-Prem"
						Enabled = $User.Enabled
						AccountExpiredDate = $AccountExpired
						EnterpriseAdmin = $EnterpriseAdmin
						DomainAdmin = $DomainAdmin
						"AzGlobalAdmin" = "N/A"
						"Known Service Account" = $KnownServiceAccount
						PasswordLastSet = $User.PasswordLastSet
						LastLogonDate = $User.LastLogonDate
						PasswordNeverExpires = $User.PasswordNeverExpires
						PasswordExpired = $User.PasswordExpired
						"Account Locked" = $User.lockedOut
						CannotChangePassword = $User.CannotChangePassword
						"Date Created" = $User.whenCreated
						"Recommended Actions" = $Recommended
						Notes = ""
						Action = ""
						"Follow Up" = ""
						Resolution = ""
					}
				} 
			# No connection to Entra ID. Process active directory users only
			} else {
				# Get recommended actions for the user
				$Recommended = Get-RecommendedActions -UserType "On-Prem" -Enabled $User.Enabled -AccountExpired $AccountExpired -EnterpriseAdmin $EnterpriseAdmin -DomainAdmin $DomainAdmin -GlobalAdmin $False -PasswordLastSet $User.PasswordLastSet -LastLogonDate $LastLogonDate -PasswordNeverExpires $User.PasswordNeverExpires -PasswordExpired $User.PasswordExpired -LockedOut $User.lockedOut -CannotChangePassword $User.CannotChangePassword -ServiceAccount $KnownServiceAccount -SamAccountName $User.samAccountName -DistinguishedName $User.DistinguishedName

				# Add the user to the UserCollection
				$UserCollection += [pscustomobject]@{
					"Name" = $User.DisplayName
					SamAccountName = $User.SamAccountName
					UserPrincipalName = $User.UserPrincipalName
					"Email Address" = $Mail
					"User Type" = "On-Prem"
					Enabled = $User.Enabled
					AccountExpiredDate = $AccountExpired
					EnterpriseAdmin = $EnterpriseAdmin
					DomainAdmin = $DomainAdmin
					"Known Service Account" = $KnownServiceAccount
					PasswordLastSet = $User.PasswordLastSet
					LastLogonDate = $User.LastLogonDate
					PasswordNeverExpires = $User.PasswordNeverExpires
					PasswordExpired = $User.PasswordExpired
					"Account Locked" = $User.lockedOut
					CannotChangePassword = $User.CannotChangePassword
					"Date Created" = $User.whenCreated
					"Recommended Actions" = $Recommended
					Notes = ""
					Action = ""
					"Follow Up" = ""
					Resolution = ""
				}
			}

			$Count += 1
		}
	}

	return $UserCollection
}

function Merge-AzUsers {
	<#
    .DESCRIPTION
    Processes the users in AzUsersToProcess for both hybrid and cloud users. 
    For hybrid users, the LastLogonTime is set according to the most recent timestamp.
    N/A is used for properties that cloud only users do not have in Entra ID.

    .PARAMETER [object]ADUsers
    Collection of all AD Users and all of their properties

    .PARAMETER [object]AzUsers
    Collection of all Entra ID Users and their relevent properties

    .PARAMETER [object]UserCollection
    Collection of all users that have already been processed

    .EXAMPLE
    Merge-AzUsers $ADUsers $AzUsers $UserCollection $ServiceAccounts
    #>

	param(
		[Parameter(Mandatory = $True)] [object]$ADUsers,
		[Parameter(Mandatory = $True)] [object]$AzUsers,
		[Parameter(Mandatory = $True)] [object]$UserCollection,
		[Parameter(Mandatory = $True)] [object]$ServiceAccounts
	)

	# Initialize user counter for progress bar
	$Count = 1

	Write-Color -Text "Processing Entra ID Users" -ShowTime

	foreach ($AzUser in $AzUsers) {
		Write-Progress -Id 1 -Activity "Processing Entra Users" -Status "Current Count: ($Count/$($AzUsers.Count))" -PercentComplete (($Count / $AzUsers.Count) * 100) -CurrentOperation "Processing... $($AzUser.DisplayName)"

		# Hybrid user
		if ($AzUser.OnPremisesSyncEnabled -eq $True) {
			$User = $ADUsers | Where-Object { $_.SID -eq $AzUser.OnPremisesSecurityIdentifier }

			if ($Null -eq $User.Enabled) {
				$Enabled = $AzUser.AccountEnabled
			} else {
				$Enabled = $User.Enabled
			}

			# Check the users samAccountName against the list of Admin Users to verify if they are a domain admin
			if (($EnterpriseAdmins.SamAccountName) -contains $User.SamAccountName) {
				$EnterpriseAdmin = $True
			} else {
				$EnterpriseAdmin = $False
			}

			# Check the users samAccountName against the list of Admin Users to verify if they are a domain admin
			if (($DomainAdmins.SamAccountName) -contains $User.SamAccountName) {
				$DomainAdmin = $True
			} else {
				$DomainAdmin = $False
			}

			# if the account has an account expiration date then consider it expired.
			if ($Null -ne $User.AccountExpirationDate) {
				$AccountExpired = $User.AccountExpirationDate
			} else {
				$AccountExpired = $Null
			}

			# Check if user is a global admin in Entra ID
			if (($GlobalAdminMembers).UserPrincipalName -contains $AzUser.UserPrincipalName) {
				$GlobalAdmin = $True
			} else {
				$GlobalAdmin = $False
			}

			# Check if user is a known service account
			if ($ServiceAccounts.Keys -contains $User.SamAccountName) {
				$KnownServiceAccount = $True
			} else {
				$KnownServiceAccount = $False
			}

			# if email property is blank then set to a blank space for formatting the spreadsheet. This stops a previous column from displaying over it.
			if ($Null -eq $User.mail) {
				$Mail = " "
			} else {
				$Mail = $User.mail
			}

			# if the tenant has a premium license then get and compare the last sign-in timestamp and lastLogonDate timestamp
			if ($PremiumEntraLicense) {
				if ($AzUser.signInActivity.lastSignInDateTime) {
					$AzlastLogonDate = [datetime]$AzUser.signInActivity.lastSignInDateTime
					# if the last sign-in timestamp is newer than set that as lastLogonDate property
					if ($User.LastLogonDate -lt $AzlastLogonDate) {
						$LastLogonDate = $AzlastLogonDate
						# Otherwise use the active directory lastLogonDate timestamp
					} else {
						$LastLogonDate = $User.LastLogonDate
					}
					# if there is no last sign-in timestamp then default to AD lastLogonDate timestamp.
				} else {
					$LastLogonDate = $User.LastLogonDate
				}
				# if the tenant doesnt have a premium license then we cant get last sign-in timestamp. Default to AD lastLogonDate timestamp.
			} else {
				$LastLogonDate = $User.LastLogonDate
			}

			# Get the recommended actions for the user
			$Recommended = Get-RecommendedActions -UserType "Hybrid" -Enabled $User.Enabled -AccountExpired $AccountExpired -EnterpriseAdmin $EnterpriseAdmin -DomainAdmin $DomainAdmin -GlobalAdmin $GlobalAdmin -PasswordLastSet $User.PasswordLastSet -LastLogonDate $LastLogonDate -PasswordNeverExpires $User.PasswordNeverExpires -PasswordExpired $User.PasswordExpired -LockedOut $User.lockedOut -CannotChangePassword $User.CannotChangePassword -ServiceAccount $KnownServiceAccount -SamAccountName $User.samAccountName -DistinguishedName $User.DistinguishedName
			
			# Add the user to the UserCollection
			$UserCollection += [pscustomobject]@{
				"Name" = $User.DisplayName
				SamAccountName = $User.SamAccountName
				"On-Prem UserPrincipalName" = $User.UserPrincipalName
				"Cloud UserPrincipalName" = $AzUser.UserPrincipalName
				"Email Address" = $Mail
				"User Type" = "Hybrid"
				Enabled = $Enabled
				AccountExpiredDate = $AccountExpired
				EnterpriseAdmin = $EnterpriseAdmin
				DomainAdmin = $DomainAdmin
				"AzGlobalAdmin" = $GlobalAdmin
				"Known Service Account" = $KnownServiceAccount
				PasswordLastSet = $User.PasswordLastSet
				LastLogonDate = $LastLogonDate
				PasswordNeverExpires = $User.PasswordNeverExpires
				PasswordExpired = $User.PasswordExpired
				"Account Locked" = $User.lockedOut
				CannotChangePassword = $User.CannotChangePassword
				"Date Created" = $User.whenCreated
				"Recommended Actions" = $Recommended
				Notes = ""
				Action = ""
				"Follow Up" = ""
				Resolution = ""
			}

		# Cloud only user
		} else {
			# Check if user is a global admin in Entra ID
			if (($GlobalAdminMembers).UserPrincipalName -contains $AzUser.UserPrincipalName) {
				$GlobalAdmin = $True
			} else {
				$GlobalAdmin = $False
			}

			# Check if user is a known cloud Sync_ service account
			if ($AzUser.UserPrincipalName -like "*Sync_*") {
				$KnownServiceAccount = $True
			} else {
				$KnownServiceAccount = $False
			}

			# if the tenant has a premium license then grab the last sign-in timestamp
			if ($PremiumEntraLicense) {
				if ($AzUser.signInActivity.lastSignInDateTime) {
					$LastLogonDate = [datetime]$AzUser.signInActivity.lastSignInDateTime
				} else {
					$LastLogonDate = $Null
				}
			} else {
				$LastLogonDate = $Null
			}

			# Check if password expiration is configured for the tenant. A value of 2147483647 indicates passwords dont expire.
			if (($AzDomains | Where-Object {$_.Id -eq ($AzUser.UserPrincipalName).Split("@")[1]} | Select-Object -ExpandProperty PasswordValidityPeriodInDays) -eq 2147483647) {
				$PasswordNeverExpires = $True
			# Else a password policy is defined and we need to check if the individual account has an exception set in the PasswordPolicies property for their account
			} else {
				# If DisablePasswordExpiration found in PasswordPolicies then the password is set to never expire in the cloud
				if ($AzUser.PasswordPolicies -like "*DisablePasswordExpiration*") {
					$PasswordNeverExpires = $True
				} else {
					$PasswordNeverExpires = $False
				}
			}

			# if email property is blank then set to a blank space for formatting the spreadsheet. This stops a previous column from displaying over it.
			if ($Null -eq $AzUser.mail) {
				$Mail = " "
			} else {
				$Mail = $AzUser.mail
			}

			# Get the recommended actions for the user
			$Recommended = Get-RecommendedActions -UserType "Cloud" -Enabled $User.Enabled -AccountExpired $AccountExpired -EnterpriseAdmin $EnterpriseAdmin -DomainAdmin $DomainAdmin -GlobalAdmin $GlobalAdmin -PasswordLastSet $User.PasswordLastSet -LastLogonDate $LastLogonDate -PasswordNeverExpires $User.PasswordNeverExpires -PasswordExpired $User.PasswordExpired -LockedOut $User.lockedOut -CannotChangePassword $User.CannotChangePassword -ServiceAccount $KnownServiceAccount

			# Add the user to the UserCollection
			$UserCollection += [pscustomobject]@{
				"Name" = $AzUser.DisplayName
				SamAccountName = "N/A"
				"On-Prem UserPrincipalName" = "N/A"
				"Cloud UserPrincipalName" = $AzUser.UserPrincipalName
				"Email Address" = $Mail
				"User Type" = "Cloud"
				Enabled = $AzUser.AccountEnabled
				AccountExpiredDate = "N/A"
				EnterpriseAdmin = $False
				DomainAdmin = "N/A"
				"AzGlobalAdmin" = $GlobalAdmin
				"Known Service Account" = $KnownServiceAccount
				PasswordLastSet = $AzUser.lastPasswordChangeDateTime
				LastLogonDate = $LastLogonDate
				PasswordNeverExpires = $PasswordNeverExpires
				PasswordExpired = "N/A"
				"Account Locked" = "N/A"
				CannotChangePassword = "N/A"
				"Date Created" = $AzUser.CreatedDateTime
				"Recommended Actions" = $Recommended
				Notes = ""
				Action = ""
				"Follow Up" = ""
				Resolution = ""
			}
		}

		# Increment counter for progress bar
		$Count += 1
	}

	return $UserCollection
}

##############################################################################################################
#                                                   Main                                                     #
##############################################################################################################
# Register the exit event to cleanup modules after the session has ended.
try {
	Clear-Host
	Write-Color -Text "__________________________________________________________________________________________" -Color White -BackgroundColor Black -HorizontalCenter $True -LinesBefore 7
	Write-Color -Text "|                                                                                          |" -Color White -BackgroundColor Black -HorizontalCenter $True
	Write-Color -Text "|","                                            .-.                                           ","|" -Color White,DarkBlue,White -BackgroundColor Black,Black,Black -HorizontalCenter $True
	Write-Color -Text "|","                                            -#-              -.    -+                     ","|" -Color White,DarkBlue,White -BackgroundColor Black,Black,Black -HorizontalCenter $True
	Write-Color -Text "|","    ....           .       ...      ...     -#-  .          =#:..  .:      ...      ..    ","|" -Color White,DarkBlue,White -BackgroundColor Black,Black,Black -HorizontalCenter $True
	Write-Color -Text "|","   +===*#-  ",".:","     #*  *#++==*#:   +===**:  -#- .#*    -#- =*#+++. +#.  -*+==+*. .*+-=*.  ","|" -Color White,DarkBlue,Cyan,DarkBlue,White -BackgroundColor Black,Black,Black,Black,Black -HorizontalCenter $True
	Write-Color -Text "|","    .::.+#  ",".:","     #*  *#    .#+   .::..**  -#-  .#+  -#=   =#:    +#. =#:       :#+:     ","|" -Color White,DarkBlue,Cyan,DarkBlue,White -BackgroundColor Black,Black,Black,Black,Black -HorizontalCenter $True
	Write-Color -Text "|","  =#=--=##. ",".:","     #*  *#     #+  **---=##  -#-   .#+-#=    =#:    +#. **          :=**.  ","|" -Color White,DarkBlue,Cyan,DarkBlue,White -BackgroundColor Black,Black,Black,Black,Black -HorizontalCenter $True
	Write-Color -Text "|","  **.  .*#. ",".:.","   =#=  *#     #+ :#=   :##  -#-    :##=     -#-    +#. :#*:  .:  ::  .#=  ","|" -Color White,DarkBlue,Cyan,DarkBlue,White -BackgroundColor Black,Black,Black,Black,Black -HorizontalCenter $True
	Write-Color -Text "|","   -+++--=      .==:   ==     =-  .=++=-==  :=:    .#=       -++=  -=    :=+++-. :=++=-   ","|" -Color White,DarkBlue,White -BackgroundColor Black,Black,Black -HorizontalCenter $True
	Write-Color -Text "|","                                                  .#+                                     ","|" -Color White,DarkBlue,White -BackgroundColor Black,Black,Black -HorizontalCenter $True
	Write-Color -Text "|","                                                  *+                                      ","|" -Color White,DarkBlue,White -BackgroundColor Black,Black,Black -HorizontalCenter $True
	Write-Color -Text "|__________________________________________________________________________________________|" -Color White -BackgroundColor Black -HorizontalCenter $True
	Write-Color -Text "Script: ","User Audit Report" -Color Yellow,White -HorizontalCenter $True -LinesBefore 1
	Write-Color -Text "Author: " ,"Mark Newton" -Color Yellow,White -HorizontalCenter $True -LinesAfter 1

	# Check if PowerShell is at least v5. That is the version that is used by default in Windows since Windows 7 and Server 2012.
	if ($PSVersionTable.PSVersion.Major -lt 5) {
		$ImportExcel = $False
		$Entra = $False
		$SupportedPS = $False
		Write-Color -Text "WARNING: The detected version of powershell on this system does not support installation of modules. Output will be on-prem AD users only and in CSV format only. You can run this script from another system with PowerShell 5.1 or higher to install modules and output to xlsx directly." -Color Yellow
	} else {
		$SupportedPS = $True
	}

	# Check if powershell is running in an admin session for the ability to install modules
	$currentPrincipal = New-Object Security.Principal.WindowsPrincipal ([Security.Principal.WindowsIdentity]::GetCurrent())
	$AdminSession = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

	# Initialize default value for checking if a disabled users OU exists in AD. This will later be set to True if one exists.
	$DisabledOU = $False

	# Check if we can import the active directory module. if not, then we can only produce a report for cloud users with the Graph API
	try {
		Import-Module ActiveDirectory
		$ActiveDirectory = $True
	} catch {
		if ($Entra) {
			Write-Color -Text "WARNING: ActiveDirectory module could not be imported. You will only be able to generate a cloud users audit report." -Color Yellow -ShowTime
			$ActiveDirectory = $False
		} else {
			Write-Color -Text "ERROR: The ActiveDirectory module could not be loaded and you are not connected to Microsoft Entra. At least one is required for this script to run." -Color Red -ShowTime
			Write-Color -Text "Exiting script..." -Color Red -ShowTime
			Exit 1
		}
	}

	# if a supported version of PS was found, check with the user if they want to install and use ImportExcel module for automatic formatting and/or use Entra to report for cloud users
	if ($SupportedPS) {
		Write-Color -Text "Checking for optional but recommended PowerShell modules" -ShowTime

		# Check for and prompt to install ImportExcel module
		$ImportExcel = Initialize-ImportExcel

		# Check for and prompt to install Microsoft.Graph module and connect to Entra ID
		$Entra,$PremiumEntraLicense,$AzUsers,$GlobalAdminMembers = Initialize-Entra
	}

	if ($ActiveDirectory) {
		# Get the domain name
		$DomainName = (Get-ADDomain).DNSRoot

		# Check if Disabled Users OU exists
		if (Get-ADOrganizationalUnit -Filter 'Name -like "*"' | Where-Object {$_.Name -like "*Disabled User*"}) {
			$DisabledOU = $True
		} Else {
			$DisabledOU = $False
		}

		# Get the Enterprise Admins group members
		$EnterpriseAdmins = Get-ADGroupMember -Identity "Enterprise Admins"

		# Get the Domain Admins group members
		$DomainAdmins = Get-ADGroupMember -Identity "Domain Admins"

		# Get all AD users with all of their properties
		$ADUsers = Get-ADUser -Filter * -Properties *

		# Get a list of all service accounts we can find on the domain
		$FoundServiceAccounts = Get-ADServiceAccounts -ADUsers $ADUsers

		if ($Entra) {
			# Get a list of all domains in Azure so we can use this to verify if synced users have a different UPN in the cloud compared to
			$AzDomains = Get-MgDomain
			# Process the AD users. if Entra is enabled then process on-prem AD users only and log hybrid users for processing with Merge-AzUsers.
			$ProcessedADUsers = Measure-ADUsers -ADUsers $ADUsers -AzUsers $AzUsers -Entra $Entra -ServiceAccounts $FoundServiceAccounts
		} else {
			# Process the AD users. if Entra is disabled then process on-prem AD users only.
			$ProcessedADUsers = Measure-ADUsers -ADUsers $ADUsers -Entra $Entra -ServiceAccounts $FoundServiceAccounts
		}
	# Initiaize variable and hashtable defaults if Active Directory if only generating a cloud user report
	} else {
		$DomainName = (Get-MgDomain | Where-Object {$_.IsDefault -eq $True}).Id
		$EnterpriseAdmins = @()
		$DomainAdmins = @()
		$ADUsers = @()
		$ProcessedADUsers = @()
	}

	if ($Entra) {
		# if Entra is enabled, process hybrid and cloud only users and merge LastLogonDate for hybrid users.
		$UserCollection = Merge-AzUsers -ADUsers $ADUsers -AzUsers $AzUsers -UserCollection $ProcessedADUsers -ServiceAccounts $FoundServiceAccounts
	} else {
		$UserCollection = $ProcessedADUsers
	}

	# Sort the user collection by Name. We have to sort before we export to Excel or csv if we want the table sorted a specific way.
	$SortedCollection = $UserCollection | Sort-Object -Property Name

	# Timestamp for Filename
	$TimeStamp = Get-Date -Format "MMddyyyy_HHmm"

	if ($ImportExcel) {
		# Format the file name with the domain name
		$FileName = "C:\Temp\$($domainName)_Users_Report_$TimeStamp.xlsx"

		# Export the sorted collection to a file and passthru to the XLSX variable to continue processing it
		$XLSX = $SortedCollection | Export-Excel $FileName -WorksheetName "AD Users" -AutoSize -FreezeTopRowFirstColumn -AutoFilter -BoldTopRow -PassThru
		# Select the AD Users worksheet we just saved to the workbook
		$Worksheet = $XLSX.Workbook.Worksheets["AD Users"]
		# Variable to represent the bottom of a column
		$lastCol = $Worksheet.Dimension.End.Row
		# Set the font size to 8 and autosize for the whole worksheet
		Set-ExcelRange -Worksheet $Worksheet -Range "A:Z" -FontSize 8 -AutoSize -BorderAround Thin

		if ($Entra) {
			#foreach ($User in $SortedCollection) {
			#	
			#}

			# Center align rows that will have "N/A" for a cleaner look and configured number formatting for date ranges
			# AccountExpired
			Set-ExcelRange -Worksheet $Worksheet -Range "H:H" -HorizontalAlignment Center -NumberFormat "MM/dd/yyyy hh:mm AM/PM"
			Set-ExcelRange -Worksheet $Worksheet -Range "H1:H1" -HorizontalAlignment Left
			Set-ExcelRange -Worksheet $Worksheet -Range "H:H" -Width 16
			# Domain Admin
			Set-ExcelRange -Worksheet $Worksheet -Range "J:J" -HorizontalAlignment Center 
			Set-ExcelRange -Worksheet $Worksheet -Range "J1:J1" -HorizontalAlignment Left
			# Global Admin
			Set-ExcelRange -Worksheet $Worksheet -Range "K:K" -HorizontalAlignment Center
			Set-ExcelRange -Worksheet $Worksheet -Range "K1:K1" -HorizontalAlignment Left
			# KnownServiceAccount
			Set-ExcelRange -Worksheet $Worksheet -Range "L:L" -HorizontalAlignment Center
			Set-ExcelRange -Worksheet $Worksheet -Range "L1:L1" -HorizontalAlignment Left
			# PasswordLastSet
			Set-ExcelRange -Worksheet $Worksheet -Range "M:M" -Width 16
			Set-ExcelRange -Worksheet $Worksheet -Range "M:M" -HorizontalAlignment Center -NumberFormat "MM/dd/yyyy hh:mm AM/PM"
			Set-ExcelRange -Worksheet $Worksheet -Range "M1:M1" -HorizontalAlignment Left
			# LastLogonDate
			Set-ExcelRange -Worksheet $Worksheet -Range "N:N" -Width 16
			Set-ExcelRange -Worksheet $Worksheet -Range "N:N" -HorizontalAlignment Center -NumberFormat "MM/dd/yyyy hh:mm AM/PM"
			Set-ExcelRange -Worksheet $Worksheet -Range "N1:N1" -HorizontalAlignment Left
			# Password Expired
			Set-ExcelRange -Worksheet $Worksheet -Range "P:P" -HorizontalAlignment Center
			Set-ExcelRange -Worksheet $Worksheet -Range "P1:P1" -HorizontalAlignment Left
			# Account Locked
			Set-ExcelRange -Worksheet $Worksheet -Range "Q:Q" -HorizontalAlignment Center
			Set-ExcelRange -Worksheet $Worksheet -Range "Q1:Q1" -HorizontalAlignment Left
			# CannotChangePassword
			Set-ExcelRange -Worksheet $Worksheet -Range "R:R" -HorizontalAlignment Center
			Set-ExcelRange -Worksheet $Worksheet -Range "R1:R1" -HorizontalAlignment Left
			# DateCreated
			Set-ExcelRange -Worksheet $Worksheet -Range "S:S" -Width 16
			Set-ExcelRange -Worksheet $Worksheet -Range "S:S" -HorizontalAlignment Center -NumberFormat "MM/dd/yyyy hh:mm AM/PM"
			Set-ExcelRange -Worksheet $Worksheet -Range "S1:S1" -HorizontalAlignment Left

			# Add conditional formatting to the data in the columns
			# Enabled Column
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "G2:G$lastCol" -RuleType Equal -ConditionValue $False -BackgroundColor Yellow
			# AccountExpired
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "H2:H$lastCol" -RuleType Equal -ConditionValue "N/A" -StopifTrue
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "H2:H$lastCol" -RuleType NotContainsBlanks -BackgroundColor Yellow
			# Enterprise Admin
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "I2:I$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
			# Domain Admin
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "J2:J$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
			# Global Admin
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "K2:K$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
			# Known Service Account
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "L2:L$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
			# PasswordLastSet
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "M2:M$lastCol" -RuleType Expression -ConditionValue "=`$N2<=(TODAY()-90)" -BackgroundColor Yellow -Bold
			# LastLogonDate
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "N2:N$lastCol" -RuleType Expression -ConditionValue "=`$N2<=(TODAY()-180)" -BackgroundColor Red -Bold
			# LastLogonDate
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "N2:N$lastCol" -RuleType Expression -ConditionValue "=`=AND(`$N2 > TODAY()-180, `$N2 < TODAY()-90)" -BackgroundColor Yellow
			# LastLogonDate
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "N2:N$lastCol" -RuleType Expression -ConditionValue "=`$N2>=(TODAY()-90)" -BackgroundColor LightGreen
			# PasswordNeverExpires
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "O2:O$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Red -Bold
			# PasswordExpired
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "P2:P$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
			# Account Locked
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "Q2:Q$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
			# CannotChangePassword
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "R2:R$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
		} else {
			# Center align rows that will have "N/A" for a cleaner look and configured number formatting for date ranges
			# AccountExpired
			Set-ExcelRange -Worksheet $Worksheet -Range "G:G" -HorizontalAlignment Center -NumberFormat "MM/dd/yyyy hh:mm AM/PM"
			Set-ExcelRange -Worksheet $Worksheet -Range "G1:G1" -HorizontalAlignment Left
			Set-ExcelRange -Worksheet $Worksheet -Range "G:G" -Width 16
			# Domain Admin
			Set-ExcelRange -Worksheet $Worksheet -Range "I:I" -HorizontalAlignment Center
			Set-ExcelRange -Worksheet $Worksheet -Range "I1:I1" -HorizontalAlignment Left
			# KnownServiceAccount
			Set-ExcelRange -Worksheet $Worksheet -Range "J:J" -HorizontalAlignment Center
			Set-ExcelRange -Worksheet $Worksheet -Range "J1:J1" -HorizontalAlignment Left
			# PasswordLastSet
			Set-ExcelRange -Worksheet $Worksheet -Range "K:K" -HorizontalAlignment Center -NumberFormat "MM/dd/yyyy hh:mm AM/PM"
			Set-ExcelRange -Worksheet $Worksheet -Range "K1:K1" -HorizontalAlignment Left
			Set-ExcelRange -Worksheet $Worksheet -Range "K:K" -Width 16
			# LastLogonDate
			Set-ExcelRange -Worksheet $Worksheet -Range "L:L" -HorizontalAlignment Center -NumberFormat "MM/dd/yyyy hh:mm AM/PM"
			Set-ExcelRange -Worksheet $Worksheet -Range "L1:L1" -HorizontalAlignment Left
			Set-ExcelRange -Worksheet $Worksheet -Range "L:L" -Width 16
			# Password Expired
			Set-ExcelRange -Worksheet $Worksheet -Range "N:N" -HorizontalAlignment Center
			Set-ExcelRange -Worksheet $Worksheet -Range "N1:N1" -HorizontalAlignment Left
			# Account Locked
			Set-ExcelRange -Worksheet $Worksheet -Range "O:O" -HorizontalAlignment Center
			Set-ExcelRange -Worksheet $Worksheet -Range "O1:O1" -HorizontalAlignment Left
			# CannotChangePassword
			Set-ExcelRange -Worksheet $Worksheet -Range "P:P" -HorizontalAlignment Center
			Set-ExcelRange -Worksheet $Worksheet -Range "P1:P1" -HorizontalAlignment Left
			# DateCreated
			Set-ExcelRange -Worksheet $Worksheet -Range "Q:Q" -HorizontalAlignment Center -NumberFormat "MM/dd/yyyy hh:mm AM/PM"
			Set-ExcelRange -Worksheet $Worksheet -Range "Q1:Q1" -HorizontalAlignment Left
			Set-ExcelRange -Worksheet $Worksheet -Range "Q:Q" -Width 16

			# Add conditional formatting to the data in the columns
			# Enabled Column
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "F2:F$lastCol" -RuleType Equal -ConditionValue $False -BackgroundColor Yellow
			# AccountExpired
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "G2:G$lastCol" -RuleType NotContainsBlanks -BackgroundColor Yellow
			# Enterprise Admin
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "H2:H$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
			# DomainAdmin
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "I2:I$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
			# Known Service Account
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "J2:J$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
			# PasswordLastSet
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "K2:K$lastCol" -RuleType Expression -ConditionValue "=`$K2<=(TODAY()-90)" -BackgroundColor Red -Bold
			# LastLogonDate
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "L2:L$lastCol" -RuleType Expression -ConditionValue "=`$L2<=(TODAY()-180)" -BackgroundColor Red -Bold
			# LastLogonDate
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "L2:L$lastCol" -RuleType Expression -ConditionValue "=`=AND(`$L2 > TODAY()-180, `$L2 < TODAY()-90)" -BackgroundColor Yellow
			# LastLogonDate
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "L2:L$lastCol" -RuleType Expression -ConditionValue "=`$L2>=(TODAY()-90)" -BackgroundColor LightGreen
			# PasswordNeverExpires
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "M2:M$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Red -Bold
			# PasswordExpired
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "N2:N$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
			# Account Locked
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "O2:O$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
			# CannotChangePassword
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "P2:P$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
		}
		Close-ExcelPackage $XLSX
	} else {
		# Format the file name with the domain name
		$FileName = "C:\Temp\$($domainName)_Users_Report_$TimeStamp.csv"
		$SortedCollection | Export-Csv -Path $FileName -NoTypeInformation
	}

	# Write comments to the PowerShell session
	Write-Color -Text "Report successfully saved to: ","$FileName" -Color Green,White -ShowTime
	If ($ImportExcel -or $Entra) {
		Write-Color -Text "You can remove the installed modules by closing this PowerShell session and running the below commands in a new PowerShell Admin session:" -Color DarkYellow -ShowTime -LinesBefore 1
		If ($ImportExcel) {
			Write-Color -Text "Uninstall-Module -Name 'ImportExcel' -Force" -Color DarkYellow -ShowTime
		}
		If ($Entra) {
			Write-Color -Text "Uninstall-Module -Name 'Microsoft.Graph.Authentication' -Force" -Color DarkYellow -ShowTime
			Write-Color -Text "Uninstall-Module -Name 'Microsoft.Graph.Users' -Force" -Color DarkYellow -ShowTime
			Write-Color -Text "Uninstall-Module -Name 'Microsoft.Graph.DirectoryObjects' -Force" -Color DarkYellow -ShowTime
			Write-Color -Text "Uninstall-Module -Name 'Microsoft.Graph.Identity.DirectoryManagement' -Force" -Color DarkYellow -ShowTime
		}
	}
	Write-Color -Text "Stay classy, Aunalytics" -Color Cyan -HorizontalCenter $True -LinesBefore 1
} catch {
	Write-Color -Text "Err Line: ","$($_.InvocationInfo.ScriptLineNumber)","Err Name: ","$($_.Exception.GetType().FullName) ","Err Msg: ","$($_.Exception.Message)" -Color Red,Magenta,Red,Magenta,Red,Magenta -ShowTime
} finally {
	Write-Host -NoNewLine 'Press any key to exit...';
	$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
	Exit 0
}
