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
		If Entra connection was successful, it processes Entra ID users
	Generate and save a report
		If ImportExcel installed the report will be saved as an xlsx file with automatic sizing and conditional formatting
		If ImportExcel is not installed, report will be saved to a csv file. Formatting will have to be performed manually.
  
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
#                                                 Globals                                                    #
##############################################################################################################

# Check if powershell is running in an admin session
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal ([Security.Principal.WindowsIdentity]::GetCurrent())
$AdminSession = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# Initialize default booleans for cleanup at script exit
$UntrustPSGallery = $False
$RemovePSGallery = $False
$RemoveNuGet = $False

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
    Color of the text. Accepts an array of colors. If more than one color is specified it will loop through colors for each string.
    If there are more strings than colors it will start from the beginning.
    Available colors are: Black, DarkBlue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow, Gray, DarkGray, DarkBlue, Green, Cyan, Red, Magenta, Yellow, White

    .PARAMETER BackGroundColor
    Color of the background. Accepts an array of colors. If more than one color is specified it will loop through colors for each string.
    If there are more strings than colors it will start from the beginning.
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
    Path to log file. If not specified no log file will be created.

    .PARAMETER DateTimeFormat
    Custom date and time format string. Default is yyyy-MM-dd HH:mm:ss

    .PARAMETER LogTime
    If set to $true it will add time to log file. Default is $true.

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
					"[$([datetime]::Now.ToString($DateTimeFormat))] $TextToFile" | Out-File -FilePath $LogFile -Encoding $Encoding -Append -ErrorAction Stop -WhatIf:$false
				}
				else {
					"$TextToFile" | Out-File -FilePath $LogFile -Encoding $Encoding -Append -ErrorAction Stop -WhatIf:$false
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
	# Initialize variably to remove ImportExcel module at end of script as Null
	$RemoveImportExcel = $Null

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
			Write-Color -Text "If authorized to install modules on this system",", would you like to temporarily install it for this script? ","(Y/N): " -Color Red,White,Yellow -NoNewline -ShowTime; $InstallImportExcel = Read-Host
			$InstallImportExcel = Read-Host ''

			switch ($InstallImportExcel) {
				"Y" {
					try {
						if ((Get-PSRepository).Name -contains "PSGallery") {
							if ((Get-PSRepository | Where-Object { $_.Name -eq 'PSGallery' }).InstallationPolicy -eq 'Untrusted') {
								Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
								$UntrustPSGallery = $True
							} else {
								$UntrustPSGallery = $False
							}
						} else {
							Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
							$RemovePSGallery = $True
						}

						if ((Get-PackageProvider).Name -notcontains 'NuGet') {
							Install-PackageProvider -Name NuGet -Force
							$RemoveNuGet = $True
						} else {
							$RemoveNuGet = $False
						}
						Write-Color -Text "Installing the ImportExcel module. Please be patient..." -ShowTime
						Install-Module -Name 'ImportExcel' -Force
						Write-Color -Text "ImportExcel module installed successfully."," It will be removed when this script exits." -Color Green,Yellow -ShowTime
						$ImportExcel = $True
						$RemoveImportExcel = $True
					} catch {
						Write-Color -Text "ERROR: ImportExcel module failed to install. See the error below. The report will output to CSV only until the error is corrected." -Color Red -ShowTime
						Write-Color -Text "Err Line: ","$($_.InvocationInfo.ScriptLineNumber)","Err Name: ","$($_.Exception.GetType().FullName) ","Err Msg: ","$($_.Exception.Message)" -Color Red,Magenta,Red,Magenta,Red,Magenta -ShowTime
						$ImportExcel = $False
						$RemoveImportExcel = $True
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
			Write-Color -Text "NOTICE: If authorized to install PowerShell modules on this system you can run this script in an admin session to install the ImportExcel module and save directly to xlsx with automated formatting" -Color Yellow -ShowTime
		}
	}

	return $ImportExcel,$RemoveImportExcel,$UntrustPSGallery,$RemovePSGallery,$RemoveNuGet
}

function Initialize-Entra {
	<#
    .DESCRIPTION
    Check if the user wants to connect to Entra ID and process cloud users.
    If the Microsoft.Graph module is not installed it will prompt the user if they want to install it.
    If the PSRepository or PackageProvider are modified or the module is installed, it will be removed at the end of the script.
    If a connection to Entra ID is successful then it grabs all Entra ID users and their relevant properties and a list of all global admins.

    .PARAMETER [boolean]RemoveGraphAPI
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
    [Returning all the output variables without inputting any variables for initial first run of the function]
    $Entra, $PremiumEntraLicense, $AzUsers, $GlobalAdminMembers, $RemoveGraphAPI, $UntrustPSGallery, $RemovePSGallery, $RemoveNuGet = Initialize-Entra
    
    [To look the function until graph API connection or user cancels]
    Initialize-Entra -RemoveGraphAPI $RemoveGraphAPI -UntrustPSGallery $UntrustPSGallery -RemovePSGallery $RemovePSGallery -RemoveNuGet $RemoveNuGet
    #>

	param(
		[boolean]$RemoveGraphAPI = $Null,
		[boolean]$UntrustPSGallery,
		[boolean]$RemovePSGallery,
		[boolean]$RemoveNuGet
	)
	Write-Color -Text "Would you like to connect to Entra ID? ","(Y/N): " -Color White,Yellow -NoNewline -ShowTime; $EntraID = Read-Host
	switch ($EntraID) {
		'Y' {
			if (Get-Module -ListAvailable -Name 'Microsoft.Graph') {
				Write-Color -Text "Microsoft.Graph module detected. Connecting to Graph API..." -Color Green -ShowTime

				# Import the ImportExcel module and set the $ImportExcel variable to True
				Import-Module Microsoft.Graph.Authentication
				Import-Module Microsoft.Graph.Users
				Import-Module Microsoft.Graph.DirectoryObjects
				Import-Module Microsoft.Graph.Identity.DirectoryManagement

				$GraphAPI = $True
				if ($Null -ne $RemoveGraphAPI) {
					$RemoveGraphAPI = $False
				}
			} else {
				# Check if we are running in an admin session. Otherwise skip trying to install the module and throw a warning to console.
				if ($AdminSession) {
					# Graph API module is not installed. Ask if allowed to install and user wants to install it.
					Write-Color -Text 'WARNING: Graph API module is not installed. The report will display on-premises AD Users only.' -Color Yellow -ShowTime
					Write-Color -Text "If authorized to install modules on this system",", would you like to temporarily install it for this script? ","(Y/N): " -Color Red,White,Yellow -NoNewline -ShowTime; $InstallGraph = Read-Host

					switch ($InstallGraph) {
						"Y" {
							try {
								if ((Get-PSRepository).Name -contains "PSGallery") {
									if ((Get-PSRepository | Where-Object { $_.Name -eq 'PSGallery' }).InstallationPolicy -eq 'Untrusted') {
										Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
										$UntrustPSGallery = $True
									} else {
										$UntrustPSGallery = $False
									}
								} else {
									Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
									$RemovePSGallery = $True
								}

								if ((Get-PackageProvider).Name -notcontains "NuGet") {
									Install-PackageProvider -Name NuGet -Force
									$RemoveNuGet = $True
								} else {
									$RemoveNuGet = $False
								}
								Install-Module -Name 'Microsoft.Graph' -Force
								Write-Color -Text "Microsoft.Graph module installed successfully. ","It will be removed when this script exits." -Color Green,Yellow -ShowTime
								Import-Module Microsoft.Graph.Authentication
								Import-Module Microsoft.Graph.Users
								Import-Module Microsoft.Graph.DirectoryObjects
								Import-Module Microsoft.Graph.Identity.DirectoryManagement

								$GraphAPI = $True
								$RemoveGraphAPI = $True
							} catch {
								Write-Color -Text "ERROR: Microsoft.Graph module failed to install. See the error below. The report will output to CSV only until the error is corrected." -Color Red -ShowTime
								Write-Color -Text "Err Line: ","$($_.InvocationInfo.ScriptLineNumber)","Err Name: ","$($_.Exception.GetType().FullName) ","Err Msg: ","$($_.Exception.Message)" -Color Red,Magenta,Red,Magenta,Red,Magenta -ShowTime
								$GraphAPI = $False
								$RemoveGraphAPI = $True
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
					Write-Color -Text "NOTICE: If authorized to install PowerShell modules on this system you can run this script in an admin session to install the Graph API module and run the audit against cloud users and combine cloud properties with on-prem properties." -Color Yellow -ShowTime
				}
			}

			# If Microsoft.Graph modules were successfully installed
			if ($GraphAPI) {
				try {
					# Interactive login to the tenant requesting the required permissions only
					Connect-MgGraph -Scopes 'Directory.Read.All, User.Read.All, AuditLog.Read.All' -NoWelcome -ErrorAction Stop
					try {
						# Try to get all users including SignInActivity which is only available with a premium license
						$AzUsers = Get-MgUser -All -Property Id,UserPrincipalName,SignInActivity,OnPremisesSyncEnabled,displayName,samAccountName,AccountEnabled,mail,lastPasswordChangeDateTime,PasswordPolicies,CreatedDateTime -ErrorAction Stop
						$PremiumEntraLicense = $True
					} catch {
						# If the tenant doesnt have a premium license get all users without including SignInActivity
						if ($_.Exception.Message -like "*Neither tenant is B2C or tenant doesn't have premium license*") {
							Write-Color -Text "WARNING: This tenant does not have a premium license. LastLogonDate will show on-premises AD datetimes only!" -Color Yellow -ShowTime
							$AzUsers = Get-MgUser -All -Property Id,UserPrincipalName,OnPremisesSyncEnabled,displayName,samAccountName,AccountEnabled,mail,lastPasswordChangeDateTime,PasswordPolicies,CreatedDateTime -ErrorAction Stop
							$PremiumEntraLicense = $False
						}
					}

					$GlobalAdminRoleId = Get-MgDirectoryRole | Where-Object { $_.DisplayName -eq "Global Administrator" } | Select-Object -ExpandProperty ID
					$GlobalAdminMembers = Get-MgDirectoryRoleMemberAsUser -DirectoryRoleId $GlobalAdminRoleId
					$Entra = $True
				} catch {
					Write-Color -Text "ERROR: Connection to Graph API failed!" -Color Red -ShowTime
					Write-Color -Text "Err Line: ","$($_.InvocationInfo.ScriptLineNumber)","Err Name: ","$($_.Exception.GetType().FullName) ","Err Msg: ","$($_.Exception.Message)" -Color Red,Magenta,Red,Magenta,Red,Magenta -ShowTime
					Write-Color -Text "Would you like to try connecting to the Graph API again? ","(Y/N): " -Color White,Yellow -NoNewline -ShowTime; $TryAgain = Read-Host
					switch ($TryAgain) {
						"Y" {
							Initialize-Entra -RemoveGraphAPI $RemoveGraphAPI -UntrustPSGallery $UntrustPSGallery -RemovePSGallery $RemovePSGallery -RemoveNuGet $RemoveNuGet
						}
						"N" {
							Write-Color -Text "Graph API module will not be used. ","Report will show on-premises AD users only." -Color White,Yellow -ShowTime
							$Entra = $False
							$PremiumEntraLicense = $False
							$AzUsers = $Null
							$GlobalAdminMembers = $Null
							if ($RemoveGraphAPI) {
								Remove-Module -Name 'Microsoft.Graph.Users' -Force
								Remove-Module -Name 'Microsoft.Graph.DirectoryObjects' -Force
								Remove-Module -Name 'Microsoft.Graph.Identity.DirectoryManagement' -Force
								Remove-Module -Name 'Microsoft.Graph.Authentication' -Force
								Uninstall-Module -Name 'Microsoft.Graph' -Force
							}
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

	return $Entra,$PremiumEntraLicense,$AzUsers,$GlobalAdminMembers,$RemoveGraphAPI,$UntrustPSGallery,$RemovePSGallery,$RemoveNuGet
}

function Measure-ADUsers {
	<#
    .DESCRIPTION
    Processes the active directory domain user accounts.
    If a connection to Entra was established it will only process on on-premises AD user accounts only. Hybrid and cloud users will be processed later.

    .PARAMETER [object]ADUsers
    Collection of all AD Users and all of their properties

    .PARAMETER [object]AzUsers
    Collection of all Entra ID Users and their relevent properties

    .PARAMETER [boolean]AzUsersToProcess
    Array of UserPrincipalNames of all the Cloud users that need to processed by this function

    .EXAMPLE
    Measure-ADUsers -ADUsers $ADUsers -AzUsers $AzUsers -Entra $True
    Measure-ADUsers -ADUsers $ADUsers -AzUsers $AzUsers -Entra $False
    Measure-ADUsers -ADUsers $ADUsers -AzUsers $AzUsers -Entra $Entra
    #>

	param(
		[Parameter(Mandatory = $True)] $ADUsers,
		[Parameter(Mandatory = $True)] $AzUsers,
		[Parameter(Mandatory = $True)] [boolean]$Entra
	)

	# Initialize arrays for UserCollection and AzUsersToProcess
	$UserCollection = @()
	$AzUsersToProcess = @()

	# Initialize user counter for progress bar
	$Count = 1

	if ($ADUsers.Count -gt 0) {
		# Process each user account found in active directory
		foreach ($User in $ADUsers) {
			Write-Color -Text "Processing Active Directory Users" -ShowTime
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

			# If the account has an account expiration date then consider it expired.
			if ($Null -ne $User.AccountExpirationDate) {
				$AccountExpired = $User.AccountExpirationDate
			} else {
				$AccountExpired = $Null
			}

			# If email property is blank then set to a blank space for formatting the spreadsheet. This stops a previous column from displaying over it.
			if ($Null -eq $User.mail) {
				$Mail = " "
			} else {
				$Mail = $User.mail
			}

			# If an Entra connection was successful then process on-prem only users and write the rest to an array to process later
			if ($Entra) {
				# On-prem user without synced cloud user
				if (($AzUsers).UserPrincipalName -notcontains $User.UserPrincipalName) {
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
						"AzGlobalAdmin" = "N/A"
						PasswordLastSet = $User.PasswordLastSet
						LastLogonDate = $User.LastLogonDate
						PasswordNeverExpires = $User.PasswordNeverExpires
						PasswordExpired = $User.PasswordExpired
						"Account Locked" = $User.lockedOut
						CannotChangePassword = $User.CannotChangePassword
						"Date Created" = $User.whenCreated
						Notes = ""
						Action = ""
						"Follow Up" = ""
						Resolution = ""
					}
					# Otherwise add the user to array AzUsersToProcess to be processed by Merge-AzUsers function
				} else {
					$AzUsersToProcess += $User.UserPrincipalName
				}
				# No connection to Entra ID. Process active directory users only
			} else {
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
					PasswordLastSet = $User.PasswordLastSet
					LastLogonDate = $User.LastLogonDate
					PasswordNeverExpires = $User.PasswordNeverExpires
					PasswordExpired = $User.PasswordExpired
					"Account Locked" = $User.lockedOut
					CannotChangePassword = $User.CannotChangePassword
					"Date Created" = $User.whenCreated
					Notes = ""
					Action = ""
					"Follow Up" = ""
					Resolution = ""
				}
			}

			$Count += 1
		}
	}

	return $UserCollection, $AzUsersToProcess
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

    .PARAMETER [array]AzUsersToProcess
    Array of UserPrincipalNames of all the Cloud users that need to processed by this function

    .PARAMETER [object]UserCollection
    Collection of all users that have already been processed

    .EXAMPLE
    Merge-AzUsers $ADUsers $AzUsers $AzUsersToProcess $UserCollection
    #>

	param(
		[Parameter(Mandatory = $True)] $ADUsers,
		[Parameter(Mandatory = $True)] $AzUsers,
		[Parameter(Mandatory = $True)] $AzUsersToProcess,
		[Parameter(Mandatory = $True)] $UserCollection
	)

	# Initialize user counter for progress bar
	$Count = 1

	foreach ($AzUser in $AzUsers) {
		Write-Color -Text "Processing Entra ID Users" -ShowTime
		Write-Progress -Id 1 -Activity "Processing Entra Users" -Status "Current Count: ($Count/$($AzUsers.Count))" -PercentComplete (($Count / $AzUsers.Count) * 100) -CurrentOperation "Processing... $($AzUser.DisplayName)"

		# On-Prem user with synced cloud user
		if ($AzUsersToProcess -contains $AzUser.UserPrincipalName) {
			$User = $ADUsers | Where-Object { $_.UserPrincipalName -eq $AzUser.UserPrincipalName }

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

			# If the account has an account expiration date then consider it expired.
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

			# If email property is blank then set to a blank space for formatting the spreadsheet. This stops a previous column from displaying over it.
			if ($Null -eq $User.mail) {
				$Mail = " "
			} else {
				$Mail = $User.mail
			}

			# If the tenant has a premium license then get and compare the last sign-in timestamp and lastLogonDate timestamp
			if ($PremiumEntraLicense) {
				if ($AzUser.signInActivity.lastSignInDateTime) {
					$AzlastLogonDate = [datetime]$AzUser.signInActivity.lastSignInDateTime
					# If the last sign-in timestamp is newer than set that as lastLogonDate property
					if ($User.LastLogonDate -lt $AzlastLogonDate) {
						$LastLogonDate = $AzlastLogonDate
						# Otherwise use the active directory lastLogonDate timestamp
					} else {
						$LastLogonDate = $User.LastLogonDate
					}
					# If there is no last sign-in timestamp then default to AD lastLogonDate timestamp.
				} else {
					$LastLogonDate = $User.LastLogonDate
				}
				# If the tenant doesnt have a premium license then we cant get last sign-in timestamp. Default to AD lastLogonDate timestamp.
			} else {
				$LastLogonDate = $User.LastLogonDate
			}

			# Add the user to the UserCollection
			$UserCollection += [pscustomobject]@{
				"Name" = $User.DisplayName
				SamAccountName = $User.SamAccountName
				UserPrincipalName = $User.UserPrincipalName
				"Email Address" = $Mail
				"User Type" = "Hybrid"
				Enabled = $User.Enabled
				AccountExpiredDate = $AccountExpired
				EnterpriseAdmin = $EnterpriseAdmin
				DomainAdmin = $DomainAdmin
				"AzGlobalAdmin" = $GlobalAdmin
				PasswordLastSet = $User.PasswordLastSet
				LastLogonDate = $LastLogonDate
				PasswordNeverExpires = $User.PasswordNeverExpires
				PasswordExpired = $User.PasswordExpired
				"Account Locked" = $User.lockedOut
				CannotChangePassword = $User.CannotChangePassword
				"Date Created" = $User.whenCreated
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

			# If the tenant has a premium license then grab the last sign-in timestamp
			if ($PremiumEntraLicense) {
				if ($AzUser.signInActivity.lastSignInDateTime) {
					$LastLogonDate = [datetime]$AzUser.signInActivity.lastSignInDateTime
				} else {
					$LastLogonDate = $Null
				}
			} else {
				$LastLogonDate = $Null
			}

			# If string found in PasswordPolicies then the password is set to never expire
			if ($AzUser.PasswordPolicies -contains "DisablePasswordExpiration") {
				$PasswordNeverExpires = $True
			} else {
				$PasswordNeverExpires = $False
			}

			# If email property is blank then set to a blank space for formatting the spreadsheet. This stops a previous column from displaying over it.
			if ($Null -eq $AzUser.mail) {
				$Mail = " "
			} else {
				$Mail = $AzUser.mail
			}

			# Add the user to the UserCollection
			$UserCollection += [pscustomobject]@{
				"Name" = $AzUser.DisplayName
				SamAccountName = "N/A"
				UserPrincipalName = $AzUser.UserPrincipalName
				"Email Address" = $Mail
				"User Type" = "Cloud"
				Enabled = $AzUser.AccountEnabled
				AccountExpiredDate = "N/A"
				EnterpriseAdmin = $False
				DomainAdmin = $False
				"AzGlobalAdmin" = $GlobalAdmin
				PasswordLastSet = $AzUser.lastPasswordChangeDateTime
				LastLogonDate = $LastLogonDate
				PasswordNeverExpires = $PasswordNeverExpires
				PasswordExpired = "N/A"
				"Account Locked" = "N/A"
				CannotChangePassword = "N/A"
				"Date Created" = $AzUser.CreatedDateTime
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
try {
	Clear-Host
	Write-Color -Text "__________________________________________________________________________________________" -Color White -BackgroundColor Black -HorizontalCenter $True -VerticalCenter $True
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
	Write-Color -Text "Script:","User Audit Report" -Color Yellow,White -BackgroundColor Black -LinesBefore 1
	Write-Color -Text "Checking for optional but recommended PowerShell modules" -ShowTime

	# Check for and prompt to install ImportExcel module
	$ImportExcel,$RemoveImportExcel,$IEUntrustPSGallery,$IERemovePSGallery,$IERemoveNuGet = Initialize-ImportExcel

	# Check for and prompt to install Microsoft.Graph module and connect to Entra ID
	$Entra,$PremiumEntraLicense,$AzUsers,$GlobalAdminMembers,$RemoveGraphAPI,$MgUntrustPSGallery,$MgRemovePSGallery,$MgRemoveNuGet = Initialize-Entra

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

	# If either value is true, set the PSGallery back to untrusted at script exit
	if ($IEUntrustPSGallery -or $MgUntrustPSGallery) {
		$UntrustPSGallery = $True
	}

	# If either value is true, set the PSGallery to be removed at script exit
	if ($IERemovePSGallery -or $MgRemovePSGallery) {
		$RemovePSGallery = $True
	}

	# If either value is true, remove the NuGet package manager at script exit
	if ($IERemoveNuGet -or $MgRemoveNuGet) {
		$RemoveNuGet = $True
	}

	if ($ActiveDirectory) {
		# Get the domain name
		$DomainName = (Get-ADDomain).DNSRoot

		# Get the Enterprise Admins group members
		$EnterpriseAdmins = Get-ADGroupMember -Identity "Enterprise Admins"

		# Get the Domain Admins group members
		$DomainAdmins = Get-ADGroupMember -Identity "Domain Admins"

		# Get all AD users with all of their properties
		$ADUsers = Get-ADUser -Filter * -Properties *

		# Process the AD users. If Entra is enabled then process on-prem AD users only and log hybrid users for processing with Merge-AzUsers.
		$ProcessedADUsers, $AzUsersToProcess = Measure-ADUsers -ADUsers $ADUsers -AzUsers $AzUsers -Entra $Entra
	} else {
		$DomainName = (Get-MgDomain | Where-Object {$_.IsDefault -eq $True}).Id
		$EnterpriseAdmins = @()
		$DomainAdmins = @()
		$ADUsers = @()
		$ProcessedADUsers = @()
		$AzUsersToProcess = @()
	}

	# If Entra is enabled, process hybrid and cloud only users and merge LastLogonDate for hybrid users.
	$UserCollection = Merge-AzUsers -ADUsers $ADUsers -AzUsers $AzUsers -AzUsersToProcess $AzUsersToProcess -UserCollection $ProcessedADUsers

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
			# Center align rows that will have "N/A" for a cleaner look
			Set-ExcelRange -Worksheet $Worksheet -Range "G:G" -HorizontalAlignment Center -NumberFormat "MM/dd/yyyy hh:mm AM/PM"
			Set-ExcelRange -Worksheet $Worksheet -Range "J:J" -HorizontalAlignment Center
			Set-ExcelRange -Worksheet $Worksheet -Range "N:N" -HorizontalAlignment Center
			Set-ExcelRange -Worksheet $Worksheet -Range "O:O" -HorizontalAlignment Center
			Set-ExcelRange -Worksheet $Worksheet -Range "P:P" -HorizontalAlignment Center

			# Add conditional formatting to the data in the columns
			# Enabled Column
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "F2:F$lastCol" -RuleType Equal -ConditionValue $False -BackgroundColor Yellow
			# AccountExpired
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "G2:G$lastCol" -RuleType NotContainsBlanks -BackgroundColor Yellow
			# Enterprise Admin
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "H2:H$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
			# Domain Admin
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "I2:I$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
			# Global Admin
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "J2:J$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
			# PasswordLastSet
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "K2:K$lastCol" -RuleType Expression -ConditionValue "=`$K2<=(TODAY()-90)" -BackgroundColor Yellow -Bold
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
		} else {
			# Add conditional formatting to the data in the columns
			# Enabled Column
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "F2:F$lastCol" -RuleType Equal -ConditionValue $False -BackgroundColor Yellow
			# AccountExpired
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "G2:G$lastCol" -RuleType NotContainsBlanks -BackgroundColor Yellow
			# Enterprise Admin
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "H2:H$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
			# DomainAdmin
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "I2:I$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
			# PasswordLastSet
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "J2:J$lastCol" -RuleType Expression -ConditionValue "=`$J2<=(TODAY()-90)" -BackgroundColor Red -Bold
			# LastLogonDate
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "K2:K$lastCol" -RuleType Expression -ConditionValue "=`$K2<=(TODAY()-180)" -BackgroundColor Red -Bold
			# LastLogonDate
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "K2:K$lastCol" -RuleType Expression -ConditionValue "=`=AND(`$K2 > TODAY()-180, `$K2 < TODAY()-90)" -BackgroundColor Yellow
			# LastLogonDate
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "K2:K$lastCol" -RuleType Expression -ConditionValue "=`$K2>=(TODAY()-90)" -BackgroundColor LightGreen
			# PasswordNeverExpires
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "L2:L$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Red -Bold
			# PasswordExpired
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "M2:M$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
			# Account Locked
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "N2:N$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
			# CannotChangePassword
			Add-ConditionalFormatting -Worksheet $Worksheet -Address "O2:O$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
		}
		Close-ExcelPackage $XLSX
	} else {
		# Format the file name with the domain name
		$FileName = "C:\Temp\$($domainName)_Users_Report_$TimeStamp.csv"
		Export-Csv -Path $FileName -NoTypeInformation
	}

	Write-Color -Text "Report successfully saved to: ","$FileName" -Color Green,White -ShowTime
	Write-Color -Text "" -Color Blue -HorizontalCenter $True -LinesBefore 1
} catch {
	Write-Color -Text "Err Line: ","$($_.InvocationInfo.ScriptLineNumber)","Err Name: ","$($_.Exception.GetType().FullName) ","Err Msg: ","$($_.Exception.Message)" -Color Red,Magenta,Red,Magenta,Red,Magenta -ShowTime
} finally {
	# When the script exits revert all packageprovider and repository changes and remove installed modules if not previously installed
	try {
		if ($UntrustPSGallery) {
			Set-PSRepository -Name 'PSGallery' -InstallationPolicy Untrusted
		}

		if ($RemovePSGallery) {
			Unregister-PSRepository -Name 'PSGallery'
		}

		if ($RemoveImportExcel) {
			Remove-Module -Name 'ImportExcel' -Force
			Uninstall-Module -Name 'ImportExcel' -Force
		}

		if ($RemoveGraphAPI) {
			Remove-Module -Name 'Microsoft.Graph.Users' -Force
			Remove-Module -Name 'Microsoft.Graph.DirectoryObjects' -Force
			Remove-Module -Name 'Microsoft.Graph.Identity.DirectoryManagement' -Force
			Remove-Module -Name 'Microsoft.Graph.Authentication' -Force
			Uninstall-Module -Name 'Microsoft.Graph' -Force
		}

		if ($RemoveNuGet) {
			Uninstall-PackageProvider -Name NuGet -Force
		}
	} catch {
		Write-Color -Text "Err Line: ","$($_.InvocationInfo.ScriptLineNumber)","Err Name: ","$($_.Exception.GetType().FullName) ","Err Msg: ","$($_.Exception.Message)" -Color Red,Magenta,Red,Magenta,Red,Magenta -ShowTime
	}
}
