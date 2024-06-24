# AD User Audit Report

##############################################################################################################
#                                                 Globals                                                    #
##############################################################################################################

# Check if powershell is running in an admin session
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$AdminSession = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# Initialize default booleans for cleanup at script exit
$UntrustPSGallery = $False
$RemovePSGallery = $False
$RemoveNuGet = $False

##############################################################################################################
#                                                Functions                                                   #
##############################################################################################################

Function Write-Color {
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
    [alias('Write-Colour')]
    [CmdletBinding()]
    param (
        [alias ('T')] [String[]]$Text,
        [alias ('C', 'ForegroundColor', 'FGC')] [ConsoleColor[]]$Color = [ConsoleColor]::White,
        [alias ('B', 'BGC')] [ConsoleColor[]]$BackGroundColor = $null,
        [bool] $VerticalCenter = $False,
        [bool] $HorizontalCenter = $False,
        [alias ('Indent')][int] $StartTab = 0,
        [int] $LinesBefore = 0,
        [int] $LinesAfter = 0,
        [int] $StartSpaces = 0,
        [alias ('L')] [string] $LogFile = '',
        [Alias('DateFormat', 'TimeFormat')][string] $DateTimeFormat = 'yyyy-MM-dd HH:mm:ss',
        [alias ('LogTimeStamp')][bool] $LogTime = $true,
        [int] $LogRetry = 2,
        [ValidateSet('unknown', 'string', 'unicode', 'bigendianunicode', 'utf8', 'utf7', 'utf32', 'ascii', 'default', 'oem')][string]$Encoding = 'Unicode',
        [switch] $ShowTime,
        [switch] $NoNewLine,
        [alias('HideConsole')][switch] $NoConsoleOutput
    )
    if (-not $NoConsoleOutput) {
        $DefaultColor = $Color[0]
        if ($null -ne $BackGroundColor -and $BackGroundColor.Count -ne $Color.Count) {
            Write-Error "Colors, BackGroundColors parameters count doesn't match. Terminated."
            return
        }
        If ($VerticalCenter) {
            for ($i = 0; $i -lt ([Math]::Max(0, $Host.UI.RawUI.BufferSize.Height / 4)); $i++) {
                Write-Host -Object "`n" -NoNewline 
            } 
        } # Center the output vertically according to the powershell window size
        if ($LinesBefore -ne 0) {
            for ($i = 0; $i -lt $LinesBefore; $i++) {
                Write-Host -Object "`n" -NoNewline 
            } 
        } # Add empty line before
        If ($HorizontalCenter) {
            $MessageLength = 0
            ForEach ($Value in $Text) {
                $MessageLength += $Value.Length
            }
            Write-Host ("{0}" -f (' ' * ([Math]::Max(0, $Host.UI.RawUI.BufferSize.Width / 2) - [Math]::Floor($MessageLength / 2)))) -NoNewline 
        } # Center the line horizontally according to the powershell window size
        if ($StartTab -ne 0) {
            for ($i = 0; $i -lt $StartTab; $i++) {
                Write-Host -Object "`t" -NoNewline 
            } 
        }  # Add TABS before text
        
        if ($StartSpaces -ne 0) {
            for ($i = 0; $i -lt $StartSpaces; $i++) {
                Write-Host -Object ' ' -NoNewline 
            } 
        }  # Add SPACES before text
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
                    for ($i = 0; $i -lt $Color.Length ; $i++) {
                        Write-Host -Object $Text[$i] -ForegroundColor $Color[$i] -NoNewline 
                        
                    }
                    for ($i = $Color.Length; $i -lt $Text.Length; $i++) {
                        Write-Host -Object $Text[$i] -ForegroundColor $DefaultColor -NoNewline 
                        
                    }
                }
                else {
                    for ($i = 0; $i -lt $Color.Length ; $i++) {
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
        }  # Add empty line after
    }
    if ($Text.Count -and $LogFile) {
        # Save to file
        $TextToFile = ""
        for ($i = 0; $i -lt $Text.Length; $i++) {
            $TextToFile += $Text[$i]
        }
        $Saved = $false
        $Retry = 0
        Do {
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
        } Until ($Saved -eq $true -or $Retry -ge $LogRetry)
    }
}

Function Initialize-ImportExcel {
    # Initialize variably to remove ImportExcel module at end of script as Null
    $RemoveImportExcel = $Null

    # Check if ImportExcel module is installed
    If (Get-Module -ListAvailable -Name 'ImportExcel') {
        Write-Color -Text "ImportExcel module detected. Will save directly to XLSX with automated formatting..." -Color Green -ShowTime

        # Import the ImportExcel module and set the $ImportExcel variable to True
        Import-Module ImportExcel
        $ImportExcel = $True
        $RemoveImportExcel = $False
    } Else {
        If ($AdminSession) {
            # ImportExcel module is not installed. Ask if allowed to install and user wants to install it.
            Write-Color -Text 'WARNING: ImportExcel module is not installed. Without it the report will output in CSV and you will have to format it manually.' -Color Yellow -ShowTime
            Write-Color -Text "If authorized to install modules on this system",", would you like to temporarily install it for this script? ", "(Y/N): " -Color Red, White, Yellow -NoNewLine -ShowTime; $InstallImportExcel = Read-Host
            $InstallImportExcel = Read-Host ''

            Switch ($InstallImportExcel) {
                "Y" {
                    Try {
                        If ((Get-PSRepository).Name -contains "PSGallery") {
                            If ((Get-PSRepository | Where-Object {$_.Name -eq 'PSGallery'}).InstallationPolicy -eq 'Untrusted') {
                                Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
                                $UntrustPSGallery = $True
                            } Else {
                                $UntrustPSGallery = $False
                            }
                        } Else {
                            Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
                            $RemovePSGallery = $True
                        }
                        
                        If ((Get-PackageProvider).Name -notcontains 'NuGet') {
                            Install-PackageProvider -Name NuGet -Force
                            $RemoveNuGet = $True
                        } Else {
                            $RemoveNuGet = $False
                        }
                        Write-Color -Text "Installing the ImportExcel module. Please be patient..." -ShowTime
                        Install-Module -Name 'ImportExcel'-Force
                        Write-Color -Text "ImportExcel module installed successfully."," It will be removed when this script exits." -Color Green, Yellow -ShowTime
                        $ImportExcel = $True
                        $RemoveImportExcel = $True
                    } Catch {
                        Write-Color -Text "ERROR: ImportExcel module failed to install. See the error below. The report will output to CSV only until the error is corrected." -Color Red -ShowTime
                        Write-Color -Text "Err Line: ","$($_.InvocationInfo.ScriptLineNumber)","Err Name: ","$($_.Exception.GetType().FullName) ","Err Msg: ","$($_.Exception.Message)" -Color Red, Magenta, Red, Magenta, Red, Magenta -ShowTime
                        $ImportExcel = $False
                        $RemoveImportExcel = $True
                    }
                }
                "N" {
                    Write-Color -Text "ImportExcel module will not be installed. ","Proceeding to save to CSV format." -Color White, Yellow -ShowTime
                    $ImportExcel = $False
                }
                Default { 
                    Write-Color -Text "No option was selected. ","Proceeding to save to CSV format." -Color White, Yellow -ShowTime
                    $ImportExcel = $False
                }
            }
        } Else {
            Write-Color -Text "NOTICE: If authorized to install PowerShell modules on this system you can run this script in an admin session to install the ImportExcel module and save directly to xlsx with automated formatting" -Color Yellow -ShowTime
        }
    }

    Return $ImportExcel, $RemoveImportExcel, $UntrustPSGallery, $RemovePSGallery, $RemoveNuGet
}

Function Initialize-Entra {
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

    param (
        [boolean]$RemoveGraphAPI = $Null,
        [boolean]$UntrustPSGallery,
        [boolean]$RemovePSGallery,
        [boolean]$RemoveNuGet
    )
    Write-Color -Text "Would you like to connect to Entra ID? ","(Y/N): " -Color White, Yellow -NoNewLine -ShowTime; $EntraID = Read-Host
    Switch ($EntraID) {
        'Y' { 
            If (Get-Module -ListAvailable -Name 'Microsoft.Graph') {
                Write-Color -Text "Microsoft.Graph module detected. Connecting to Graph API..." -Color Green -ShowTime
        
                # Import the ImportExcel module and set the $ImportExcel variable to True
                Import-Module Microsoft.Graph.Authentication
                Import-Module Microsoft.Graph.Users
                Import-Module Microsoft.Graph.DirectoryObjects
                Import-Module Microsoft.Graph.Identity.DirectoryManagement

                $GraphAPI = $True
                If ($Null -ne $RemoveGraphAPI) {
                    $RemoveGraphAPI = $False
                }
            } Else {
                If ($AdminSession) {
                    # Graph API module is not installed. Ask if allowed to install and user wants to install it.
                    Write-Color -Text 'WARNING: Graph API module is not installed. The report will display on-premises AD Users only.' -Color Yellow -ShowTime
                    Write-Color -Text "If authorized to install modules on this system",", would you like to temporarily install it for this script? ", "(Y/N): " -Color Red, White, Yellow -NoNewLine -ShowTime; $InstallGraph = Read-Host
        
                    Switch ($InstallGraph) {
                        "Y" {
                            Try {
                                If ((Get-PSRepository).Name -contains "PSGallery") {
                                    If ((Get-PSRepository | Where-Object {$_.Name -eq 'PSGallery'}).InstallationPolicy -eq 'Untrusted') {
                                        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
                                        $UntrustPSGallery = $True
                                    } Else {
                                        $UntrustPSGallery = $False
                                    }
                                } Else {
                                    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
                                    $RemovePSGallery = $True
                                }

                                If ((Get-PackageProvider).Name -notcontains "NuGet") {
                                    Install-PackageProvider -Name NuGet -Force
                                    $RemoveNuGet = $True
                                } Else {
                                    $RemoveNuGet = $False
                                }
                                Install-Module -Name 'Microsoft.Graph'-Force
                                Write-Color -Text "Microsoft.Graph module installed successfully. ","It will be removed when this script exits." -Color Green, Yellow -ShowTime
                                Import-Module Microsoft.Graph.Authentication
                                Import-Module Microsoft.Graph.Users
                                Import-Module Microsoft.Graph.DirectoryObjects
                                Import-Module Microsoft.Graph.Identity.DirectoryManagement

                                $GraphAPI = $True
                                $RemoveGraphAPI = $True
                            } Catch {
                                Write-Color -Text "ERROR: Microsoft.Graph module failed to install. See the error below. The report will output to CSV only until the error is corrected." -Color Red -ShowTime
                                Write-Color -Text "Err Line: ","$($_.InvocationInfo.ScriptLineNumber)","Err Name: ","$($_.Exception.GetType().FullName) ","Err Msg: ","$($_.Exception.Message)" -Color Red, Magenta, Red, Magenta, Red, Magenta -ShowTime
                                $GraphAPI = $False
                                $RemoveGraphAPI = $True
                            }
                        }
                        "N" {
                            Write-Color -Text "Graph API module will not be installed. ","Report will show on-premises AD users only." -Color White, Yellow -ShowTime
                            $GraphAPI = $False
                        }
                        Default { 
                            Write-Color -Text "No option was selected. ","Graph API module will not be installed. Report will show on-premises AD users only." -Color White, Yellow -ShowTime
                            $GraphAPI = $False
                        }
                    }
                } Else {
                    Write-Color -Text "NOTICE: If authorized to install PowerShell modules on this system you can run this script in an admin session to install the Graph API module and run the audit against cloud users and combine cloud properties with on-prem properties." -Color Yellow -ShowTime
                }
            }

            If ($GraphAPI) {
                Try {
                    Connect-MgGraph -Scopes 'Directory.Read.All, User.Read.All, AuditLog.Read.All' -NoWelcome -ErrorAction Stop
                    Try {
                        $AzUsers = Get-MgUser -All -Property Id, UserPrincipalName, SignInActivity, OnPremisesSyncEnabled, displayName, samAccountName, AccountEnabled, mail, lastPasswordChangeDateTime, PasswordPolicies, CreatedDateTime -ErrorAction Stop
                        $PremiumEntraLicense = $True
                    } Catch {
                        If ($_.Exception.Message -like "*Neither tenant is B2C or tenant doesn't have premium license*") {
                            Write-Color -Text "WARNING: This tenant does not have a premium license. LastLogonDate will show on-premises AD datetimes only!" -Color Yellow -ShowTime
                            $AzUsers = Get-MgUser -All -Property Id, UserPrincipalName, OnPremisesSyncEnabled, displayName, samAccountName, AccountEnabled, mail, lastPasswordChangeDateTime, PasswordPolicies, CreatedDateTime -ErrorAction Stop
                            $PremiumEntraLicense = $False
                        }
                    }

                    $GlobalAdminRoleId = Get-MgDirectoryRole | Where-Object {$_.DisplayName -eq "Global Administrator"} | Select-Object -ExpandProperty ID
                    $GlobalAdminMembers = Get-MgDirectoryRoleMemberAsUser -DirectoryRoleId $GlobalAdminRoleId
                    $Entra = $True
                } Catch {
                    Write-Color -Text "ERROR: Connection to Graph API failed!" -Color Red -ShowTime
                    Write-Color -Text "Would you like to try connecting to the Graph API again? ","(Y/N): " -Color White, Yellow -NoNewLine -ShowTime; $TryAgain = Read-Host
                    Switch ($TryAgain) {
                        "Y" {
                            Initialize-Entra -RemoveGraphAPI $RemoveGraphAPI -UntrustPSGallery $UntrustPSGallery -RemovePSGallery $RemovePSGallery -RemoveNuGet $RemoveNuGet
                        }
                        "N" {
                            Write-Color -Text "Graph API module will not be used. ","Report will show on-premises AD users only." -Color White, Yellow -ShowTime
                            $Entra = $False
                            $PremiumEntraLicense = $False
                            $AzUsers = $Null
                            $GlobalAdminMembers = $Null
                            If ($RemoveGraphAPI) {
                                Remove-Module -Name 'Microsoft.Graph.Users' -Force
                                Remove-Module -Name 'Microsoft.Graph.DirectoryObjects' -Force
                                Remove-Module -Name 'Microsoft.Graph.Identity.DirectoryManagement' -Force
                                Remove-Module -Name 'Microsoft.Graph.Authentication' -Force
                                Uninstall-Module -Name 'Microsoft.Graph' -Force
                            }
                        }
                    }
                }
            } Else {
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

    Return $Entra, $PremiumEntraLicense, $AzUsers, $GlobalAdminMembers, $RemoveGraphAPI, $UntrustPSGallery, $RemovePSGallery, $RemoveNuGet
}

Function Measure-ADUsers {
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

    param (
        [Parameter(Mandatory = $True)]$ADUsers,
        [Parameter(Mandatory = $True)]$AzUsers,
        [Parameter(Mandatory = $True)][boolean]$Entra
    )

    # Initialize arrays for UserCollection and AzUsersToProcess
    $UserCollection = @()
    $AzUsersToProcess = @()

    # Initialize user counter for progress bar
    $Count = 1

    # Process each user account found in active directory
    ForEach ($User in $ADUsers) {
        Write-Color -Text "Processing Active Directory Users" -ShowTime
        Write-Progress -Id 1 -Activity "Processing AD Users" -Status "Current Count: ($Count/$($ADUsers.Count))" -PercentComplete (($Count / $ADUsers.Count) * 100) -CurrentOperation "Processing... $($User.DisplayName)"

        # Check the users samAccountName against the list of Admin Users to verify if they are a domain admin
        If (($EnterpriseAdmins.SamAccountName) -contains $User.samAccountName) {
            $EnterpriseAdmin = $True
        } Else {
            $EnterpriseAdmin = $False
        }

        # Check the users samAccountName against the list of Admin Users to verify if they are a domain admin
        If (($DomainAdmins.SamAccountName) -contains $User.samAccountName) {
            $DomainAdmin = $True
        } Else {
            $DomainAdmin = $False
        }

        # If the account has an account expiration date then consider it expired.
        If ($Null -ne $User.AccountExpirationDate) {
            $AccountExpired = $User.AccountExpirationDate
        } Else {
            $AccountExpired = $Null
        }

        # If email property is blank then set to a blank space for formatting the spreadsheet. This stops a previous column from displaying over it.
        If ($Null -eq $User.mail) {
            $Mail = " "
        } Else {
            $Mail = $User.mail
        }

        # If an Entra connection was successful then process on-prem only users and write the rest to an array to process later
        If ($Entra) {
            # On-prem user without synced cloud user
            If (($AzUsers).UserPrincipalName -notcontains $User.UserPrincipalName) {
                # Add the user to the UserCollection
                $UserCollection += [PSCustomObject]@{
                    "Name" = $User.displayName
                    SamAccountName = $User.samAccountName
                    UserPrincipalName = $User.userPrincipalName
                    "Email Address" = $Mail
                    "User Type" = "On-Prem"
                    Enabled = $User.enabled
                    AccountExpiredDate = $AccountExpired
                    EnterpriseAdmin = $EnterpriseAdmin
                    DomainAdmin = $DomainAdmin
                    "AzGlobalAdmin" = "N/A"
                    PasswordLastSet = $User.passwordLastSet
                    LastLogonDate = $User.lastLogonDate
                    PasswordNeverExpires = $User.passwordNeverExpires
                    PasswordExpired = $User.passwordExpired
                    "Account Locked" = $User.lockedOut
                    CannotChangePassword = $User.cannotChangePassword
                    "Date Created" = $User.whenCreated
                    Notes = ""
                    Action = ""
                    "Follow Up" = ""
                    Resolution = ""
                }
            # Otherwise add the user to array AzUsersToProcess to be processed by Merge-AzUsers function
            } Else {
                $AzUsersToProcess += $User.UserPrincipalName
            }
        # No connection to Entra ID. Process active directory users only
        } Else {
            # Add the user to the UserCollection
            $UserCollection += [PSCustomObject]@{
                "Name" = $User.displayName
                SamAccountName = $User.samAccountName
                UserPrincipalName = $User.userPrincipalName
                "Email Address" = $Mail
                "User Type" = "On-Prem"
                Enabled = $User.enabled
                AccountExpiredDate = $AccountExpired
                EnterpriseAdmin = $EnterpriseAdmin
                DomainAdmin = $DomainAdmin
                PasswordLastSet = $User.passwordLastSet
                LastLogonDate = $User.lastLogonDate
                PasswordNeverExpires = $User.passwordNeverExpires
                PasswordExpired = $User.passwordExpired
                "Account Locked" = $User.lockedOut
                CannotChangePassword = $User.cannotChangePassword
                "Date Created" = $User.whenCreated
                Notes = ""
                Action = ""
                "Follow Up" = ""
                Resolution = ""
            }
        }

        $Count += 1
    }

    Return $UserCollection, $AzUsersToProcess
}

Function Merge-AzUsers {
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

    param (
        [Parameter(Mandatory = $True)]$ADUsers,
        [Parameter(Mandatory = $True)]$AzUsers,
        [Parameter(Mandatory = $True)]$AzUsersToProcess,
        [Parameter(Mandatory = $True)]$UserCollection
    )

    # Initialize user counter for progress bar
    $Count = 1

    ForEach ($AzUser in $AzUsers) {
        Write-Color -Text "Processing Entra ID Users" -ShowTime
        Write-Progress -Id 1 -Activity "Processing Entra Users" -Status "Current Count: ($Count/$($AzUsers.Count))" -PercentComplete (($Count / $AzUsers.Count) * 100) -CurrentOperation "Processing... $($AzUser.DisplayName)"

        # On-Prem user with synced cloud user
        If ($AzUsersToProcess -contains $AzUser.UserPrincipalName) {
            $User = $ADUsers | Where-Object {$_.UserPrincipalName -eq $AzUser.UserPrincipalName}

            # Check the users samAccountName against the list of Admin Users to verify if they are a domain admin
            If (($EnterpriseAdmins.SamAccountName) -contains $User.samAccountName) {
                $EnterpriseAdmin = $True
            } Else {
                $EnterpriseAdmin = $False
            }

            # Check the users samAccountName against the list of Admin Users to verify if they are a domain admin
            If (($DomainAdmins.SamAccountName) -contains $User.samAccountName) {
                $DomainAdmin = $True
            } Else {
                $DomainAdmin = $False
            }

            # If the account has an account expiration date then consider it expired.
            If ($Null -ne $User.AccountExpirationDate) {
                $AccountExpired = $User.AccountExpirationDate
            } Else {
                $AccountExpired = $Null
            }

            # Check if user is a global admin in Entra ID
            If (($GlobalAdminMembers).UserPrincipalName -contains $AzUser.UserPrincipalName) {
                $GlobalAdmin = $True
            } Else {
                $GlobalAdmin = $False
            }

            # If email property is blank then set to a blank space for formatting the spreadsheet. This stops a previous column from displaying over it.
            If ($Null -eq $User.mail) {
                $Mail = " "
            } Else {
                $Mail = $User.mail
            }

            # If the tenant has a premium license then get and compare the last sign-in timestamp and lastLogonDate timestamp
            If ($PremiumEntraLicense) {
                If ($AzUser.signInActivity.lastSignInDateTime) { 
                    $AzlastLogonDate = [DateTime]$AzUser.signInActivity.lastSignInDateTime
                    # If the last sign-in timestamp is newer than set that as lastLogonDate property
                    If ($User.lastLogonDate -lt $AzlastLogonDate) {
                        $LastLogonDate = $AzlastLogonDate
                    # Otherwise use the active directory lastLogonDate timestamp
                    } Else {
                        $LastLogonDate = $User.lastLogonDate
                    }
                # If there is no last sign-in timestamp then default to AD lastLogonDate timestamp.
                } Else {
                    $LastLogonDate = $User.lastLogonDate
                }
            # If the tenant doesnt have a premium license then we cant get last sign-in timestamp. Default to AD lastLogonDate timestamp.
            } Else {
                $LastLogonDate = $User.lastLogonDate
            }
            
            # Add the user to the UserCollection
            $UserCollection += [PSCustomObject]@{
                "Name" = $User.displayName
                SamAccountName = $User.samAccountName
                UserPrincipalName = $User.userPrincipalName
                "Email Address" = $Mail
                "User Type" = "Hybrid"
                Enabled = $User.enabled
                AccountExpiredDate = $AccountExpired
                EnterpriseAdmin = $EnterpriseAdmin
                DomainAdmin = $DomainAdmin
                "AzGlobalAdmin" = $GlobalAdmin
                PasswordLastSet = $User.passwordLastSet
                LastLogonDate = $LastLogonDate
                PasswordNeverExpires = $User.passwordNeverExpires
                PasswordExpired = $User.passwordExpired
                "Account Locked" = $User.lockedOut
                CannotChangePassword = $User.cannotChangePassword
                "Date Created" = $User.whenCreated
                Notes = ""
                Action = ""
                "Follow Up" = ""
                Resolution = ""
            }

        # Cloud only user
        } Else {
            # Check if user is a global admin in Entra ID
            If (($GlobalAdminMembers).UserPrincipalName -contains $AzUser.UserPrincipalName) {
                $GlobalAdmin = $True
            } Else {
                $GlobalAdmin = $False
            }

            # If the tenant has a premium license then grab the last sign-in timestamp
            If ($PremiumEntraLicense) {
                If ($AzUser.signInActivity.lastSignInDateTime) { 
                    $LastLogonDate = [DateTime]$AzUser.signInActivity.lastSignInDateTime
                } Else {
                    $LastLogonDate = $Null
                }
            } Else {
                $LastLogonDate = $Null
            }

            # If string found in PasswordPolicies then the password is set to never expire
            If ($AzUser.PasswordPolicies -contains "DisablePasswordExpiration") {
                $PasswordNeverExpires = $True
            } Else {
                $PasswordNeverExpires = $False
            }

            # If email property is blank then set to a blank space for formatting the spreadsheet. This stops a previous column from displaying over it.
            If ($Null -eq $AzUser.mail) {
                $Mail = " "
            } Else {
                $Mail = $AzUser.mail
            }

            # Add the user to the UserCollection
            $UserCollection += [PSCustomObject]@{
                "Name" = $AzUser.displayName
                SamAccountName = "N/A"
                UserPrincipalName = $AzUser.userPrincipalName
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

    Return $UserCollection
}

##############################################################################################################
#                                                   Main                                                     #
##############################################################################################################
Try {
    Clear-Host
    Write-Color -Text "__________________________________________________________________________________________" -Color White -BackGroundColor Black -HorizontalCenter $True -VerticalCenter $True
    Write-Color -Text "|                                                                                          |" -Color White -BackGroundColor Black -HorizontalCenter $True
    Write-Color -Text "|","                                            .-.                                           ","|" -Color White, DarkBlue, White -BackGroundColor Black, Black, Black -HorizontalCenter $True
    Write-Color -Text "|","                                            -#-              -.    -+                     ","|" -Color White, DarkBlue, White -BackGroundColor Black, Black, Black -HorizontalCenter $True
    Write-Color -Text "|","    ....           .       ...      ...     -#-  .          =#:..  .:      ...      ..    ","|" -Color White, DarkBlue, White -BackGroundColor Black, Black, Black -HorizontalCenter $True
    Write-Color -Text "|","   +===*#-  ",".:","     #*  *#++==*#:   +===**:  -#- .#*    -#- =*#+++. +#.  -*+==+*. .*+-=*.  ","|" -Color White, DarkBlue, Cyan, DarkBlue, White -BackGroundColor Black, Black, Black, Black, Black -HorizontalCenter $True
    Write-Color -Text "|","    .::.+#  ",".:","     #*  *#    .#+   .::..**  -#-  .#+  -#=   =#:    +#. =#:       :#+:     ","|" -Color White, DarkBlue, Cyan, DarkBlue, White -BackGroundColor Black, Black, Black, Black, Black -HorizontalCenter $True
    Write-Color -Text "|","  =#=--=##. ",".:","     #*  *#     #+  **---=##  -#-   .#+-#=    =#:    +#. **          :=**.  ","|" -Color White, DarkBlue, Cyan, DarkBlue, White -BackGroundColor Black, Black, Black, Black, Black -HorizontalCenter $True
    Write-Color -Text "|","  **.  .*#. ",".:.","   =#=  *#     #+ :#=   :##  -#-    :##=     -#-    +#. :#*:  .:  ::  .#=  ","|" -Color White, DarkBlue, Cyan, DarkBlue, White -BackGroundColor Black, Black, Black, Black, Black -HorizontalCenter $True
    Write-Color -Text "|","   -+++--=      .==:   ==     =-  .=++=-==  :=:    .#=       -++=  -=    :=+++-. :=++=-   ","|" -Color White, DarkBlue, White -BackGroundColor Black, Black, Black -HorizontalCenter $True
    Write-Color -Text "|","                                                  .#+                                     ","|" -Color White, DarkBlue, White -BackGroundColor Black, Black, Black -HorizontalCenter $True
    Write-Color -Text "|","                                                  *+                                      ","|" -Color White, DarkBlue, White -BackGroundColor Black, Black, Black -HorizontalCenter $True
    Write-Color -Text "|__________________________________________________________________________________________|" -Color White -BackGroundColor Black -HorizontalCenter $True
    Write-Color -Text "Script:","User Audit Report" -Color Yellow, White -BackGroundColor Black -LinesBefore 1
    Write-Color -Text "Checking for optional but recommended PowerShell modules" -ShowTime
    $ImportExcel, $RemoveImportExcel, $IEUntrustPSGallery, $IERemovePSGallery, $IERemoveNuGet = Initialize-ImportExcel
    $Entra, $PremiumEntraLicense, $AzUsers, $GlobalAdminMembers, $RemoveGraphAPI, $MgUntrustPSGallery, $MgRemovePSGallery, $MgRemoveNuGet = Initialize-Entra

    If ($IEUntrustPSGallery -or $MgUntrustPSGallery) {
        $UntrustPSGallery = $True
    }

    If ($IERemovePSGallery -or $MgRemovePSGallery) {
        $RemovePSGallery = $True
    }

    If ($IERemoveNuGet -or $MgRemoveNuGet) {
        $RemoveNuGet = $True
    }

    # Get the domain name
    $DomainName = (Get-ADDomain).DNSRoot

    # Get the Enterprise Admins group members
    $DomainAdmins = Get-ADGroupMember -Identity "Enterprise Admins"

    # Get the Domain Admins group members
    $DomainAdmins = Get-ADGroupMember -Identity "Domain Admins"

    # Create CSV of AD Users
    $ADUsers = Get-ADUser -Filter * -Properties *

    # Process the AD users. If Entra is enabled then process on-prem AD users only.
    $ProcessedADUsers, $AzUsersToProcess = Measure-ADUsers -ADUsers $ADUsers -AzUsers $AzUsers -Entra $Entra

    # If Entra is enabled, process hybrid and cloud only users and merge LastLogonDate for hybrid users.
    $UserCollection = Merge-AzUsers -ADUsers $ADUsers -AzUsers $AzUsers -AzUsersToProcess $AzUsersToProcess -UserCollection $ProcessedADUsers

    # Sort the user collection by DisplayName. We have to sort before we export to Excel if we want the table sorted a specific way.
    $SortedCollection = $UserCollection | Sort-Object -Property Name

    # Timestamp for Filename
    $TimeStamp = Get-Date -Format "MMddyyyy_HHmm"

    If ($ImportExcel) {
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

        If ($Entra) {
            # Center align rows that will have "N/A" for a cleaner look
            Set-ExcelRange -Worksheet $Worksheet -Range "G:G" -HorizontalAlignment Center -NumberFormat "MM/dd/yyyy hh:mm AM/PM"
            Set-ExcelRange -Worksheet $Worksheet -Range "J:J" -HorizontalAlignment Center
            Set-ExcelRange -Worksheet $Worksheet -Range "N:N" -HorizontalAlignment Center
            Set-ExcelRange -Worksheet $Worksheet -Range "O:O" -HorizontalAlignment Center
            Set-ExcelRange -Worksheet $Worksheet -Range "P:P" -HorizontalAlignment Center

            # Add conditional formatting to the data in the columns
            # Enabled Column
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "F2:F$lastCol" -RuleType Equal -ConditionValue $False -BackgroundColor Yellow
            # AccountExpired
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "G2:G$lastCol" -RuleType NotContainsBlanks -BackgroundColor Yellow
            # Enterprise Admin
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "H2:H$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
            # Domain Admin
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "I2:I$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
            # Global Admin
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "J2:J$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
            # PasswordLastSet
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "K2:K$lastCol" -RuleType Expression -ConditionValue "=`$K2<=(TODAY()-90)" -BackgroundColor Yellow -Bold
            # LastLogonDate
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "L2:L$lastCol" -RuleType Expression -ConditionValue "=`$L2<=(TODAY()-180)" -BackgroundColor Red -Bold
            # LastLogonDate
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "L2:L$lastCol" -RuleType Expression -ConditionValue "=`=AND(`$L2 > TODAY()-180, `$L2 < TODAY()-90)" -BackgroundColor Yellow 
            # LastLogonDate
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "L2:L$lastCol" -RuleType Expression -ConditionValue "=`$L2>=(TODAY()-90)" -BackgroundColor LightGreen
            # PasswordNeverExpires
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "M2:M$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Red -Bold
            # PasswordExpired
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "N2:N$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
            # Account Locked
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "O2:O$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
            # CannotChangePassword
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "P2:P$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
        } Else {
            # Add conditional formatting to the data in the columns
            # Enabled Column
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "F2:F$lastCol" -RuleType Equal -ConditionValue $False -BackgroundColor Yellow
            # AccountExpired
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "G2:G$lastCol" -RuleType NotContainsBlanks -BackgroundColor Yellow
            # Enterprise Admin
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "H2:H$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
            # DomainAdmin
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "I2:I$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
            # PasswordLastSet
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "J2:J$lastCol" -RuleType Expression -ConditionValue "=`$J2<=(TODAY()-90)" -BackgroundColor Red -Bold
            # LastLogonDate
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "K2:K$lastCol" -RuleType Expression -ConditionValue "=`$K2<=(TODAY()-180)" -BackgroundColor Red -Bold
            # LastLogonDate
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "K2:K$lastCol" -RuleType Expression -ConditionValue "=`=AND(`$K2 > TODAY()-180, `$K2 < TODAY()-90)" -BackgroundColor Yellow 
            # LastLogonDate
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "K2:K$lastCol" -RuleType Expression -ConditionValue "=`$K2>=(TODAY()-90)" -BackgroundColor LightGreen
            # PasswordNeverExpires
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "L2:L$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Red -Bold
            # PasswordExpired
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "M2:M$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
            # Account Locked
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "N2:N$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
            # CannotChangePassword
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "O2:O$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
        }
        Close-ExcelPackage $XLSX
    } Else {
        # Format the file name with the domain name
        $FileName = "C:\Temp\$($domainName)_Users_Report_$TimeStamp.csv"
        Export-Csv -Path $FileName -NoTypeInformation
    }

    Write-Color -Text "Report successfully saved to: ", "$FileName" -Color Green, White -ShowTime
    Write-Color -Text "" -Color Blue -HorizontalCenter $True -LinesBefore 1
} Catch {
    Write-Error "Err Line: $($_.InvocationInfo.ScriptLineNumber) Err Name: $($_.Exception.GetType().FullName) Err Msg: $($_.Exception.Message)"
} Finally {
    # When the script exits revert all packageprovider and repository changes and remove installed modules if not previously installed
    Try {
        If ($UntrustPSGallery) {
            Set-PSRepository -Name 'PSGallery' -InstallationPolicy Untrusted
        }

        If ($RemovePSGallery) {
            Unregister-PSRepository -Name 'PSGallery'
        }

        If ($RemoveImportExcel) {
            Remove-Module -Name 'ImportExcel' -Force
            Uninstall-Module -Name 'ImportExcel' -Force
        }

        If ($RemoveGraphAPI) {
            Remove-Module -Name 'Microsoft.Graph.Users' -Force
            Remove-Module -Name 'Microsoft.Graph.DirectoryObjects' -Force
            Remove-Module -Name 'Microsoft.Graph.Identity.DirectoryManagement' -Force
            Remove-Module -Name 'Microsoft.Graph.Authentication' -Force
            Uninstall-Module -Name 'Microsoft.Graph' -Force
        }

        If ($RemoveNuGet) {
            Uninstall-PackageProvider -Name NuGet -Force
        }
    } Catch {

    }
}
