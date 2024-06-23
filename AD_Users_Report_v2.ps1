# AD User Audit Report

# Check if powershell is running in an admin session
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$AdminSession = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

# Globals
$RemovePSGallery = $False
$RemoveNuGet = $False
$RemoveImportExcel = $False
$RemoveGraphAPI = $False

Try {
    # Get the domain name
    $DomainName = (Get-ADDomain).DNSRoot

    # Check if ImportExcel module is installed
    If (Get-Module -ListAvailable -Name 'ImportExcel') {
        Write-Host -ForegroundColor Yellow "ImportExcel module detected. Will save directly to XLSX with formatting..."

        # Import the ImportExcel module and set the $ImportExcel variable to True
        Import-Module ImportExcel
        $ImportExcel = $True
        $RemoveImportExcel = $False
    } Else {
        If ($AdminSession) {
            # ImportExcel module is not installed. Ask if allowed to install and user wants to install it.
            Write-Warning 'ImportExcel module is not installed. Without it the report will output in CSV and you will have to format it manually.'
            $InstallImportExcel = Read-Host 'If allowed to install modules on this system, would you like to temporarily install it for this script? (Y/N)'

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
                        Install-Module -Name 'ImportExcel'-Force
                        Write-Host -ForegroundColor Green "ImportExcel installed successfully. It will be removed after running this script."
                        $ImportExcel = $True
                        $RemoveImportExcel = $True
                    } Catch {
                        Write-Error "ImportExcel module failed to install. See the error below. The report will output to CSV only until the error is corrected. The ImportExcel module will be uninstalled at the end of this script"
                        Write-Error "Err Line: $($_.InvocationInfo.ScriptLineNumber) Err Name: $($_.Exception.GetType().FullName) Err Msg: $($_.Exception.Message)"
                        $ImportExcel = $False
                        $RemoveImportExcel = $True
                    }
                }
                "N" {
                    Write-Host -ForegroundColor Yellow "ImportExcel will not be installed. Proceeding to use CSV format."
                    $ImportExcel = $False
                }
                Default { 
                    Write-Host -ForegroundColor Yellow "No option was selected. Proceeding to use CSV format."
                    $ImportExcel = $False
                }
            }

            $EntraID = Read-Host "Would you like to connect to Entra ID? (Y/N)"
            Switch ($EntraID) {
                'Y' { 
                    If (Get-Module -ListAvailable -Name 'Microsoft.Graph') {
                        Write-Host -ForegroundColor Yellow "Microsoft.Graph module detected. Connecting to Graph API..."
                
                        # Import the ImportExcel module and set the $ImportExcel variable to True
                        Import-Module Microsoft.Graph.Users
                        Import-Module Microsoft.Graph.Groups
                        Import-Module Microsoft.Graph.DirectoryObjects
                        $GraphAPI = $True
                        $RemoveGraphAPI = $False
                    } Else {
                        # Graph API module is not installed. Ask if allowed to install and user wants to install it.
                        Write-Warning 'Graph API module is not installed. Without it the report will display on-premises AD Users only.'
                        $InstallGraph = Read-Host 'If allowed to install modules on this system, would you like to temporarily install it for this script? (Y/N)'
            
                        Switch ($InstallGraph) {
                            "Y" {
                                Try {
                                    Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
                                    If ((Get-PackageProvider).Name -notcontains "NuGet") {
                                        Install-PackageProvider -Name NuGet -Force
                                        $RemoveNuGet = $True
                                    } Else {
                                        $RemoveNuGet = $False
                                    }
                                    Install-Module -Name 'Microsoft.Graph'-Force
                                    Write-Host -ForegroundColor Green "Microsoft.Graph installed successfully. It will be removed after running this script."
                                    $GraphAPI = $True
                                    $RemoveGraphAPI = $True
                                } Catch {
                                    Write-Error "Graph API module failed to install. See the error below. The report will include on-premises AD users only until the error is corrected. The Graph API module will be uninstalled at the end of this script"
                                    Write-Error "Err Line: $($_.InvocationInfo.ScriptLineNumber) Err Name: $($_.Exception.GetType().FullName) Err Msg: $($_.Exception.Message)"
                                    $GraphAPI = $False
                                    $RemoveGraphAPI = $True
                                }
                            }
                            "N" {
                                Write-Host -ForegroundColor Yellow "Graph API will not be installed. Report will show on-premises AD users only."
                                $GraphAPI = $False
                            }
                            Default { 
                                Write-Host -ForegroundColor Yellow "No option was selected. Graph API will not be installed. Report will show on-premises AD users only."
                                $GraphAPI = $False
                            }
                        }
                    }

                    If ($GraphAPI) {
                        Try {
                        Connect-MgGraph -Scopes 'Directory.Read.All, User.Read.All, Group.Read.All, AuditLog.Read.All' -NoWelcome -ErrorAction Stop
                            Try {
                                $AzUsers = Get-MgUser -All -Property Id, UserPrincipalName, SignInActivity, OnPremisesSyncEnabled, displayName, samAccountName, AccountEnabled, mail, lastPasswordChangeDateTime, PasswordPolicies, CreatedDateTime -ErrorAction Stop
                                $PremiumEntraLicense = $True
                            } Catch {
                                If ($_.Exception.Message -like "*Neither tenant is B2C or tenant doesn't have premium license*") {
                                    $AzUsers = Get-MgUser -All -Property Id, UserPrincipalName, OnPremisesSyncEnabled, displayName, samAccountName, AccountEnabled, mail, lastPasswordChangeDateTime, PasswordPolicies, CreatedDateTime -ErrorAction Stop
                                    $PremiumEntraLicense = $False
                                }
                            }

                            $GlobalAdminRoleId = Get-MgDirectoryRole | Where-Object {$_.DisplayName -eq "Global Administrator"} | Select-Object -ExpandProperty ID
                            $GlobalAdminMembers = Get-MgDirectoryRoleMemberAsUser -DirectoryRoleId $GlobalAdminRoleId
                            $Entra = $True
                        } Catch {
                            Write-Warning "Connection to Graph API failed. Report will show on-premises AD users only."
                            $Entra = $False
                        }
                    } Else {
                        Write-Warning "Connection to Graph API failed. Report will show on-premises AD users only."
                        $Entra = $False
                    }
                }
                'N' {
                    $Entra = $False
                }
                Default {
                    $Entra = $False
                }
            }
        } Else {
            Write-Host -ForegroundColor Yellow "If allowed to install PowerShell modules on this system you can run this script in an admin session to install the Graph API and ImportExcel modules and save directly to xlsx with formatting"
        }
    }

    # Get the Enterprise Admins group members
    $DomainAdmins = Get-ADGroupMember -Identity "Enterprise Admins"

    # Get the Domain Admins group members
    $DomainAdmins = Get-ADGroupMember -Identity "Domain Admins"

    # Create CSV of AD Users
    $ADUsers = Get-ADUser -Filter * -Properties * 

    # Initialize an empty array to store the results
    $UserCollection = @()

    # Initialize user counter for progress bar
    $Count = 1

    ForEach ($User in $ADUsers) {
        Write-Progress -Id 1 -Activity "Processing Users" -Status "Current Count: ($Count/$($ADUsers.Count))" -PercentComplete (($Count / $ADUsers.Count) * 100) -CurrentOperation "Processing..."

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

        If ($Null -ne $User.AccountExpirationDate) {
            $AccountExpired = $User.AccountExpirationDate
        } Else {
            $AccountExpired = $Null
        }

        If ($Entra) {
            # On-prem user without synced cloud user
            If (($AzUsers).UserPrincipalName -notcontains $User.UserPrincipalName) {
                $UserCollection += [PSCustomObject]@{
                    "Name" = $User.displayName
                    SamAccountName = $User.samAccountName
                    UserPrincipalName = $User.userPrincipalName
                    "Email Address" = $User.mail
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
            } Else {
                ForEach ($AzUser in $AzUsers) {
                    # On-Prem user with synced cloud user
                    If (($User).UserPrincipalName -contains $AzUser.UserPrincipalName) {
                        # Check if user is a global admin in Entra ID
                        If (($GlobalAdminMembers).UserPrincipalName -contains $AzUser.UserPrincipalName) {
                            $GlobalAdmin = $True
                        } Else {
                            $GlobalAdmin = $False
                        }

                        #TODO Compare last sign-in date from AD and Graph and use the latest sign-in date
                        If ($PremiumEntraLicense) {
                            If ($AzUser.signInActivity.lastSignInDateTime) { 
                                $AzlastLogonDate = [DateTime]$AzUser.signInActivity.lastSignInDateTime
                                If ($User.lastLogonDate -lt $AzlastLogonDate) {
                                    $LastLogonDate = $AzlastLogonDate
                                } Else {
                                    $LastLogonDate = $User.lastLogonDate
                                }
                            } Else {
                                $LastLogonDate = $User.lastLogonDate
                            }
                        } Else {
                            $LastLogonDate = $User.lastLogonDate
                        }
                        
                        $UserCollection += [PSCustomObject]@{
                            "Name" = $User.displayName
                            SamAccountName = $User.samAccountName
                            UserPrincipalName = $User.userPrincipalName
                            "Email Address" = $User.mail
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
                        If (($GlobalAdminMembers).UserPrincipalName -contains $User.UserPrincipalName) {
                            $GlobalAdmin = $True
                        } Else {
                            $GlobalAdmin = $False
                        }

                        If ($PremiumEntraLicense) {
                            If ($AzUser.signInActivity.lastSignInDateTime) { 
                                $LastLogonDate = [DateTime]$AzUser.signInActivity.lastSignInDateTime
                            } Else {
                                $LastLogonDate = $Null
                            }
                        } Else {
                            $LastLogonDate = $Null
                        }

                        If ($AzUser.PasswordPolicies -contains "DisablePasswordExpiration") {
                            $PasswordNeverExpires = $True
                        } Else {
                            $PasswordNeverExpires = $False
                        }

                        $UserCollection += [PSCustomObject]@{
                            "Name" = $AzUser.displayName
                            SamAccountName = $AzUser.samAccountName
                            UserPrincipalName = $AzUser.userPrincipalName
                            "Email Address" = $AzUser.mail
                            Enabled = $AzUser.AccountEnabled
                            AccountExpiredDate = $Null
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
                }
            }
        } Else {
            $UserCollection += [PSCustomObject]@{
                "Name" = $User.displayName
                SamAccountName = $User.samAccountName
                UserPrincipalName = $User.userPrincipalName
                "Email Address" = $User.mail
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

    # Sort the user collection by samAccountName. We have to sort before we export to Excel if we want the table sorted a specific way.
    $SortedCollection = $UserCollection | Sort-Object -Property samAccountName -Descending

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
        $lastCol = $Worksheet.Dimension.End.Column
        # Set the font sizw to 8 for the whole document
        Set-ExcelRange -Worksheet $Worksheet -Range "A:Z" -FontSize 8
        # Autosize the columns again after changing the font size
        $XLSX.Workbook.Worksheets.Cells.AutoFitColumns()

        If ($Entra) {
            # Add conditional formatting to the data in the columns
            # Enabled Column
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "E2:E$lastCol" -RuleType Equal -ConditionValue $False -BackgroundColor Yellow
            # AccountExpired
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "F2:F$lastCol" -RuleType NotContainsBlanks -BackgroundColor Yellow
            # Enterprise Admin
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "G2:G$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
            # Domain Admin
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "H2:H$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
            # Global Admin
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "I2:I$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
            # PasswordLastSet
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "J2:J$lastCol" -RuleType Expression -ConditionValue "=`$J2<=(TODAY()-90)" -BackgroundColor Yellow -Bold
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
            # What was this for?
            #Add-ConditionalFormatting -WorkSheet $Worksheet -address "P2:P$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
        } Else {
            # Add conditional formatting to the data in the columns
            # Enabled Column
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "E2:E$lastCol" -RuleType Equal -ConditionValue $False -BackgroundColor Yellow
            # AccountExpired
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "F2:F$lastCol" -RuleType NotContainsBlanks -BackgroundColor Yellow
            # Enterprise Admin
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "G2:G$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
            # DomainAdmin
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "H2:H$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor LightGreen -Bold
            # PasswordLastSet
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "I2:I$lastCol" -RuleType Expression -ConditionValue "=`$I2<=(TODAY()-90)" -BackgroundColor Yellow -Bold
            # LastLogonDate
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "J2:J$lastCol" -RuleType Expression -ConditionValue "=`$J2<=(TODAY()-180)" -BackgroundColor Red -Bold
            # LastLogonDate
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "J2:J$lastCol" -RuleType Expression -ConditionValue "=`=AND(`$J2 > TODAY()-180, `$J2 < TODAY()-90)" -BackgroundColor Yellow 
            # LastLogonDate
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "J2:J$lastCol" -RuleType Expression -ConditionValue "=`$J2>=(TODAY()-90)" -BackgroundColor LightGreen
            # PasswordNeverExpires
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "K2:K$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Red -Bold
            # PasswordExpired
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "L2:L$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
            # Account Locked
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "M2:M$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
            # CannotChangePassword
            Add-ConditionalFormatting -WorkSheet $Worksheet -address "N2:N$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
            # What was this for?
            # Add-ConditionalFormatting -WorkSheet $Worksheet -address "O2:O$lastCol" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
        }
        Close-ExcelPackage $XLSX
    } Else {
        # Format the file name with the domain name
        $FileName = "C:\Temp\$($domainName)_Users_Report_$TimeStamp.csv"
        Export-Csv -Path $FileName -NoTypeInformation
    }

    If ($UntrustPSGallery) {
        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Untrusted
    }

    If ($RemovePSGallery) {
        Unregister-PSRepository -Name 'PSGallery'
    }

    If ($RemoveImportExcel) {
        Remove-Module -Name 'ImportExcel'
        Uninstall-Module -Name 'ImportExcel' -Force
    }

    If ($RemoveGraphAPI) {
        Uninstall-Module -Name 'Microsoft.Graph' -Force
    }

    If ($RemoveNuGet) {
        Uninstall-PackageProvider -Name NuGet -Force
    }
} Catch {
    Write-Error "Err Line: $($_.InvocationInfo.ScriptLineNumber) Err Name: $($_.Exception.GetType().FullName) Err Msg: $($_.Exception.Message)"
}
