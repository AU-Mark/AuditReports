$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$AdminSession = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

$TimeStamp = Get-Date -Format "MMddyyyy_HHmm"

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

        # Format the file name with the domain name
        $FileName = "C:\Temp\$($domainName)_Users_Report_$TimeStamp.xlsx"
    } Else {
        If ($AdminSession) {
            $Install = Read-Host 'ImportExcel module is not installed, would you like to temporarily install it for this script? (Y/N)'

            Switch ($Install) {
                "Y" {
                    Try {
                        Set-PSRepository -Name 'PSGallery' -InstallationPolicy Trusted
                        Install-PackageProvider -Name NuGet
                        Install-Module -Name 'ImportExcel'-Force
                        Write-Host -ForegroundColor Green "ImportExcel installed successfully. It will be removed after running this script."
                        $ImportExcel = $True
                        $RemoveImportExcel = $True
                    } Catch {

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
        } Else {
            Write-Host -ForegroundColor Yellow "If allowed to install PowerShell modules on this system you can run this script in an admin session to install the ImportExcel module and save directly to xlsx with formatting"
        }

        $FileName = "C:\Temp\$($domainName)_Users_Report_$TimeStamp.csv"
    }

    # Get the Domain Admins group
    $DomainAdmins = Get-ADGroupMember -Identity "Domain Admins"

    # Create CSV of AD Users
    $ADUsers = Get-ADUser -Filter * -Properties * 

    # Initialize an empty array to store the results
    $UserCollection = @()

    # Initialize user counter for progress bar
    $Count = 1

    ForEach ($User in $ADUsers) {
        Write-Progress -Id 1 -Activity "Processing Domain Users" -Status "Current Count: ($Count/$($ADUsers.Count))" -PercentComplete (($Count / $ADUsers.Count) * 100) -CurrentOperation "Processing $($User.displayName)..."


        If (($DomainAdmins.SamAccountName) -contains $User.samAccountName) {
            $DomainAdmin = $True
        } Else {
            $DomainAdmin = $False
        }

        If ($Null -ne $User.AccountExpirationDate) {
            $AccountExpired = $True
        } Else {
            $AccountExpired = $False
        }

        $UserCollection += [PSCustomObject]@{
            Name = $User.displayName
            samAccountName = $User.samAccountName
            UserPrincipalName = $User.userPrincipalName
            Enabled = $User.enabled
            AccountExpired = $AccountExpired
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

        $Count += 1
    }

    # Check if user can/wants to install ImportExcel module

    If ($ImportExcel) {
        $XLSX = $UserCollection | Export-Excel $FileName -WorksheetName "AD Users" -AutoSize -FreezeTopRow -AutoFilter -BoldTopRow -PassThru 
        $Worksheet = $XLSX.Workbook.Worksheets["AD Users"]
        $lastRow = $Worksheet.Dimension.End.Row
        Add-ConditionalFormatting -WorkSheet $Worksheet -address "D2:D$Lastrow" -RuleType Equal -ConditionValue $False -BackgroundColor Yellow
        Add-ConditionalFormatting -WorkSheet $Worksheet -address "E2:E$Lastrow" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
        Add-ConditionalFormatting -WorkSheet $Worksheet -address "F2:F$Lastrow" -RuleType Equal -ConditionValue $True -BackgroundColor Green
        Add-ConditionalFormatting -WorkSheet $Worksheet -address "G2:G$Lastrow" -RuleType Expression -ConditionValue "=`$G2<=(TODAY()-90)" -BackgroundColor Yellow
        Add-ConditionalFormatting -WorkSheet $Worksheet -address "H2:H$Lastrow" -RuleType Expression -ConditionValue "=`$H2<=(TODAY()-180)" -BackgroundColor Red
        Add-ConditionalFormatting -WorkSheet $Worksheet -address "H2:H$Lastrow" -RuleType Expression -ConditionValue "=`=AND(`$H2 > TODAY()-180, `$H2 < TODAY()-90)" -BackgroundColor Yellow 
        Add-ConditionalFormatting -WorkSheet $Worksheet -address "H2:H$Lastrow" -RuleType Expression -ConditionValue "=`$H2>=(TODAY()-90)" -BackgroundColor Green
        Add-ConditionalFormatting -WorkSheet $Worksheet -address "I2:I$Lastrow" -RuleType Equal -ConditionValue $True -BackgroundColor Red
        Add-ConditionalFormatting -WorkSheet $Worksheet -address "J2:J$Lastrow" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
        Add-ConditionalFormatting -WorkSheet $Worksheet -address "K2:K$Lastrow" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
        Add-ConditionalFormatting -WorkSheet $Worksheet -address "L2:L$Lastrow" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
        Add-ConditionalFormatting -WorkSheet $Worksheet -address "M2:M$Lastrow" -RuleType Equal -ConditionValue $True -BackgroundColor Yellow
        Close-ExcelPackage $XLSX

        If ($RemoveImportExcel) {
            Set-PSRepository -Name 'PSGallery' -InstallationPolicy Untrusted
            Uninstall-PackageProvider -Name NuGet
            Remove-Module -Name 'ImportExcel'
            Uninstall-Module -Name 'ImportExcel'
        }
    } Else {
        Export-Csv -Path $FileName -NoTypeInformation
    }
} Catch {
    Write-Error "Err Line: $($_.InvocationInfo.ScriptLineNumber) Err Name: $($_.Exception.GetType().FullName) Err Msg: $($_.Exception.Message)"
}