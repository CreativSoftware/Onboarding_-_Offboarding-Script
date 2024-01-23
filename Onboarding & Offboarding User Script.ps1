#Pulls information from Excel spreadsheet
Install-Module -Name ImportExcel
$users = Import-Excel -Path 'S:\Onboarding&Offboarding\new_users.xlsx'

#Verifies Authentication
$authenticate = $true
$attempts = 3
while ($authenticate) {
    $domain_username = Read-Host -Prompt "Enter YOUR ADMIN domain\username"
    $credentials = Get-Credential -UserName $domain_username -Message 'Enter Admin Password'
    try {
        $session = New-PSSession -ComputerName 'doidc02' -Credential $credentials -ErrorAction Stop
        Remove-PSSession $session
        Write-Host "Authentication successful" -ForegroundColor Green
        $authenticate = $false
    } catch {
        $attempts = $attempts - 1
        if ($attempts -eq 0){
            Write-Host "Too many failed attempts. Exiting console." -ForegroundColor Red
            exit
        }
        Write-Host "Failed to authenticate please try again. $attempts attempts remaining." -ForegroundColor Red
    }
}

#Get login of the user who opens this script.
$name = whoami.exe
$name = $name.split("\")
$login = $name[1]

#Reads each Row in Spreadsheet and creates the variables for the account creation
$user_answer = Read-Host -Prompt "Are you setting up a New User Account (Y or N)?"
if ($user_answer.ToLower() -eq "y"){
    $user_answer_two = Read-Host -Prompt "Did you update the Excel speadsheet(Y or N)?"
        if ($user_answer_two.ToLower() -eq "y"){
            Invoke-Command -ComputerName "doidc02" -Credential $credentials -ScriptBlock{
                #Global Varibles
                $credentials = $using:credentials
                $users = $using:users
                $Password = "Welcome1"
                $SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
                $AccountEnabled = $true
                $Street = '180 Maiden Lane'
                $City = 'New York'
                $Zip = '10038'
                $Country = "US"
                $State = 'New York'
                $Company = 'NYCDOI'
                $TelephoneNumber = "212-825-2541"

                foreach($user in $users){
                    $Name = $user.Name
                    $Email = $user.Email
                    $JobTitle = $user.Title
                    $Department = $user.Squad

                    $fullname = $user.Name.split(" ")
                    $FirstName = $fullname[0]
                    $LastName  = $fullname[1]
                    $Username = $FirstName[0].tostring() + $LastName
                    
                    $Mname = $user.Manager.split(" ")
                    $Mfirst = $Mname[0]
                    $Mlast = $Mname[1]
                    $Manager = Get-ADUser -Filter {GivenName -eq $Mfirst -and Surname -eq $Mlast} | Select-Object -First 1 |Select-Object -ExpandProperty SamAccountName

                    if ($Department -eq "squad1"){$OU = "OU=Squad1,OU=DOI Users,DC=DOI,DC=NYCNET"}
                    elseif ($Department -eq "squad2"){$OU = "OU=Squad2,OU=DOI Users,DC=DOI,DC=NYCNET"}
                    elseif ($Department -eq "squad3"){$OU = "OU=Squad3,OU=DOI Users,DC=DOI,DC=NYCNET"}
                    elseif ($Department -eq "squad4"){$OU = "OU=Squad4,OU=DOI Users,DC=DOI,DC=NYCNET"}
                    elseif ($Department -eq "squad5"){$OU = "OU=Squad5,OU=DOI Users,DC=DOI,DC=NYCNET"} 
                    elseif ($Department -eq "squad6"){$OU = "OU=Squad6,OU=DOI Users,DC=DOI,DC=NYCNET"}
                    elseif ($Department -eq "execunit"){$OU = "OU=ExecUnit,OU=DOI Users,DC=DOI,DC=NYCNET"}
                    elseif ($Department -eq "background"){$OU = "OU=Background,OU=DOI Users,DC=DOI,DC=NYCNET"}
                    elseif ($Department -eq "data analytics"){$OU = "OU=Data Analytics,OU=DOI Users,DC=DOI,DC=NYCNET"}
                    elseif ($Department -eq "digitalforensics"){$OU = "OU=DigitalForensics,OU=DOI Users,DC=DOI,DC=NYCNET"}
                    elseif ($Department -eq "facilities"){$OU = "OU=Facilities,OU=DOI Users,DC=DOI,DC=NYCNET"}
                    elseif ($Department -eq "finance"){$OU = "OU=Finance,OU=DOI Users,DC=DOI,DC=NYCNET"}
                    elseif ($Department -eq "fingerprint"){$OU = "OU=Fingerprint,OU=DOI Users,DC=DOI,DC=NYCNET"}
                    elseif ($Department -eq "intern"){$OU = "OU=Interns,OU=DOI Users,DC=DOI,DC=NYCNET"}
                    else {$OU = "OU=TestUsers,OU=DOI Users,DC=DOI,DC=NYCNET"}

                    #Creates New ADUser Account
                    New-ADUser `
                        -Name $Name `
                        -UserPrincipalName "$Username@doi.nyc.gov" `
                        -SamAccountName $Username `
                        -EmailAddress $Email `
                        -AccountPassword $SecurePassword `
                        -Enabled $AccountEnabled `
                        -Path $OU `
                        -GivenName $FirstName `
                        -Surname $LastName `
                        -DisplayName $Name `
                        -StreetAddress $Street `
                        -City $City `
                        -PostalCode $Zip `
                        -Country $Country `
                        -State $State `
                        -Title $JobTitle `
                        -Department $Department `
                        -Manager $Manager `
                        -Company $Company `
                        -HomeDirectory "\\doi\doi_share\home_folder\$Username" `
                        -HomeDrive 'I:' `
                        -OfficePhone $TelephoneNumber `
                        -Credential $credentials
                        
                    #Adds Specific Membership Groups
                    if ($OU -eq "OU=Squad1,OU=DOI Users,DC=DOI,DC=NYCNET"){
                        $groups = @('squad1', 'Investigation_Division_EDOCS', 'InvestigationDivision', 'RightFax_Squad1', 'Squad1_Media')
                        foreach($group in $groups){
                            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
                        }    
                    }
                    elseif ($OU -eq "OU=Squad2,OU=DOI Users,DC=DOI,DC=NYCNET"){
                        $groups = @('squad2', 'Investigation_Division_EDOCS', 'InvestigationDivision', 'RightFax_Squad2')
                        foreach($group in $groups){
                            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
                        }  
                    }
                    elseif ($OU -eq "OU=Squad3,OU=DOI Users,DC=DOI,DC=NYCNET"){
                        $groups = @('squad3', 'Investigation_Division_EDOCS', 'InvestigationDivision', 'RightFax_Squad3')
                        foreach($group in $groups){
                            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
                        }
                    }
                    elseif ($OU -eq "OU=Squad4,OU=DOI Users,DC=DOI,DC=NYCNET"){
                        $groups = @('squad4', 'Investigation_Division_EDOCS', 'InvestigationDivision', 'RightFax_Squad4')
                        foreach($group in $groups){
                            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
                        }
                    }
                    elseif ($OU -eq "OU=Squad5,OU=DOI Users,DC=DOI,DC=NYCNET"){
                        $groups = @('squad5', 'Investigation_Division_EDOCS', 'InvestigationDivision', 'RightFax_Squad5')
                        foreach($group in $groups){
                            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
                        }
                    }
                    elseif ($OU -eq "OU=Squad6,OU=DOI Users,DC=DOI,DC=NYCNET"){
                        $groups = @('squad6', 'Investigation_Division_EDOCS', 'InvestigationDivision', 'RightFax_Squad6')
                        foreach($group in $groups){
                            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
                        }
                    }
                    elseif ($OU -eq "OU=ExecUnit,OU=DOI Users,DC=DOI,DC=NYCNET"){
                        $groups = @('EXEC-UNIT','GeneralCounsel', 'Investigation_Division_EDOCS', 'InvestigationDivision', 'RightFax_Executives')
                        foreach($group in $groups){
                            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
                        }
                    }
                    elseif ($OU -eq "OU=Background,OU=DOI Users,DC=DOI,DC=NYCNET"){
                        $groups = @('BackgroundUnitStaff', 'Investigation_Division_EDOCS', 'InvestigationDivision', 'RightFax_Background')
                        foreach($group in $groups){
                            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
                        }
                    }
                    elseif ($OU -eq "OU=Data Analytics,OU=DOI Users,DC=DOI,DC=NYCNET"){
                        $groups = @('Data Analytics', 'Investigation_Division_EDOCS', 'Data_Analytics_SG')
                        foreach($group in $groups){
                            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
                        }
                    }
                    elseif ($OU -eq "OU=DigitalForensics,OU=DOI Users,DC=DOI,DC=NYCNET"){
                        $groups = @('Digital-Forensics-Unit', 'Investigation_Division_EDOCS', 'Forensics', 'InvestigationDivision')
                        foreach($group in $groups){
                            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
                        }
                    }
                    elseif ($OU -eq "OU=Facilities,OU=DOI Users,DC=DOI,DC=NYCNET"){
                        $groups = @('FacilitiesStaff')
                        foreach($group in $groups){
                            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
                        }
                    }
                    elseif ($OU -eq "OU=Finance,OU=DOI Users,DC=DOI,DC=NYCNET"){
                        $groups = @('Fiscal', 'FiscalServices', 'RightFax_Finance-Fiscal-Procurement')
                        foreach($group in $groups){
                            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
                        }
                    }
                    elseif ($OU -eq "OU=Fingerprint,OU=DOI Users,DC=DOI,DC=NYCNET"){
                        $groups = @('FingerprintStaff', 'Fingerprint_Share', 'RightFax_Fingerprint')
                        foreach($group in $groups){
                            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
                        }
                    }
                    elseif ($OU -eq "OU=Interns,OU=DOI Users,DC=DOI,DC=NYCNET"){
                        $groups = @('Interns')
                        foreach($group in $groups){
                            Add-ADGroupMember -Identity $group -Members @($Username) -Credential $credentials
                        }
                    }
                    
                    #Adds General Membership Groups
                    $groups = @('Azure_MFA_PasswordReset', 'MFA_PaloAlto', 'RightFaxUsers', 'VPNUsers', 'M365LicenseGroup', 'MFALicenseGroup')
                    foreach($group in $groups){
                        Add-ADGroupMember `
                            -Identity $group `
                            -Members @($Username) `
                            -Credential $credentials
                    }

                    #Connects to Azure and performs an ADsync
                    Invoke-Command -ComputerName doiazad -Credential $credentials -ScriptBlock {
                        Start-ADSyncSyncCycle -PolicyType Delta
                    }
                    
                    #Pauses script for 30 seconds
                    Start-Sleep -Seconds 30

                    #Connects to the Exchange Server and Enables Remote mailbox 
                    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://doiexchmb01.doi.nycnet/PowerShell/ -Authentication Kerberos -Credential $credentials
                    Import-PSSession $Session -DisableNameChecking

                    Enable-RemoteMailbox -Identity $Name -RemoteRoutingAddress $Username@nycdoi365.mail.onmicrosoft.com
                }

                #Removes Connection to Exchange Server
                Remove-PSSession $Session
            }
    }
    else {
        exit
    }
}
elseif ($user_answer.ToLower() -eq "n"){
    $user_answer_three = Read-Host -Prompt "Are you Departing a User Account (Y or N)"
    if ($user_answer_three.ToLower() -eq "y"){
        Invoke-Command -ComputerName "doidc02" -Credential $credentials -ScriptBlock{
            $login = $using:login
            $login_name = Get-ADUser -Identity $login
            $From = $login_name.UserPrincipalName
            $EmailTo = "desktoptechs@doi.nyc.gov", "SecurityAlert@doi.nyc.gov", "garroyo@doi.nyc.gov"
            
            #Ask for Terminated useraccount, check to make sure the username is active and not already departed.
            $credentials = $using:credentials
            
            $validusername = $true
            while ($validusername){
                $username_test = Read-Host -Prompt "Please enter the username of the terminated account"
                try {
                    $username_details = Get-ADUser -Identity $username_test -ErrorAction Stop
                    $name_string = $username_details.Name.ToString()
                    if ($username_details.distinguishedName -eq "CN=$name_string,OU=Departed Users,DC=DOI,DC=NYCNET"){
                        Write-Host "The user $name_string is already departed." -ForegroundColor Red
                        $choice = Read-Host "Would you like to try another username? (Y/N)"
                        if ($choice -eq 'N' -or $choice -eq 'n'){
                            exit
                        }else{
                            continue
                        }
                    }
                    $username = $username_details.SamAccountName
                    $validusername = $false
                    
                } catch {
                    Write-Host "The username '$username_test' does not exist." -ForegroundColor Red
                    $choice = Read-Host "Would you like to try another username? (Y/N)"
                    if ($choice -eq 'N' -or $choice -eq 'n'){
                        exit
                    }
                }
            }
    
            #Specify a time the script will run. 
            $time = Read-Host -Prompt "What time would you like to disable the account?"
            $targetTime = [datetime]$time
            $buffer = [timespan]::FromMinutes(5)
            $currentTime = Get-Date
    
            #Verify the Account Termination
            $account_name = $username_details.Name
            $username_verify = Read-Host -Prompt "Are you sure you want to Terminate the following user? (Y/N) $account_name"
            if ($username_verify -eq 'Y' -or $username_verify -eq 'y'){
                
            }else{
                exit
            }
    
            while ($currentTime -lt $targetTime -or $currentTime -gt ($targetTime + $buffer)) {
                Start-Sleep -Seconds 5  
                $currentTime = Get-Date
            }
    
            #Reset Password
            Set-ADAccountPassword -Identity $username -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "Welcome!@#" -Force)
    
            #Assigned memberships
            $assignedgroups = Get-ADPrincipalGroupMembership -Identity $username | Select-Object Name | Out-String
    
            #Disable user account
            Disable-ADAccount -Identity $username -Credential $credentials
    
            #clear the Manager and Direct report fields
            Set-ADUser -Identity $username -Clear Manager -Credential $credentials
            $directreports = Get-ADUser -Identity $username -properties DirectReports | select-object -ExpandProperty DirectReports
            foreach($user in $directreports){
                Set-ADUser -Identity $user -Clear Manager -Credential $credentials
            }
    
            #Remove all memberships from AD account
            $membershipgroups = Get-ADPrincipalGroupMembership -Identity $username
    
            foreach ($membership in $membershipgroups){
                if ($membership.distinguishedName -eq 'CN=Domain Users,OU=General SG,OU=Security Groups,OU=Groups,DC=DOI,DC=NYCNET')
                {
                continue
                }
                Remove-ADPrincipalGroupMembership -Identity $username -MemberOf $membership.distinguishedName -Credential $credentials -Confirm:$false
            }
    
            #Move AD account to Departed User's OU
            $username_details = Get-ADUser -Identity $username
            Move-ADObject -Identity $username_details.distinguishedName -TargetPath 'OU=Departed Users,DC=DOI,DC=NYCNET' -Credential $credentials
    
            #Move the Home and Profile folders to the Archive server. 
            $Folder_Name = $username
            $Path1 = "\\doiarchive01\home_archive\$Folder_Name"
            New-Item -Path $Path1 -ItemType Directory 
            $Path2 = "\\doiarchive01\profile_archive\$Folder_Name"
            New-Item -Path $Path2 -ItemType Directory 
    
            $Source_Home_Folder = "\\doi.nycnet\doi_share\home_folder\$Folder_Name"
            $Destination_Home_Folder = "\\DOIARCHIVE01\HOME_ARCHIVE\$Folder_name"
    
            $Source_Profile_folder = "\\DOIPROFILE01\USER_FOLDER_REDIRECTION\$Folder_name"
            $Destination_Profile_folder = "\\DOIARCHIVE01\PROFILE_ARCHIVE\$Folder_name"
    
            #Robocopy Execute
            robocopy $Source_Home_Folder $Destination_Home_Folder /COPYALL /Z /E /W:1 /R:2 /tee /Move 
            robocopy $Source_Profile_folder $Destination_Profile_folder /COPYALL /Z /E /W:1 /R:2 /tee /Move 
    
            #Sends Email with user's memberships
            $fullname = $username_details.Name
            Send-MailMessage -From $From -To $EmailTo -Subject "Departed User $fullname" -body "The Departed account $fullname is now completed. Their home and profile folders have been moved to the Archived Server. Here is a list of Group Memberships he/she was assigned to: `n$assignedgroups" -SmtpServer 'smtp.doi.nycnet' -Port '25'
        }
    }
    else{
        exit
    }    
}
