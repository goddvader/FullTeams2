###############################################################################
# Complete Teams adminy script
# Assign & Unassign Phonenumbers for Users & ressource Accs
# 22.08.2021 Initial Version - Roman Padrun
###############################################################################
# Requires -Modules MicrosoftTeams
Import-Module MicrosoftTeams

#$cred = Get-Credential
Connect-MicrosoftTeams

#Function menu to choose the function u want
function Show-Menu
{
    param (
        [string]$Title = "Teams Functions by isolutions"
    )
    Clear-Host
    Write-Host "================= $Title ================="

    Write-Host "Add Number to a User:                   Press `1` for this option."
    Write-Host "Add Number to a Ressource Account:      Press `2` for this option."
    Write-Host "Remove Number from a User:              Press `3` for this option."
    Write-Host "Reassign Number from a User:            Press `4` for this option."
    Write-Host "Show all assgined Numbers:              Press `5` for this option"
    Write-Host "Massmigrate Users to Teams:             Press `6` for this option"
    Write-Host "End the program:                        Press `q` for this option."
}

do
    {
        Show-Menu
        $selection = Read-Host "Please choose the option"
        switch ($selection)
        {
            '1' {
                'You choose option 1'
                Clear-Host
                Write-Host "Please enter the Username and Phonenumber to assign and specify the Teams Policy"
                Write-Host "If Swisscom: SwisscomET4T" -ForegroundColor Blue
                Write-Host "If Sunrise: SunriseUnlimited" -ForegroundColor Red

                $username =  read-Host "Username"
                $phonenumber = Read-Host "Phonenumber (Example: +413424232)"
                $policy = Read-Host "Policy"

                if ($policy -eq 'SunriseUnlimited' -and "SwisscomET4T"){
                    Set-CsUser -Identity $username -EnterpriseVoiceEnabled $true -HostedVoiceMail $true -OnPremLineURI tel:$phonenumber
                    Grant-CsOnlineVoiceRoutingPolicy -Identity $username -PolicyName $policy
                    #Grant-CsTeamsUpgradePolicy –Identity $username –PolicyName UpgradeToTeams

                    Write-Host "$username has successfully $phonenumber assigned" -ForegroundColor Green
                    }
                        else{
                        Write-Host "You have not choose a policy" -ForegroundColor Red
                    }
            } '2' {
                'You choose option 2'
                Clear-Host
                Write-Host "Please enter the Autoaattendant and the Phonenumber to assign"

                $AutoAttendant = Read-Host "AutoAttendant"
                $phonenumber2 = Read-Host "Phonenumber"

                Set-CsOnlineApplicationInstance -Identity $AutoAttendant -OnpremPhoneNumber $phonenumber2

                Write-Host "$AutoAttendant has successfully $phonenumber assigned" -ForegroundColor Green

            } '3' {
                'You choose option 3'
                Clear-Host
                Write-Host "Please enter the Username where you want to remove the phonenumber"

                $username3 =  read-Host "Username"

                Set-CsUser -Identity $username3 -EnterpriseVoiceEnabled $true -HostedVoiceMail $true -OnPremLineURI $null

                Write-Host "$username3 has successfully the number removed" -ForegroundColor Green
                
            } '4' {
                'You choose option 4'
                Clear-Host
                Write-Host "Please enter the Username where you want to reassign the phonenumber"

                $username4 = read-Host "Username"
                $phonenumber4 = Get-CsOnlineUser | Where-Object { $_.UserPrincipalName -Match "$username4"} | Format-Table LineURI -HideTableHeaders
                
                Set-CsUser -Identity $username4 -EnterpriseVoiceEnabled $true -HostedVoiceMail $true -OnPremLineURI $null
                
                Start-Sleep -s 5
                
                Set-CsUser -Identity $username4 -EnterpriseVoiceEnabled $true -HostedVoiceMail $true -OnPremLineURI $phonenumber4
                
                Write-Host "$username4 has successfully $phonenumber4 reassigned" -ForegroundColor Green
            } '5' {
                'You choose option 5'
                $allnumbers = Get-CsOnlineUser | Where-Object  { $_.LineURI -notlike $null } | Format-Table DisplayName,UserPrincipalName,LineURI
                $allnumbers
                pause

            } '6' {
                'You choose option 6'
                Write-Host "The CSV File need to look like this -> name;phone;policy"
                $csvpath = Read-Host "please enter the Path to the CSV file"

                if ($csvpath -eq '0'){
                    $userfull = Import-Csv -Path $csvpath -Delimiter ";"
                    foreach($user in $userfull){
                        $username6 = $($user.name)
                        $phonenumber6 = $($user.phone)
                        $policy6 = $($user.polciy)
                        Set-CsUser -Identity $username6 -EnterpriseVoiceEnabled $true -HostedVoiceMail $true -OnPremLineURI $phonenumber6
                        Grant-CsOnlineVoiceRoutingPolicy -Identity $username6 -PolicyName $policy6
                        Grant-CsTeamsUpgradePolicy –Identity $username6 –PolicyName UpgradeToTeams
                        Write-Host "$username6 successfully assigned the $phonnumber6" -ForegroundColor Green
                    }
                }
                    else {
                        Write-Host "$csvpath is no valid path"
                    }
                pause

        }
        
    }
    }
    until ($selection -eq 'q')