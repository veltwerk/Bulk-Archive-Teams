<#
# v1.5 :
    Added option to skip archiving and only remove members
    Added optional domain filter to only remove members matching a domain (e.g. students)

# v1.4 : 
    Minor text changes
    Added option to filter on class teams only (Visibility = HiddenMembership)
    Added variable $teamsType to show in output whether filtering on class teams or not
    Changed wording of some user prompts
    
# v1.3 : 
    Moved all user input to the start of the script so it can finish without further user interaction
    double progress bar showing progress for both member removal and team archiving
    suppression of displaying member names

# V 1.2 : Script checks if the Set-UnifiedGroup command is available; if not, the logged-in user does not have the Exchange admin role
#>
# Script designed to search, select and archive Teams groups

Write-Host "This script allows you to bulk archive Teams groups and hide them from the address lists." -ForegroundColor Yellow
Write-Host "You have the option to also remove all members after archiving." -ForegroundColor Yellow

# Prerequisites

    <# Install and import ExchangeOnlineManagement module
    if (-not(Get-InstalledModule Microsoft.Graph.Teams -ErrorAction SilentlyContinue)){
        Write-Host "ExchangeOnlineManagement module installeren..." -ForegroundColor Yellow
        Install-Module ExchangeOnlineManagement -Confirm:$false -Force -Scope CurrentUser
    }
    Write-Host "Installatie ExchangeOnlineManagement module voltooid." -ForegroundColor Green

    if (-not(Get-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue)){
        Write-Host "ExchangeOnlineManagement module importeren..." -ForegroundColor Yellow
        Import-Module ExchangeOnlineManagement
    }
    Write-Host "Importeren ExchangeOnlineManagement module voltooid." -ForegroundColor Green

    # Install and import Microsoft.Graph.Teams module
    if (-not(Get-InstalledModule Microsoft.Graph.Teams -ErrorAction SilentlyContinue)){
        Write-Host "Microsoft.Graph.Teams module installeren..." -ForegroundColor Yellow
        Install-Module Microsoft.Graph.Teams -Confirm:$false -Force -Scope CurrentUser
    }
    Write-Host "Installatie Microsoft.Graph.Teams module voltooid." -ForegroundColor Green

    if (-not(Get-Module Microsoft.Graph.Teams -ErrorAction SilentlyContinue)){
        Write-Host "Microsoft.Graph.Teams module importeren..." -ForegroundColor Yellow
        Import-Module Microsoft.Graph.Teams
    }
    Write-Host "Importeren Microsoft.Graph.Teams module voltooid." -ForegroundColor Green

    # Install and import Microsoft.Graph.Groups module
    if (-not(Get-InstalledModule Microsoft.Graph.Groups -ErrorAction SilentlyContinue)){
        Write-Host "Microsoft.Graph.Groups module installeren..." -ForegroundColor Yellow
        Install-Module Microsoft.Graph.Groups -Confirm:$false -Force -Scope CurrentUser
    }
    Write-Host "Installatie Microsoft.Graph.Groups module voltooid." -ForegroundColor Green 

    if (-not(Get-Module Microsoft.Graph.Groups -ErrorAction SilentlyContinue)){
        Write-Host "Microsoft.Graph.Groups module importeren..." -ForegroundColor Yellow
        Import-Module Microsoft.Graph.Groups
    }
    Write-Host "Importeren Microsoft.Graph.Groups module voltooid." -ForegroundColor Green
    
    # Install and import Microsoft.Graph.Users module
    if (-not(Get-InstalledModule Microsoft.Graph.Users -ErrorAction SilentlyContinue)){
        Write-Host "Microsoft.Graph.Users module installeren..." -ForegroundColor Yellow
        Install-Module Microsoft.Graph.Users -Confirm:$false -Force -Scope CurrentUser
    }
    Write-Host "Installatie Microsoft.Graph.Users module voltooid." -ForegroundColor Green 

    if (-not(Get-Module Microsoft.Graph.Users -ErrorAction SilentlyContinue)){
        Write-Host "Microsoft.Graph.Users module importeren..." -ForegroundColor Yellow
        Import-Module Microsoft.Graph.Users
    }
    Write-Host "Importeren Microsoft.Graph.Users module voltooid." -ForegroundColor Green
    #>
    
    Add-Type -AssemblyName System.Windows.Forms
    Import-Module ExchangeOnlineManagement
    Import-Module Microsoft.Graph.Groups
    Import-Module Microsoft.Graph.Users
    Import-Module Microsoft.Graph.Teams

    #Disconnect any existing sessions
    Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue > $null
    Disconnect-MgGraph -ErrorAction SilentlyContinue > $null
    
    # Connect to tenant
    # EXO eerst: Microsoft.Graph laadt een oudere Microsoft.Identity.Client.dll die de WAM-broker van EXO breekt
    Write-Host "Connecting to tenant..." -ForegroundColor Yellow
    $exoParams = @{ ShowBanner = $false; ErrorAction = 'Stop' }
    if ((Get-Command Connect-ExchangeOnline).Parameters.ContainsKey('DisableWAM')) { $exoParams['DisableWAM'] = $true }
    Connect-ExchangeOnline @exoParams
    Connect-MgGraph -Scopes "TeamSettings.ReadWrite.All, Group.ReadWrite.All, User.Read.All" -ErrorAction Stop
    Write-Host "Connection to tenant established." -ForegroundColor Green

    Write-host "Checking if Exchange module is complete."
    # check if the Set-UnifiedGroup command is available
    if (-not (Get-Command Set-UnifiedGroup -ErrorAction SilentlyContinue)) {
        Write-Host "Set-UnifiedGroup cmdlet is not available. Ensure the logged-in user has the Exchange Admin role!" -ForegroundColor Red
        return
    }

    # Filter Teams groups
    # Input search term
    Write-Host "`n`nEnter a search term that will be used to filter on MailNickname. Press [ENTER] for no filter." -ForegroundColor Yellow
    Write-Host "[note] You will be able to search and filter in the gridview after the teams have been retrieved."
    $queryGroups = Read-Host

    # Ask in advance for class teams only
    # Input search term
    $classTeamsOnlyInput = Read-Host "Show class / education teams only (Visibility = HiddenMembership)? (Y/n)"
    $classTeamsOnly = if ($classTeamsOnlyInput -match '^(n|no)$') { "no";$teamsType="" } else { "yes";$teamsType="class" }

    # Ask in advance if the selected teams should be archived
    $doArchiveInput = Read-Host "Should the selected teams be archived (and hidden from address lists)? (Y/n)"
    $doArchive = if ($doArchiveInput -match '^(n|no)$') { "no" } else { "yes" }

    # Ask in advance if members should be removed
    $removeMembersInput = Read-Host "Should team members be removed after archiving? (y/N)"
    $removeMembers = if ($removeMembersInput -match '^(y|yes)$') { "yes" } else { "no" }

    # Ask for an optional domain filter to only remove members with a matching UPN/mail domain (e.g. students)
    $memberDomainFilter = ""
    if ($removeMembers -eq "yes"){
        $memberDomainFilter = Read-Host "Enter a domain name to only remove members with that domain (e.g. student.domain.nl). Press [ENTER] to remove all members"
    }

    if ($queryGroups) { $preFilter="with filter '$queryGroups'"} else { $preFilter="without filter" }
    Write-host "Retrieving $teamsType Teams $preFilter..."

    # Get all Teams groups
    if ("no" -eq $classTeamsOnly) {
        $allGroups = Get-MgGroup -All | Where-Object {$_.Team -ne $null } | Select-Object -Property DisplayName, MailNickname, Id
    }else{
        $allGroups = Get-MgGroup -All | Where-Object {$_.Team -ne $null -and $_.Visibility -eq "HiddenMembership"} | Select-Object -Property DisplayName, MailNickname, Id
    }

    # Output
    $queriedGroups = $allGroups | Where-Object {$_.MailNickname -like "*$queryGroups*"}
    # [System.Windows.Forms.MessageBox]::Show("Select the Teams you want to archive.")
    $selectedGroups = $queriedGroups | Out-GridView -OutputMode Multiple -Title "Select Teams to Archive"

    $groupCount = $selectedGroups.Count

# Archive selected groups
foreach ($selectedGroup in $selectedGroups) {
    if ($doArchive -eq "yes"){
        Write-Progress -Activity "Archiving teams" -Status "Archiving $($selectedGroup.DisplayName)" -PercentComplete (($selectedGroups.IndexOf($selectedGroup) / $groupCount) * 100)
        Set-UnifiedGroup -Identity $selectedGroup.Id -HiddenFromAddressListsEnabled:$true 
        Invoke-MgArchiveTeam -TeamId $selectedGroup.Id -Confirm:$false 
    }
    # Option to remove members
    if ($removeMembers -eq "yes"){
        Write-Progress -Activity "Archiving teams" -Status "Removing members from $($selectedGroup.DisplayName)" -PercentComplete (($selectedGroups.IndexOf($selectedGroup) / $groupCount) * 100)
        $users = Get-MgGroupMember -All -GroupId $selectedGroup.Id | Select-Object -Property Id, AdditionalProperties
        # Filter members by mail/UPN domain when a filter was provided (e.g. only remove students)
        if ($memberDomainFilter) {
            $users = $users | Where-Object {
                $upn = $_.AdditionalProperties["userPrincipalName"]
                $upn -and $upn -like "*@$memberDomainFilter"
            }
        }
        if ($users -and $users.Count -and $users.Count -gt 0) {
            $usercount = $users.count
        } else {
            $usercount = 1 # error voorkomen met progressbar wanneer er maar 1 lid is
        }
        $curUser = 0
        foreach ($user in $users){
            $curUser++
            Write-Progress -Activity "Removing members" -id 1 -Status "User $($curUser) / $($usercount)" -PercentComplete (($curUser / $usercount) * 100)
            # Get-MgUser -UserId $user.Id | Select-Object -Property DisplayName, Id 
            Remove-MgGroupMemberByRef -GroupId $selectedGroup.Id -DirectoryObjectId $user.Id > $null
        }
        Write-Progress -Activity "Removing members" -id 1 -Status "Complete" -Completed
    }
}
if ($removeMembers -eq "no"){
    Write-Host "No members removed." -ForegroundColor Red
} elseif ($memberDomainFilter) {
    Write-Host "Only members with domain '$memberDomainFilter' were removed." -ForegroundColor Yellow
}
if ($doArchive -eq "no"){
    Write-Host "Teams were not archived." -ForegroundColor Red
}
Write-Progress -Activity "Archiving teams" -Status "Complete" -Completed

if ($selectedGroups.Count -gt 0){
    # Export edited Teams to CSV
    $selectedGroups | Export-Csv -Path $HOME\archiveTeams.csv
    Invoke-Item -Path $HOME\archiveTeams.csv
}

Disconnect-ExchangeOnline -Confirm:$false
Disconnect-MgGraph 