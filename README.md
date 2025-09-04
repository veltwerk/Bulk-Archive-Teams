# Bulk-Archive-Teams

- Sjoerd Derks : s.derks@veldwerk.nl
- Michiel In 't Velt : m.intvelt@veldwerk.nl

## disclaimer

You need several admin roles in your tenant to use this script. If you have those, then you should know what your doing when running a bulk action powershell script against your Azure tenant. Still... use at your own risk.

## what the script is for

In educational tenants lots and lots of class teams / educational teams have to be created each year and those teams have to be archived after the schoolyear is over. This script uses the Graph API and Exchange Online Management powershell modules to help with the archiving. 

It lets you select the teams to be archived (and removed from the Exchange Address List) with the option to also remove all members from those teams.

## prerequisites

- An admin account for your tenant with at least the following roles : Exchange Admin + Groups Admin + Teams Admin + Application Admin **or** Exchange Admin + Global Admin.
- The admin account must have a valid Teams license. A1 is sufficient.
- Powershell. Version **5.1** works best because in higher versions, filtering in the gridview when selecting teams, can crash the process.
- ExchangeOnlineManagement, Microsoft.Graph.Groups, Microsoft.Graph.Users, Microsoft.Graph.Teams powershell modules installed. Unquote the _prerequisites_ section to run module installers with the script.

## how to use it

Run the script without any parameters. It will show a popup (powershell 5.1) or a browser tab to log onto graph and then, once more, to log onto exchange online.

- It will ask you to enter a search term that will be used to filter on the _mailNickname_. When Veldwerk creates classTeams for a schoolyear, we allways put the schoolyear in the mailNickname of the team/group. For example: search on "2526" to retrieve teams like: `VBS-2526-H3A-NE@domain.edu`

- It will ask you to show only class Teams; default = **yes**. This filters the groups to only show teams with `Visibility = HiddenMembership'. In most tenants; only the educational teams, i.e. teams with a class notebook and other edu apps, have visibility set to hide the members.

- The script will ask if you want to remove all members from the selected teams; default = **no**.

_In large tenants with 1000+ teams, it takes a while to retrieve all the teams._

- You are presented with a gridView where you can filter (if running on powershell 5.1) to find the teams you need to archive.
Click on `OK` to start the process.

After the archiving and member removal processes are finished, a .CSV file will be writen to your home folder and opened with whatever application is your default for such files.

## notes

- There is minimal error handling in this script.
- The script doesn't check to see if a team is already archived. When an already archived team is processed you will see an error flashing by along with a warning from exchange that the call `Set-UnifiedGroup -Identity $selectedGroup.Id -HiddenFromAddressListsEnabled:$true` has altered no properties.
- _**Be patient: archiving 1000 teams and removing all members can take 2 hours or more!**_
