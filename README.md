# Bulk-Archive-Teams

- Sjoerd Derks : s.derks@veldwerk.nl
- Michiel In 't Velt : m.intvelt@veldwerk.nl

## disclaimer

You need several admin roles in your tenant to use this script. If you have those, then you should know what your doing when running a bulk action powershell script against your Azure tenant. Stil... use at your own risk.

## what the script is for

In educational tenants lots and lots of class teams / educational teams have to be created each year and those teams have to be archived after the schoolyear is over. This script uses the Graph API and Exchange Online Management powershell modules to help with the archiving. 

It lets you select the teams to be archived (and removed from the Exchange Address List) with the option to also remove all members from those teams.

## prerequisites

- An admin account for your tenant with at least the following roles : Exchange Admin + Groups Admin + Teams Admin + Application Admin **or** Exchange Admin + Global Admin. Note that the Exchange Admin role is needed even if you are global admin.
- The admin account must have a valid Teams license. A1 is sufficient.
- Powershell. Version **5.1** works best because in higher versions, filtering in the gridview when selecting teams, can crash the process.
- ExchangeOnlineManagement, Microsoft.Graph.Groups, Microsoft.Graph.Users, Microsoft.Graph.Teams powershell modules installed.

## how to use it

Run the script without any parameters. It will show a popup (powershell 5.1) or a browser tab to log onto graph and then again, to log onto exchange online.

It will ask you to enter a search term that will be used to filter on the _mailNickname_. When Veldwerk creates classTeams for a schoolyear, we allways put the schoolyear in the mailNickname of the team/group; `VBS-**2526**-H3A-NE`
