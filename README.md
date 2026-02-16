# osm-powershell-tools

Online Scout Manager (OSM) tools written in PowerShell

## Setup

Follow these steps to create an Online Scout Manager (OSM) application which is needed to use the API:

1. Log in to [Online Scout Manager (OSM)](https://www.onlinescoutmanager.co.uk).
2. Expand **Settings** from the left-hand menu
3. Select **My Account Details**
4. Click **Developer Tools** from the left-hand menu
5. Click **Create Application**
6. Enter **osm-powershell-tools** as the application name
7. Click **Save**
8. Close the **Application Created** pop-up (**X** top right)
9. Click **Edit** on the **osm-powershell-tools** application
10. Tick the **Client Credentials Grant** box and click **Save**
11. Click **Regenerate Keys** on the **osm-powershell-tools** application
12. Type **CONFIRM** and click **Confirm**
13. Type **I am a developer** and click **Reveal Credentials**
14. The **OAuth Client ID** and **OAuth Secret** will be displayed **only once**. Make sure to note them down.
15. Close the **Application Keys Regenerated** pop-up (**X** top right)

## Usage

Download the latest release into a directory on your computer & extract.

Open PowerShell in the directory and run `Import-Module .\Osm.PowerShell.Tools.ps1`

First time running? You will be prompted to **Enter your OSM OAuth Client ID** & **Enter your OSM OAuth Client Secret**

Once imported you should see an output to the screen of sections you have access to. You will need this reference table when running the commands.

## Commands

#### New-OsmParentRota

Used to create a new parent rota for the current term for the provided section. Rota will be created in your downloads folder and optionally printed to your default printer.

Create and populate `exclude_{section_name}.txt` (replace {section_name} with relevant value) in your downloads folder with list (one per line) of surnames to exclude from rota.

`New-OsmParentRota -sectionId xxxxx` or `New-OsmParentRota -sectionId xxxxx -print`

#### Get-OsmPaperRegister

Used to get a paper register for the current term for the provided section. Register will be created in your downloads folder and optionally printed to your default printer. If required you can pass the `order` parameter to sort the register by `firstname` or `lastname` or `dob` (age) or `patrolid` (six).

`Get-OsmPaperRegister -sectionId xxxxx` or `Get-OsmPaperRegister -sectionId xxxxx -order firstname` or 
`Get-OsmPaperRegister -sectionId xxxxx -print` or `Get-OsmPaperRegister -sectionId xxxxx -order firstname -print`

#### New-OsmMeetings

Used to create meetings for the current term for the provided section.

`New-OsmMeetings -sectionId xxxxx -day xxxxx` or `New-OsmMeetings -sectionId xxxxx -day mon`

#### Copy-OsmMeetings

Used to copy meetings from one section to another for the current term.

`Copy-OsmMeetings -fromSectionId xxxxx -toSectionId xxxxx -day xxxxx` or `Copy-OsmMeetings -fromSectionId xxxxx -toSectionId xxxxx -day mon`

#### Get-OsmPhotoConsent

Generates a report of members with no photograph consent. Report is saved as an HTML file in the Downloads folder and optionally sent to a printer.

`Get-OsmPhotoConsent -sectionId xxxxx` or `Get-OsmPhotoConsent -sectionId xxxxx -print`

#### Get-OsmDietary

Generates a report of members with allergies & dietary requirements. Report is saved as an HTML file in the Downloads folder and optionally sent to a printer.

`Get-OsmDietary -sectionId xxxxx` or `Get-OsmDietary -sectionId xxxxx -print`

## Contributions

If you find any issues or would like additional features then please open a GitHub issue providing as much detail as possible.

If you can code in PowerShell and would like to contribute then feel free to pick up a GitHub issue, create a branch and develop. Once complete open a PR request to merge into main branch.

## Acknowledgements

Thanks to [@Alan Tiller](https://github.com/alantiller) for his 'osm-for-wordpress' code which I scoured for OSM API info :smiley:
