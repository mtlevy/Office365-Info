# O365-SH
## Office365 Service Health Tools
All services use an Azure AD App to fetch Office365 Service Health information

## Office 365 Service Health Dashboard (Dashboard\O365SH-Dashboard.ps1)
### Overview:
	A multipage dashboard showing your favourite workloads and status.
	Active incidents are displayed along with recently closed incidets.
	A full list of existing workloads with a traffic light indicator showing current status

### Features:
	A version of the 'Wall' showing all of your workloads and features. Colour coded to show outages and health,
	hover-over gives a description of the status.
			
### Incidents:
	A list of all closed incidents. Links are provided to local versions of the incident pages.
	All documents can be rebuilt using the -rebuilddocs parameter

### Advisories:
	Similar to incidents, these advisories are linked to local docs.
	Advisories are broken down into 'Prevent / Fix Issues', 'Plan for Change' and 'Other Messages''

### Roadmap:
	Pulls the Microsoft 365 Roadmap and Azure Updates into one place

### Log:
	Simple logging about the number of items downloaded.
	Also has links to the Information Wall and Diagnostics page.

## Office 365 Diagnostics (Toolbox\O365SH-Toolbox.ps1)
### Diagnostics:
	Some basic client connectivity checks

### Licences:
	Lists all licence SKUs and sub-SKUs available on your tenant, and their provisioning status.

### IP and URLs:
	Microsoft URL and IP changes for the Office 365 environment.

### URLs
	A single page display showing all of your Office 365 Workloads, Features and their status.

### CNAMES
	Track CNAME changes and lookups from multiple DNS resolvers

### Logs
	Simple logging about the number of items downloaded.


## Office365 Monitor (Monitor\O365SH-Monitor.ps1)
A simple script that checks for outages and sends an email alert when new incidents are detected.
Incidents can be logged to the event log.
Changes to CNAMEs are monitored and new results are logged.

## Office365 Usage Reports (Usage\O365SH-Usage.ps1)
A simple script to download the Office 365 Usage reports.
Can be scheduled to run on a regular basis and download the appropriate reports for available periods (D7, D30 etc)
