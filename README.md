# Office365-Info
Office 365  Service health, usage reports and monitoring.
Screenshots available: https://jc-nts.blogspot.com/2019/06/office-365-service-health-reporting-on.html

This little project initially started out as a web page or two to show Office 365 service health information. It quickly grew into a suite of powershell scripts which did similar, but distinct, tasks.

# Dashboard
This is the main script. It display office 365 service health information on various tabs of a web page.
## Overview
Highlight the most critical workloads and display these as colour coded 'cards'
Open and recent incidents are listed, with a full workload list and 'traffic light' icons indicating health status 
## Incidents
A full list of closed incidents
## Messages
A list of messages - Plan / Fix, Plan for Change, and no-action required messages
## Diagnostics
some simple diagnostics to verify Office 365 connectivity (DNS lookups, web page access etc)

# Monitor
Monitor simply checks for new/closed incidents on office 365 service health and sends an email.

# Usage
Usage reports download the Office 365 usage reports in CSV format

# Wall
This script produces a single page dashboard of all the office 365 workloads available in your tenant and their features.
Those with issues are highlighted


I'm hoping to have more functionality soon
