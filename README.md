# sling-to-gcal (DEPRECATED)

Deprecated Python script to import Sling .xls shift schedules into Google Calendar. This W.I.P. script was discontinued as my organization did not renew their contract with Sling, instead migrating to another platform for shift scheduling.

## Description

Sling is an employee scheduling app. It offers no easy way to export those shifts into another calendar app. I have had to put shifts into Google Calendar manually for a while now, and it has become a cumbersome process.

I decided, at my peer's suggestion, that I would sit down and write a script to automate this process. For anyone else using Sling, I hope this will be helpful. Thanks for viewing the page.

## Usage

First, to use this script, you'll need to set up the Google Calendar API to work with the script. Google has a guide for that [here.](https://developers.google.com/calendar/api/quickstart/python#enable_the_api)
This will involve going to Google Cloud, making a new project (name it whatever you want), configuring OAuth, and getting a credentials.json file. Just place that credentials.json file in the same directory as the script when you're done.

Second, you must download export a .xls file from the Sling month view. Place that shifts-export.xls file into the same directory as the script, and then you can run the script! It'll change the .xls file to a .xlsx, and then process your data.
***Do not, under any circumstances, change the name of the .xls or .xlsx file!!***

Note that on the first run the script will also ask you to log in to your Google account. If you're using your own Google Cloud project, I can't see the data the script handles nor access your Google account.
It's perfectly safe as long as you're only using your own credentials.json.
You'll also need to install the two packages I've listed in the **Requirements** section.

## Requirements

> the [openpyxl](https://openpyxl.readthedocs.io/en/stable/tutorial.html) package
> the [Google Client Library](https://developers.google.com/calendar/api/quickstart/python#install_the_google_client_library) package
> a [Google Cloud Project](https://developers.google.com/workspace/guides/create-project)
> a month's shifts-export.xls file from [Sling](https://app.getsling.com/shifts)
> an installation of Python...
> and a Google Account with a calendar!
