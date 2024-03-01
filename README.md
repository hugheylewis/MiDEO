# mde-offboarder
A Python script to automate the offboarding of devices from Microsoft Defender for Endpoint (MDE)

This implementation currently takes an exported CSV file from Defender for Endpoint (security.microsoft.com) as the input and will offboard <b>all machines in that file</b>.<br>
The offboarded devices are then inserted into a SQLite database, stored in the current working directory.<br>
*** It is important to leave the columns and the format of the CSV file from Defender as-is. main.py is expecting to read the list of device IDs and hostnames from columns 0 and 1, respectively. Rows, however, can be removed as needed to prevent offboarding specific endpoints.***

# Getting Started
<h1>Azure App Registration</h1>
You will need to create an App Registration with Microsoft Azure. Follow the instructions below on how to generate your app's secret keys and granting the appropriate application permissions.
<ol>
<li>Follow this Microsoft guide on creating an app registration. When the reach the API permissions section, move on to Step 2. https://learn.microsoft.com/en-us/azure/healthcare-apis/register-application</li>
<li>The following API permissions are required for this application: `Machine.Read.All` and `Machine.Offboard`</li>
<li>Take note of your Azure tenant ID, the app ID and the app secret (store this somewhere safe: if you lose it, you will have to generate a new one and break any pre-existing connections with this app)</li>
<li>Clone this repo to your location machine. Navigate to `config/.env` and paste your tenant ID, app ID and app secret in the appropriate fields</li>
</ol>
<h1>pip installs</h1>
The only package required to be installed is `dotenv`, a Python module that assists this project with the secure handling of API keys, application secret keys, etc.<br>
&#96;pip install python-dotenv&#96;

# Roadmap / TODO
1.  Azure Key Vault to replace config.py and .env files (used only for local script execution)
