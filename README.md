# mde-offboarder
A Python script to automate the offboarding of devices from Microsoft Defender for Endpoint (MDE)

This implementation currently takes an exported CSV file from Defender for Endpoint (security.microsoft.com) as the input and will offboard <b>all machines in that file</b>.
*** It is important to leave the columns of the CSV file as-is. main.py is expecting to read the list of device IDs and hostnames from columns 0 and 1, respectively.


# TODO
1. Clean-up the final API response to only show relevant fields to the user
