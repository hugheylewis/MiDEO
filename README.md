# mde-offboarder
A Python script to automate the offboarding of devices from Microsoft Defender for Endpoint (MDE)

This implementation currently only supports the offboarding of a single device at a time. Future revisions of this script will allow multiple devices - either manually entered by the user or uploaded via CSV file - to be offboarded at a time.

# TODO
1. Allow users to upload .csv file of hostnames or IP addresses of machines to offboard.
2. Clean-up the final API response to only show relevant fields to the user
