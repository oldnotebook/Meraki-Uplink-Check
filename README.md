# Meraki-Uplink-Check
-The script sends us a list of Meraki devices that have uplinks listed as ‘failed’ or ‘not connected’ for over 45 minutes at a minimum.   This was done since Meraki doesn’t have an automatic alert to let you know if a backup wan/uplink port is down or failed.  
-This was designed to be run on a Windows server.  Solarwinds may be a good candidate, for example.
-This is for Python version 3.7.  

# Curl Setup
-Curl needs to be on the machine that runs the scripts
-Curl setup for a Windows box: paste curl.exe, curl-ca-bundle.crt, and libcurl-x64.dll, into c:\Windows\System32.  You'll then be able to run the curl command in the cmd prompt.

# The Scripts
-Uplink status check.py - This is the script that does all the 'magic.'  You'll need to input your meraki API key for it to work.  Curl will also need to be working through the cmd prompt.
	a. Overview:
		i. The script pulls the organization ID from the api then it...
		ii. Pulls the network list with the org ID.  After that, it...
		iii. Pulls the device info for each network ID. Then...
		iv. It'll select the info it needs and save it to a report.  It'll check for any differences from the last time it was run.  If the time is over 45 minutes that an uplink has been down, it'll be added to the alert report.
		v. The final step is to have the report emailed to the email address specified in the EmailUplinkReport.py script.
-EmailUplinkReport.py - This is the script that sends the excel report via email.  You will need to update the SMTP server info and email addresses.

# To run this script on a Windows box via the task scheduler:
1. Create a task that runs whether the user is logged on or not.  Run with highest privileges.
2. Set the trigger to be whenever you want.  I have mine set to daily at noon.
3. The action is to start a program.  The program will be a .bat file. The .bat file kicks off the python script.

After the script runs a second time, you'll be emailed a list of devices that match the alert threshold.   

Created by Nate Revello 2019 January
