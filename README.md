# SplunkO365MessageTrace

This is a PS Script created in PS 5.
It connects to Office 365 and utilizes the Get-MessageTrace command let to collect Message Logs and converts them to JSON. 
A Splunk forwarder can monitor the output directory and send those logs to Splunk. 

There is a SplunkO365 Initialization function that prompts for username/password and stores them in the registry as a secure string. 
All registry items can be found in "HKCU:\Software\O365"

The task needs to be created with Triggers enforcing persistence.
I have mine set with At Start Up and at 11:00AM. I have them both set to repeat every 1 minute indefinately. 
Under settings I have the "Do not start a new instance" enabled. This prevents the script from running multiple instances. 
Create an Action to start a program. The program is powershell.exe with the following arguments:
"-executionpolicy bypass C:\Scripts\Office365\Splunk_o365_messagetrace.ps1"

