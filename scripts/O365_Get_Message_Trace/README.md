## O365_Get_Message_Trace.ps1

### Use case

Use this PowerShell script to pull O365 message tracking info.  If you need to pull messages sent by a specific sender or recipient to determine whether they have sent or received suspicious email, use this script.

### Brief description

Run script and enter O365 admin creds. This PowerShell script uses Inbound/Outbound mail actions to pull e-mail flow meta data and outputs a csv file to the Desktop. 

### Data sensitivity
E-mail Meta Data: 
- To
- From
- Subject
- IP Addresses

### Local environment requirements
None

### License and/or permissions required

A1 or higher

### How to run the script
```` powershell
.\O365_Get_Message_Trace.ps1
````

### Known issues
None

### Acknowledgements and Credits
Josh Hartley/University of Missouri
