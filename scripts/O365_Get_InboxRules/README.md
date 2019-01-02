## O365_Get_InboxRules.ps1

### Use case

Use this PowerShell script to get all O365 inbox rules set on a particular mailbox.  This is a good script to use after an account has been compromised to check for malicious looking inbox rules the attacker may have set.

### Brief description

Run script and enter O365 admin creds. This PowerShell script uses Exchange Online Actions to pull inbox rules configurations on a particular mailbox and outputs data to the screen. 

### Data sensitivity
None

### Local environment requirements
None

### License and/or permissions required

A1 or higher

### How to run the script
```` powershell
.\O365_Get_InboxRules.ps1
````

### Known issues
None

### Acknowledgements and Credits
Josh Hartley/University of Missouri
