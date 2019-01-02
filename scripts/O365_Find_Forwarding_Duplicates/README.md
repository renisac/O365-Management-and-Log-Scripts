## O365_Find_Forwarding_Duplicates.ps1

### Use case

Use this PowerShell script to pull All O365 mailbox forwards and find any duplicates.  Once an attacker has compromised an O365 account, sometimes they will setup forwarding to receive a copy of the victims e-mail.

### Brief description

Run script, enter O365 admin creds (Warning: this script can take a couple of hours (or more) to run depending on the number of mailboxes in your tenant). This PowerShell script uses Exchange Online Actions and outputs a CSV file to the Desktop. 

### Data sensitivity
Users’ e-mail addresses

### Local environment requirements
None

### License and/or permissions required

A1 or higher

### How to run the script
```` powershell
.\O365_Find_Forwarding_Duplicates.ps1
````

### Known issues
None

### Acknowledgements and Credits
Josh Hartley/University of Missouri
