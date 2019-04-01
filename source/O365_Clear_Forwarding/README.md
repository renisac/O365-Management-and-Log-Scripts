## O365_Clear_Forwarding.ps1

### Use case

Use this PowerShell script to set O365 mailbox forwarding to NULL.  Once an attacker has compromised an O365 account, sometimes they will setup forwarding to receive a copy of the victims e-mail.  If you find a "funny" looking forward, use this script to clear the forward.

### Brief description

Run script and enter O365 admin creds. This PowerShell script uses Exchange Online Actions, changes the mailbox forwarding property to NULL and outputs data to the screen. 

### Data sensitivity
None

### Local environment requirements
None

### License and/or permissions required

A1 or higher

### How to run the script
```` powershell
.\O365_Clear_Forwarding.ps1
````

### Known issues
None

### Acknowledgements and Credits
Josh Hartley/University of Missouri
