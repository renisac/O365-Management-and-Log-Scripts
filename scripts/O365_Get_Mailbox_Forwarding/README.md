## O365_Get_Mailbox_Forwarding

### Use case

Use the this PowerShell script to pull O365 mailbox forwarding info for a particular mailbox.  Once an attacker has compromised an O365 account, sometimes they will setup forwarding to receive a copy of the victims e-mail.  If you find a "funny" looking forward with this script, to can use the O365_Clear_Forwarding script to clear the bad forward.

### Brief description

Run script and enter O365 admin creds. This PowerShell script uses Exchange Online Actions to pull the forwarding configuration on a particular mailbox and outputs data to the screen. 

### Data sensitivity
None

### Local environment requirements
None

### License and/or permissions required

A1 or higher

### How to run the script
```` powershell
.\O365_Get_Mailbox_Forwarding
````

### Known issues
None

### Acknowledgements and Credits
Josh Hartley/University of Missouri
