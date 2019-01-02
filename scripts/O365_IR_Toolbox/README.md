## O365_IR_Toolbox.ps1

### Use case

This PowerShell script is an all-in-one O365 IR toolbox with the following functions:
- Get Message Trace Logs
- Get Azure AD Audit Logs
- Get Last Password Reset
- Get Inbox Rules
- Get Mailbox Forwarding
- Clear Mailbox Forwarding
- Reset Account Password

### Brief description

Run script and enter O365 admin creds.
- Get Message Trace Logs > outputs csv file to Desktop
- Get Azure AD Audit Logs > outputs csv file to Desktop
- Get Last Password Reset > outputs LastPasswordChangedTimestamp to screen
- Get Inbox Rules > outputs to screen
- Get Mailbox Forwarding > outputs to screen
- Clear Mailbox Forwarding > sets forwarding to NULL and outputs change to screen
- Reset Account Password > set new pw and outputs change to screen

### Data sensitivity
- E-mail meta data
- Authentication data

### Local environment requirements
None

### License and/or permissions required

A1 or higher

### How to run the script
```` powershell
.\O365_IR_Toolbox.ps1
````

### Known issues
None

### Acknowledgements and Credits
Josh Hartley/University of Missouri
