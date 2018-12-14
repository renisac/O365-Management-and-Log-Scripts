# Template for Script README.md 

Please use the following format for the README.md for script submissions. View this document Raw and copy/paste to grab the markdown formatting.

***

## Script_Name

The script name should be descriptive, e.g. O365_Get_InboxRules.

### Use case

Describe the operational purpose for this script. For example, considering the O365_Get_InboxRules mentioned above, the Use Case may be:

> Use after an account has been compromised to check for inbox rules the attacker may have set.

### Brief description

Succinctly describe the:
- action(s) performed by the script, 
- log collection method, 
- logs touched, and
- character of the output data

For example the O365_Get_InboxRules Brief Description may be: 

> Get all inbox rules set on a specified mailbox. This PowerShell script uses the Exchange Online Actions log and outputs a list of the inbox rules on the specified mailbox.

### Data sensitivity

Make note of any sensitivity considerations for processed data in intermediate or final output forms.

### Local environment requirements

Environment variables, configuration files, &c; or, "none".

### License and/or permissions required

Microsoft license/tenant level, &c.

### How to run the script

Essentially, the --help information if not available directly via the script.

### Known issues

Any known issues.

### Acknowledgements and Credits

If the author wishes attribution and/or other credit(s) should be acknowledged.
