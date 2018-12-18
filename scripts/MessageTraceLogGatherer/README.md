## getmessagetracelogs.ps1

### Use case

This script will pull message trace logs and output them to JSON, in order to be ingested/parsed by a SIEM

### Brief description 

When running this script, it iwll connect out to M365, and pull message trace logs for the dates/times it does not already have a log for.

### Data sensitivity

Meta data on email. To/From, Subject, IP Addresses

### Local environment requirements

Powershell and Windows.

### License and/or permissions required

None.

### How to run the script
Update the two JSON config files [Authentication.json](Authentication.json) and [MessageTraceConfiguration.json](MessageTraceConfiguration.json)

```` powershell
.\getmessagetracelogs.ps1
````

### Known issues

Not all message logs are pulled. this is due to M365 not having the log available when we pull it. If you want to ensure more logs are pulled, change [line #36](https://github.com/renisac/O365Logs/blob/b9cd3788a7dc0bc811b4f72164a785825b0bc5fc/scripts/MessageTraceLogGatherer/getmessagetracelogs.ps1#L36) to only collect after 6 hours (-1 to -6).

### Acknowledgements and Credits

[Ryan McVicar](https://github.com/mcvic1rj)/[Central Michigan University](https://github.com/orgs/CentralMichigan/)
