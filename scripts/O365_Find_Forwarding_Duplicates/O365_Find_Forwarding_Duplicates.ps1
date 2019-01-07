#
# Use the following PowerShell script to pull All O365 mailbox forwarding and find any duplicates.  Once an attacker
# has compromised an O365 account, sometimes they will setup forwarding to receive a copy
# of the victims e-mail.  If you find a "funny" looking forward with this script, to can
# use the O365_Set_Forwarding_To_Null script to clear the forward.
# 
# Dependiences:
#  1) Install the 64-bit version of the Microsoft Online Services Sign-in Assistant:
# 	  	- https://go.microsoft.com/fwlink/p/?LinkId=286152
#  2) Install the MSOnline module
#	  	- Open an administrator-level PowerShell command prompt.
#		- Run the Install-Module MSOnline command
# MS White paper for more detail: https://docs.microsoft.com/en-us/office365/enterprise/powershell/connect-to-office-365-powershell
# 
# Instructions:
# 
# 1. Run script and enter your O365 admin creds (Warning: this script can take a couple of hours (or more) to run depending on the number of mailboxes in your tenant)
# 2. Output: CSV output to desktop
#
#
# Written By: Josh Hartley
# University of Missouri
#



##### Auth code and control logic #####
# Do while to run authenication
Do {
    
    # set connection vars to NULL
    $UserCredential = $NULL
    $Session = $NULL

    # Get Credentials
    $UserCredential = Get-Credential

    # If credentials were entered, try connection
    If ($UserCredential -ne $NULL) {
        # Import MSOnline Module
        Import-Module MSOnline

        # Open session
        $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

        Import-PSSession $Session

        # Create a connection to O365
        $Connection = Connect-MsolService –Credential $UserCredential
    }

    # If the Session variable is NULL credentials failed
    If ($Session -eq $NULL -and $UserCredential -ne $NULL) {
        [System.Windows.Forms.MessageBox]::Show('Bad Username or Password','Error')
    }

 # Do until auth has not failed or creds were null (which means user clicked cancel)
} Until (($Session -ne $NULL) -or ($UserCredential -eq $NULL))




##### Main program control logic #####
# If auth was successful or credentials not NULL, then run main program
If (($Session -ne $NULL) -or ($UserCredential -ne $NULL)) {
  

    # Get date and store as string
    $CurrentDate = Get-Date -UFormat "%m-%d-%Y_%H-%M-%S"


    # Set filename and path"
    $currentForwardsFile = $PSScriptRoot + '\o365_mailbox_forwarding_' + $CurrentDate + '.csv'
    $dupMailboxFile = $PSScriptRoot + '\o365_duplicate_mailbox_forwarding_' + $CurrentDate + '.csv'


    # Get mailbox forwards and export to CSV
    Get-Mailbox -ResultSize unlimited |
    where {(($_.forwardingsmtpaddress -ne $null) -or ($_.forwardingaddress -ne $null))} | 
    Select-Object userprincipalname,forwardingsmtpaddress,DeliverToMailboxAndForward,WhenChanged | 
    Export-Csv $currentForwardsFile -NoTypeInformation


    # Close O365 session
    Remove-PSSession $Session -ErrorAction $WarningPreference
    Remove-Module MSOnline


    # find any duplicatses and export to CSV
    Import-Csv $currentForwardsFile |
    Group-Object -Property ForwardingSmtpAddress |
    Where-Object { $_.count -ge 2 } |
    Foreach-Object { $_.Group } |
    Select UserPrincipalName, ForwardingAddress, ForwardingSmtpAddress, DeliverToMailboxAndForward, WhenChanged |
    Export-csv -Path $dupMailboxFile -NoTypeInformation
}

