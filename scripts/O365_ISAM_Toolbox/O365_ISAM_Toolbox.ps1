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
# 1. Run script and enter your O365 admin creds
#
# Written By: Josh Hartley
# University of Missouri
#


# load windows form support
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
 

##### Menu Form Functons #####
function mainMenu() {
    
    # Set the size of your form
    $Form = New-Object System.Windows.Forms.Form
    $Form.width = 450
    $Form.height = 480
    $Form.Text = ”O365 ISAM Toolbox"
    $form.StartPosition = 'CenterScreen'
 
    # Set the font of the text to be used within the form
    $Font = New-Object System.Drawing.Font("Times New Roman",12)
    $Form.Font = $Font

    # Create a group that will contain your radio buttons
    $radioGroupBox = New-Object System.Windows.Forms.GroupBox
    $radioGroupBox.Location = '40,20'
    $radioGroupBox.size = '350,330'
    $radioGroupBox.text = "Toolbox Functions"
    
    # Create the collection of radio buttonns
    $RadioButton1 = New-Object System.Windows.Forms.RadioButton
    $RadioButton1.Location = '20,40'
    $RadioButton1.size = '300,20'
    $RadioButton1.Checked = $false
    $RadioButton1.Text = "Get Message Trace"

    $RadioButton2 = New-Object System.Windows.Forms.RadioButton
    $RadioButton2.Location = '20,70'
    $RadioButton2.size = '300,20'
    $RadioButton2.Checked = $false 
    $RadioButton2.Text = "Get Last Password Reset"
 
    $RadioButton3 = New-Object System.Windows.Forms.RadioButton
    $RadioButton3.Location = '20,100'
    $RadioButton3.size = '300,20'
    $RadioButton3.Checked = $false
    $RadioButton3.Text = "Get Inbox Rules"

    $RadioButton4 = New-Object System.Windows.Forms.RadioButton
    $RadioButton4.Location = '20,130'
    $RadioButton4.size = '300,20'
    $RadioButton4.Checked = $false
    $RadioButton4.Text = "Get Mailbox Forwarding"
    
    $RadioButton5 = New-Object System.Windows.Forms.RadioButton
    $RadioButton5.Location = '20,160'
    $RadioButton5.size = '300,20'
    $RadioButton5.Checked = $false
    $RadioButton5.Text = "Clear Mailbox Forwading"

    $RadioButton6 = New-Object System.Windows.Forms.RadioButton
    $RadioButton6.Location = '20,190'
    $RadioButton6.size = '300,20'
    $RadioButton6.Checked = $false
    $RadioButton6.Text = "Reset Account Password"

    $RadioButton7 = New-Object System.Windows.Forms.RadioButton
    $RadioButton7.Location = '20,220'
    $RadioButton7.size = '300,20'
    $RadioButton7.Checked = $false
    $RadioButton7.Text = "Get Audit Logs (Beta)"
 
    # Add an OK button
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.Location = '175,380'
    $OKButton.Size = '100,40' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult=[System.Windows.Forms.DialogResult]::OK
 
    #Add a cancel button
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Location = '285,380'
    $CancelButton.Size = '100,40'
    $CancelButton.Text = "Quit Toolbox"
    $CancelButton.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
 
    # Add all the GroupBox controls on one line
    $radioGroupBox.Controls.AddRange(@($Radiobutton1,$RadioButton2,$Radiobutton3,$Radiobutton4,$Radiobutton5,$Radiobutton6,$Radiobutton7))

    # Add all the Form controls on one line 
    $form.Controls.AddRange(@($radioGroupBox,$OKButton,$CancelButton))
    
    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $form.AcceptButton = $OKButton
    $form.CancelButton = $CancelButton
 
    # Activate the form
    $form.Add_Shown({$form.Activate()})    
    
    # Get the results from the button click
    $dialogResult = $form.ShowDialog()

    $form.Topmost = $true

    # If the OK button is selected return sender and date info, else cancel button was selected so quit search
    if ($dialogResult -eq "OK"){

        # Check the current state of each radio button and respond accordingly
        if ($RadioButton1.Checked){
            "getmessagetrace"
        }elseif ($RadioButton2.Checked){
            "getlastpw"
        }elseif ($RadioButton3.Checked){
            "getinboxrules"
        }elseif ($RadioButton4.Checked){
            "getforwarding"
        }elseif ($RadioButton5.Checked){
            "setforwarding"
        }elseif ($RadioButton6.Checked){
            "setpw"
        }elseif ($RadioButton7.Checked){
            "auditlogs"
        }else {
            "quit"
        }

    } else {
        "quit"
    }
}

function traceForm($formName) {
    
    # Set the size of your form
    $Form = New-Object System.Windows.Forms.Form
    $Form.width = 600
    $Form.height = 550
    $Form.Text = ”$formName"
    $form.StartPosition = 'CenterScreen'
 
    # Set the font of the text to be used within the form
    $Font = New-Object System.Drawing.Font("Times New Roman",12)
    $Form.Font = $Font

    # Create address textbox
    $inputBox = New-Object System.Windows.Forms.TextBox
    $inputBox.Location = '40,25'
    $inputBox.size = '400,400'
    $inputBox.text = "Enter E-mail Address Here"
    $form.Controls.Add($inputBox)

    # Create a group that will contain your radio buttons
    $radioGroupBox = New-Object System.Windows.Forms.GroupBox
    $radioGroupBox.Location = '40,70'
    $radioGroupBox.size = '350,100'
    $radioGroupBox.text = "Search Type"
    
    # Create the collection of radio buttons
    $RadioButton1 = New-Object System.Windows.Forms.RadioButton
    $RadioButton1.Location = '20,40'
    $RadioButton1.size = '300,20'
    $RadioButton1.Checked = $true 
    $RadioButton1.Text = "Sender Search"
 
    $RadioButton2 = New-Object System.Windows.Forms.RadioButton
    $RadioButton2.Location = '20,70'
    $RadioButton2.size = '300,20'
    $RadioButton2.Checked = $false
    $RadioButton2.Text = "Recipient Search"
 
    # Create a group that will contain your start calendar
    $startGroupBox = New-Object System.Windows.Forms.GroupBox
    $startGroupBox.Location = '40,180'
    $startGroupBox.size = '250,250'
    $startGroupBox.text = "Start Date"

    # Create a group that will contain your end calendar
    $endGroupBox = New-Object System.Windows.Forms.GroupBox
    $endGroupBox.Location = '300,180'
    $endGroupBox.size = '250,250'
    $endGroupBox.text = "End Date"
    
    # Create start calendar
    $startCalendar = New-Object System.Windows.Forms.MonthCalendar
    $startCalendar.Location = '10,25'
    $startCalendar.ShowTodayCircle = $false
    $startCalendar.MaxSelectionCount = 1
    $form.Controls.Add($startCalendar)

    # Create end calendar
    $endCalendar = New-Object System.Windows.Forms.MonthCalendar
    $endCalendar.Location = '10,25'
    $endCalendar.ShowTodayCircle = $false
    $endCalendar.MaxSelectionCount = 1
    $form.Controls.Add($endCalendar)
 
    # Add an OK button
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.Location = '300,450'
    $OKButton.Size = '100,40' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult=[System.Windows.Forms.DialogResult]::OK
 
    #Add a cancel button
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Location = '425,450'
    $CancelButton.Size = '100,40'
    $CancelButton.Text = "Quit Function"
    $CancelButton.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
 
    # Add all the GroupBox controls on one line
    $startGroupBox.Controls.AddRange(@($startCalendar))
    $endGroupBox.Controls.AddRange(@($endCalendar,$endDateBox))
    $radioGroupBox.Controls.AddRange(@($Radiobutton1,$RadioButton2))

    # Add all the Form controls on one line 
    $form.Controls.AddRange(@($inputBox,$radioGroupBox,$startGroupBox,$endGroupBox,$OKButton,$CancelButton))
    
    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $form.AcceptButton = $OKButton
    $form.CancelButton = $CancelButton
 
    # Activate the form
    $form.Add_Shown({$form.Activate()})    
    

    # Get the results from the button click
    $dialogResult = $form.ShowDialog()

    $form.Topmost = $true

    # If the OK button is selected return sender and date info, else cancel button was selected so quit search
    if ($dialogResult -eq "OK"){
        
        $inputBox.Text
        $startCalendar.SelectionStart
        $endCalendar.SelectionStart

        # Check the current state of each radio button and respond accordingly
        if ($RadioButton1.Checked){
            "sender"
        }
        elseif ($RadioButton2.Checked){
            # Call the function
            "recipient"
        }

    } else {
        "quit"
    }
}

function auditForm($formName) {
    
    # Set the size of your form
    $Form = New-Object System.Windows.Forms.Form
    $Form.width = 600
    $Form.height = 420
    $Form.Text = ”$formName"
    $form.StartPosition = 'CenterScreen'
 
    # Set the font of the text to be used within the form
    $Font = New-Object System.Drawing.Font("Times New Roman",12)
    $Form.Font = $Font

    # Create address textbox
    $inputBox = New-Object System.Windows.Forms.TextBox
    $inputBox.Location = '40,25'
    $inputBox.size = '400,400'
    $inputBox.text = "Enter the O365 Primary Address Here"
    $form.Controls.Add($inputBox)
 
    # Create a group that will contain your start calendar
    $startGroupBox = New-Object System.Windows.Forms.GroupBox
    $startGroupBox.Location = '40,60'
    $startGroupBox.size = '250,250'
    $startGroupBox.text = "Start Date"

    # Create a group that will contain your end calendar
    $endGroupBox = New-Object System.Windows.Forms.GroupBox
    $endGroupBox.Location = '300,60'
    $endGroupBox.size = '250,250'
    $endGroupBox.text = "End Date"
    
    # Create start calendar
    $startCalendar = New-Object System.Windows.Forms.MonthCalendar
    $startCalendar.Location = '10,25'
    $startCalendar.ShowTodayCircle = $false
    $startCalendar.MaxSelectionCount = 1
    $form.Controls.Add($startCalendar)

    # Create end calendar
    $endCalendar = New-Object System.Windows.Forms.MonthCalendar
    $endCalendar.Location = '10,25'
    $endCalendar.ShowTodayCircle = $false
    $endCalendar.MaxSelectionCount = 1
    $form.Controls.Add($endCalendar)
 
    # Add an OK button
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.Location = '300,325'
    $OKButton.Size = '100,40' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult=[System.Windows.Forms.DialogResult]::OK
 
    #Add a cancel button
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Location = '425,325'
    $CancelButton.Size = '100,40'
    $CancelButton.Text = "Quit Function"
    $CancelButton.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
 
    # Add all the GroupBox controls on one line
    $startGroupBox.Controls.AddRange(@($startCalendar))
    $endGroupBox.Controls.AddRange(@($endCalendar))

    # Add all the Form controls on one line 
    $form.Controls.AddRange(@($inputBox,$startGroupBox,$endGroupBox,$OKButton,$CancelButton))
    
    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $form.AcceptButton = $OKButton
    $form.CancelButton = $CancelButton
 
    # Activate the form
    $form.Add_Shown({$form.Activate()})    
    
    # Get the results from the button click
    $dialogResult = $form.ShowDialog()

    $form.Topmost = $true

    # If the OK button is selected return sender and date info, else cancel button was selected so quit search
    if ($dialogResult -eq "OK"){
        
        $inputBox.Text
        $startCalendar.SelectionStart
        $endCalendar.SelectionStart

    } else {
        "quit"
    }
}

function functionInput($formName) {
    
    # Set the size of your form
    $Form = New-Object System.Windows.Forms.Form
    $Form.width = 600
    $Form.height = 200
    $Form.Text = $formName
    $form.StartPosition = 'CenterScreen'
 
    # Set the font of the text to be used within the form
    $Font = New-Object System.Drawing.Font("Times New Roman",12)
    $Form.Font = $Font

    # Create input box for account address
    $inputBox = New-Object System.Windows.Forms.TextBox
    $inputBox.Location = '45,25'
    $inputBox.size = '400,400'
    $inputBox.text = "Enter Primary SMTP Address Here"
    $form.Controls.Add($inputBox)
 
    # Add an OK button
    $OKButton = new-object System.Windows.Forms.Button
    $OKButton.Location = '300,100'
    $OKButton.Size = '100,40' 
    $OKButton.Text = 'OK'
    $OKButton.DialogResult=[System.Windows.Forms.DialogResult]::OK
 
    #Add a cancel button
    $CancelButton = new-object System.Windows.Forms.Button
    $CancelButton.Location = '425,100'
    $CancelButton.Size = '100,40'
    $CancelButton.Text = "Quit Function"
    $CancelButton.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
 
    # Add all the Form controls on one line 
    $form.Controls.AddRange(@($inputBox,$OKButton,$CancelButton))
    
    # Assign the Accept and Cancel options in the form to the corresponding buttons
    $form.AcceptButton = $OKButton
    $form.CancelButton = $CancelButton
 
    # Activate the form
    $form.Add_Shown({$form.Activate()})    
    
    # Get the results from the button click
    $dialogResult = $form.ShowDialog()

    $form.Topmost = $true

    # If the OK button is selected run search, else cancel button was selected so quit search
    if ($dialogResult -eq "OK"){
        
        $inputBox.Text

    } else {
        "quit"
    }          
}


##### Toolbox Functions #####
function getmessagetrace() {
    # run search type form and return type picked
        $searchData = traceForm "O365 Message Trace Search"
        
        # set form function return values
        $searchAddress = $searchData[0]
        $startDate = $searchData[1] 
        $endDate = $searchData[2]
        $searchType = $searchData[3]

        Write-host $searchData[0]
        Write-host $searchData[1] 
        Write-host $searchData[2]
        Write-host $searchData[3]
	    
        # if user did not click the quit button, then continue running program
	    If ($searchData -ne "quit") {
        
            # Get date and store as string
            $CurrentDate = Get-Date -UFormat "%m-%d-%Y_%H-%M-%S"
            # set results path to desktop
            $exportResultsPath = ([Environment]::GetFolderPath("Desktop")+"\O365_Message_Trace_Results_$CurrentDate.csv")
            $traceData = $NULL
            $FoundCount = 0

            # If search type is sender then run sender search code, else run recipient search code
            If ( $searchType -eq "sender" ) {
                  
                For($i = 1; $i -le 1000; $i++)  # Maximum allowed pages is 1000
                {
                    $Messages = Get-MessageTrace -SenderAddress $searchAddress -StartDate $startDate -EndDate $endDate -PageSize 5000 -Page $i | select Received, SenderAddress, RecipientAddress, Subject, Status, FromIP, ToIP, StartDate, EndDate
                    
                    If($Messages.count -gt 0) {

                        foreach ($row in $Messages)
                        {
                            $row.Received = $row.Received.ToLocalTime()   
                        }
                        $Messages | Export-Csv $exportResultsPath -NoTypeInformation -Append

                        $FoundCount += $Messages.Count
                    }
                    Else{

                        Break
                    }
                }  

                Write-Host $FoundCount " message were found"

                if ($FoundCount -gt 0) {
                    # display meassege with report location
                    [System.Windows.Forms.MessageBox]::Show($exportResultsPath,'Your report is located here:')
                }
           
            } else {

                # call message trace commandlet with user entered values
                For($i = 1; $i -le 1000; $i++)  # Maximum allowed pages is 1000
                {
                    $Messages = Get-MessageTrace -RecipientAddress $searchAddress -StartDate $startDate -EndDate $endDate -PageSize 5000 -Page $i | select Received, SenderAddress, RecipientAddress, Subject, Status, FromIP, ToIP, StartDate, EndDate
                    
                    If($Messages.count -gt 0) {

                        foreach ($row in $Messages)
                        {
                            $row.Received = $row.Received.ToLocalTime()   
                        }
                        $Messages | Export-Csv $exportResultsPath -NoTypeInformation -Append

                        $FoundCount += $Messages.Count
                    }
                    Else{

                        Break
                    }
                }  

                Write-Host $FoundCount " message were found"

                if ($FoundCount -gt 0) {
                    # display meassege with report location
                    [System.Windows.Forms.MessageBox]::Show($exportResultsPath,'Your report is located here:')
                }
            }
	    }
}

function getlastpw() {
    $searchAddress = functionInput "Get Account Last Password Reset"
	$resetInfo = $NULL

	if ($searchAddress -ne "quit") {
		
		# Get and print last password reset info:
        $resetInfo = Get-MsolUser -UserPrincipalName $searchAddress | select UserPrincipalName,DisplayName,LastPasswordChangeTimeStamp
        $resetInfo.LastPasswordChangeTimeStamp = $resetInfo.LastPasswordChangeTimeStamp.ToLocalTime()
            
        if ($resetInfo -ne $NULL) {
                
            $resetInfo = $resetInfo | Format-List | Out-String
            Write-Host $resetInfo
            # display account reset info
            [System.Windows.Forms.MessageBox]::Show($resetInfo,'Last Reset')

        } else {
                
            $resetInfo = 'Account not found or no data was found.'
            [System.Windows.Forms.MessageBox]::Show($resetInfo,'Error')

        }
	}
}

function getinboxrules() {
    $searchAddress = functionInput "Get Account Inbox Rules"
	
	if ($searchAddress -ne "quit") {
		
		# Get and print inbox rule info:
        $returnedInfo = Get-InboxRule -Mailbox $searchAddress | Select-Object Identity, Enabled, Name, Description, ForwardTo, DeleteMessage | Format-List
            
        if ($returnedInfo -ne $NULL) {
                
            $returnedInfo = $returnedInfo | Format-List | Out-String
            Write-Host $returnedInfo
            # display account info
            [System.Windows.Forms.MessageBox]::Show($returnedInfo,'Inbox Rules')

        } else {
                
            $returnedInfo = 'Account not found or no data was found.'
            [System.Windows.Forms.MessageBox]::Show($returnedInfo,'Error')

        }
	}
}

function getforwarding() {
    $searchAddress = functionInput "Get Mailbox Forwarding"
	
    if ($searchAddress -ne "quit") {
		
    # Get and print Forwording info:
        $returnedInfo = Get-Mailbox  $searchAddress | Select-Object userprincipalname,forwardingsmtpaddress,DeliverToMailboxAndForward,WhenChanged | Format-List
            
        if ($returnedInfo -ne $NULL) {
                
            $returnedInfo = $returnedInfo | Format-List | Out-String
            Write-Host $returnedInfo
            # display account reset info
            [System.Windows.Forms.MessageBox]::Show($returnedInfo,'Forwarding Info')

        } else {
                
            $returnedInfo = 'Account not found or no data was found.'
            [System.Windows.Forms.MessageBox]::Show($returnedInfo,'Error')

        }
    }
}

function setforwarding() {
    $searchAddress = functionInput "Clear Account Forwarding"
	
	if ($searchAddress -ne "quit") {
		
		# set forwarding to null
		Get-Mailbox $searchAddress -ResultSize 1 | Set-Mailbox -ForwardingSmtpAddress $NULL -DeliverToMailboxAndForward $false

		# Get and print forwarding info:
		$returnedInfo = Get-Mailbox  $searchAddress | Select-Object userprincipalname,forwardingsmtpaddress,DeliverToMailboxAndForward,WhenChanged | Format-List

        if ($returnedInfo -ne $NULL) {
                
            $returnedInfo = $returnedInfo | Format-List | Out-String
            Write-Host $returnedInfo
            # display account info
            [System.Windows.Forms.MessageBox]::Show($returnedInfo,'Cleared Forwarding')

        } else {
                
            $returnedInfo = 'Account Not Found'
            [System.Windows.Forms.MessageBox]::Show($returnedInfo,'Error')

        }
	}
}

function setpw() {
    $searchAddress = functionInput "Reset Account Password"
	$returnedInfo = $Null 
    
	if ($searchAddress -ne "quit") {
		
		# Reset with random password
        Set-MsolUserPassword -UserPrincipalName $searchAddress

		# Get and print password reset:
		$returnedInfo = Get-MsolUser -UserPrincipalName $searchAddress | select UserPrincipalName,DisplayName,LastPasswordChangeTimeStamp
        $returnedInfo.LastPasswordChangeTimeStamp = $returnedInfo.LastPasswordChangeTimeStamp.ToLocalTime()

        if ($returnedInfo -ne $NULL) {
                
            $returnedInfo = $returnedInfo | Format-List | Out-String
            Write-Host $returnedInfo
            # display account info
            [System.Windows.Forms.MessageBox]::Show($returnedInfo,'Password Reset')

        } else {
                
            $returnedInfo = 'Account Not Found'
            [System.Windows.Forms.MessageBox]::Show($returnedInfo,'Error')

        }
	}
}

function creddump() {

    # Import the IMAP dll
    $dll = $PSScriptRoot + '\' + 'imap_client\ImapX.dll'
    [Reflection.Assembly]::LoadFile($dll)

    # Create a client object
    $client = New-Object ImapX.ImapClient
    $client.Host = "outlook.office365.com"
    $client.Port = 993
    $client.UseSsl = $true

    Function Get-FileName($initialDirectory){
        [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.initialDirectory = $initialDirectory
        $OpenFileDialog.filter = "txt (*.txt)| *.txt"
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.filename
    }

    # Get date and store as string
    $currentDate = Get-Date -UFormat "%m-%d-%Y_%H-%M-%S"

    # Display dump file format note
    [System.Windows.Forms.MessageBox]::Show('Make sure your cred dump text file is in the following format with each cred on a newline: User01@mail.missouri.edu:Pa$$W0rd!','Dump File Format')

    # Get dump file path and store creds
    $dumpPath = Get-FileName ([Environment]::GetFolderPath("Desktop"))
    $accounts=@()

    if ($dumpPath -ne ""){

        $accounts = Get-Content $dumpPath
    }

    # Set output data to current desktop folder
    $output = ([Environment]::GetFolderPath("Desktop") +"\O35_Creds_Checked_$currentDate.csv")    #Path to output file here

    if ($accounts -ne $NULL){

        $results=@()
        $count = 1

        # printing csv headers
	        echo "Account`tPassword`tSuccessful Auth`t" >> $output

        foreach ($account in $accounts) {
	
	        # Parse email address and password 
	        $useraccount, $password = $account.Split(":")
	
	        # Print processing count to screen
	        write-host "processing account # "$count
            Write-Host "account: " $useraccount
	
	        # pass creds and test IMAP connection and returns value to result variable
            $connection = $client.Connect()
	        $result = $client.Login($useraccount,$password)
	
	        # Print result to screen 
	        Write-Host $result
	        Write-Host ""
	
	        # output results in tab delimited format
	        echo "$useraccount`t$password`t$result`t" >> $output 
	
	        $count++
            $client.Disconnect()
        }

        # display meassege with report location
        [System.Windows.Forms.MessageBox]::Show($output,'Your results are located here:')

    } else{
    
        [System.Windows.Forms.MessageBox]::Show('No file selected or contains no data','Error')

    }

}

function auditlogs() {
    $searchData = auditForm "O365 Get Audit Logs"
        
    # set form function return values
    $searchAddress = $searchData[0]
    $startDate = $searchData[1] 
    $endDate = $searchData[2]
	    
    # if user did not click the quit button, then continue running program
	If ($searchData -ne "quit") {
        
        # Get date and store as string
        $CurrentDate = Get-Date -UFormat "%m-%d-%Y_%H-%M-%S"
        # set results path to desktop
        $exportResultsPath = ([Environment]::GetFolderPath("Desktop")+"\O365_Audit_log_Results_$CurrentDate.csv")
        $sessionID = [DateTime]::Now.ToString().Replace('/', '_')
        $auditData = $NULL
            
        $auditData = Search-UnifiedAuditLog -UserIDs $searchAddress -StartDate $startDate -EndDate $endDate -ResultSize 5000

        if ($auditData -ne $NULL) {
            
            $auditData | Select-Object * | Export-Csv $exportResultsPath -NoTypeInformation
            
            [System.Windows.Forms.MessageBox]::Show('Note: Audit log timestamps are in UTC and are limited to 5,000 logs per request!')
            # display meassege with report location
            [System.Windows.Forms.MessageBox]::Show($exportResultsPath,'Your Audit report is located here:')

        } else {
                
            $returnedInfo = 'No Data Found! Check user ID or expand search dates.'
            [System.Windows.Forms.MessageBox]::Show($returnedInfo,'Error')

        }
	}
}


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
    
    # loop program until user clicks the quit button
    Do { 
        
        $selectedFunctions = mainMenu

        # Check the current state of each radio button and respond accordingly
        if ($selectedFunctions -eq "getmessagetrace"){
            
            Write-Host $selectedFunctions
            # call function
            getmessagetrace

        }elseif ($selectedFunctions -eq "getlastpw"){
            
            Write-Host $selectedFunctions
            # call function
            getlastpw

        
        }elseif ($selectedFunctions -eq "getinboxrules"){
            
            Write-Host $selectedFunctions
            # call function
            getinboxrules
            

        }elseif ($selectedFunctions -eq "getforwarding"){
            
            Write-Host $selectedFunctions
            # call function
            getforwarding

        }elseif ($selectedFunctions -eq "setforwarding"){
            
            Write-Host $selectedFunctions
            # call function
            setforwarding

        }elseif ($selectedFunctions -eq "setpw"){
            
            Write-Host $selectedFunctions
            # call function
            setpw

        }elseif ($selectedFunctions -eq "creddump"){
            
            Write-Host $selectedFunctions
            creddump

        }elseif ($selectedFunctions -eq "auditlogs"){
        
            Write-Host $selectedFunctions
            # call function
	        auditlogs
        }


    } Until ($selectedFunctions -eq "quit")

    #Close session
    Remove-PSSession $Session -ErrorAction $WarningPreference
    Remove-Module MSOnline
}
