#
# Use the following PowerShell script to pull O365 message tracking info.  If you need to 
# pull messages sent by a specific sender or recipient to determine whether they have sent 
# or received suspicious email, use the script.
#  
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
# 2. Pick the type of search, sender or recipient
# 3. Enter the sender or recipient email address and select a date range 
#
# Output: csv file located on your desktop with filename format: O365_Message_Trace_Results_CurrentDateStamp.csv
#
# Written By: Josh Hartley
# University of Missouri
#


# load windows form support
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 


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
    $inputBox.text = "Enter the O365 Primary Address Here"
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
    $CancelButton.Text = "Quit Script"
    $CancelButton.DialogResult=[System.Windows.Forms.DialogResult]::Cancel
 
    # Add all the GroupBox controls on one line
    $startGroupBox.Controls.AddRange(@($startCalendar))
    $endGroupBox.Controls.AddRange(@($endCalendar))
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


# If auth was successful or credentials not NULL, then run main program
If (($Session -ne $NULL) -or ($UserCredential -ne $NULL)) {
    
    # loop program until user clicks the quit button
    Do { 
        
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

    } Until ($searchData -eq "quit")

    #Close session
    Remove-PSSession $Session -ErrorAction $WarningPreference
    Remove-Module MSOnline
}
