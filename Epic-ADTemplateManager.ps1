Import-Module ActiveDirectory
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

##### IDEA: Use object to store variabel values like $yPos. Ones taht change and need a default value #####
# Sets variables
$userFilePath = '..\SupplementaryFiles\ADUserFiles\Users.csv'
$userTemplatePath = '' ##### Not in use yet #####
$defaultWebsitesFilePath = '..\SupplementaryFiles\ADUserFiles\Sites.json'

# Loads data ###### Turn this into a function that reloads all data that needs to be imported
$fileUsers = Import-csv $userFilePath
$adUsers = Get-ADUser -Filter * -Properties *
$adGroups = Get-ADGroup -Filter * -Properties *
#$defaultWebsites = Import-csv $defaultWebsitesFilePath
$jsonContent = Get-Content -Raw -Path $defaultWebsitesFilePath
$defaultWebsites = $jsonContent | ConvertFrom-Json


# Defines functions that create a Windows form object
function New-Label{
    param (
        [string]$Text,
        [object]$Location,
        [int]$xPos,
        [int]$yPos,
        [int]$xSize,
        [int]$ySize
    )
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Size = New-Object System.Drawing.Size($xSize, $ySize)
    $label.Location = New-Object System.Drawing.Point($xPos, $yPos)
    $Location.Controls.Add($label)

    return $label #Return $label to it can be used outside of this fuction ##### Check to make sure this statement is true
}

function New-TextBox{
    param (
        [object]$addTo,
        [string]$Text,
        [int]$xPos,
        [int]$yPos,
        [int]$xSize,
        [int]$ySize
    )
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Text = $Text
    $textBox.Size = New-Object System.Drawing.Size($xSize, $ySize)
    $textBox.Location = New-Object System.Drawing.Point($xPos, $yPos)
    $addTo.Controls.Add($textBox)

    return $textBox #Returns $textbox to it can be used outside of this fuction
}

function New-Button{
    param (
        [object]$addTo,
        [string]$Text,
        [int]$xPos,
        [int]$yPos,
        [int]$xSize,
        [int]$ySize,
        [int]$yAdd ##### Check to see if this works, I don't think it will
    )
    $Button = New-Object Windows.Forms.Button
    $Button.Text = $Text
    $Button.Size = New-Object System.Drawing.Size($xSize, $ySize)
    $Button.Location = New-Object System.Drawing.Point($xPos, $yPos)
    $addTo.Controls.Add($Button)
    $yPos += $yAdd

    return $Button
}

function New-TabPage{
    param (
        [object]$addTo,
        [string]$Text
    )
    $tabPage = New-Object System.Windows.Forms.TabPage
    $tabPage.Text = $Text
    $addTo.TabPages.Add($tabPage)

    return $tabPage ##### Do I need to return this if I add it to a page?
}

function New-Panel{
    param (
        [object]$addTo,
        [int]$xPos,
        [int]$yPos,
        [int]$xSize,
        [int]$ySize
    )
        $panel= New-Object Windows.Forms.panel
        $panel.AutoScroll = $true
        $panelPos = 510 - $yPos     
        $panel.Size = New-Object Drawing.Size($xSize, $panelPos)
        $panel.Location = New-Object Drawing.Point($xPos, $yPos)
        $addTo.Controls.Add($panel)

        return $panel
}

function New-TabControl{
    param (
        [object]$addTo,
        [int]$xPos,
        [int]$yPos,
        [int]$xSize,
        [int]$ySize
    )
    $tabControl = New-Object System.Windows.Forms.TabControl
    $tabControl.Size = New-Object System.Drawing.Size($xSize, $ySize)
    $tabControl.Location = New-Object System.Drawing.Point($xPos, $yPos)
    $addTo.Controls.Add($tabControl)

    return $tabControl
}

function New-ComboBox{
    param (
        [object]$addTo,
        [object]$addRange,
        [string]$Text,
        [int]$xPos,
        [int]$yPos,
        [int]$xSize,
        [int]$ySize
    )
    $comboBox = New-Object System.Windows.Forms.ComboBox
    $comboBox.Size = New-Object System.Drawing.Size($xSize, $ySize)
    $comboBox.Location = New-Object System.Drawing.Point($xPos, $yPos)
    $comboBox.Text = $Text
    foreach($item in $defaultWebsites){
        if ($item.ShowInDropDown){ # shows only items with this value
            $combobox.Items.Add($item.Sites)
        }
    }
    $addTo.Controls.Add($comboBox)

    return $comboBox
}

function New-CheckBox{
    param (
        [object]$addTo,
        [string]$Text,
        [int]$xPos,
        [int]$yPos,
        [int]$xSize,
        [int]$ySize
    )
    $checkBox = New-Object Windows.Forms.CheckBox
    $checkBox.Text = $Text
    $checkBox.Size = New-Object Drawing.Size($xSize, $ySize)
    $checkBox.Location = New-Object Drawing.Point($xPos, $yPos)
    $addTo.Controls.Add($checkBox)

    return $checkBox
}

function New-RadioButton{
    param (
        [object]$addTo,
        [string]$Text,
        [string]$Name,
        [int]$xPos,
        [int]$yPos,
        [int]$xSize,
        [int]$ySize
    )
    $radioButton = New-Object Windows.Forms.RadioButton
    $radioButton.Text = $Text
    $radioButton.Name = $Name
    $radioButton.Size = New-Object Drawing.Size($xSize, $ySize)
    $radioButton.Location = New-Object Drawing.Point($xPos, $yPos)
    $addTo.Controls.Add($radioButton)

    return $radioButton
}


# These arrays hold data that is used to create/edit Windows form objects
$userDataTypes = @( #Holds text for tabPage_B1
    'GivenName',
    'Initials',
    'SurName',
    'Name',
    'DisplayName',
    'SamAccountName',
    'Description',
    'Title',
    'Office',
    'Department',
    'Company'
)

$moreDataTypes = @( #Holds text for tabPage_B2
    "CannotChangePassword",
    "PasswordNeverExpires",
    "Enabled"
) 


# Creates the main Windows form window
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Epic Active Directory Template Manager'
$form.Size = New-Object System.Drawing.Size(1024,768)
$font = New-Object System.Drawing.Font("Arial", 15, [System.Drawing.FontStyle]::Bold)
$form.Font = $font


# Creates a tab control to hold tabs
$tabControl_A = New-TabControl -xSize 500 -ySize 550 -xPos 3 -yPos 10 -addTo $form

    $tabPage_A1 = New-TabPage -Text 'File Users' -addTo $tabControl_A

        $yPos = 20

        # Search box for searching users
        $searchBox_A1 = New-TextBox -xSize 220 -ySize 30 -xPos 20 -yPos $yPos -addTo $tabPage_A1

        # Makes button to be used wtih $searchBox_A1
        $searchButton_A1 = New-Button -Text 'Search' -xSize 220 -ySize 30 -xPos 250 -yPos $yPos -addTo $tabPage_A1
        $yPos += 35

        # Makes Button that will be used to add user info the the User Info section in tabPage_B1
        $button_A1 = New-Button -Text 'Add User to User Info' -xSize 450 -ySize 30 -xPos 20 -yPos $yPos -addTo $tabPage_A1
        $yPos += 35

        # Create a scrollable panel_A1 to hold the radio buttons
        $panel_A1 = New-Panel -xSize 490 -xPos 0 -yPos $yPos -addTo $tabPage_A1

        # Sets the Name for each user. This is not imported with the csv so it is generated here
        foreach ($user in $fileUsers) {
            $nameTemp = $user.GivenName + ' ' + $user.Initials + '. ' + $user.SurName
            $user | Add-Member -MemberType NoteProperty -Name Name -Value ''
            $user.Name = $nameTemp
        }

        # Sets the DisplayName for each user. This is not imported with the csv so it is generated here
        foreach ($user in $fileUsers) {
            $displayNameTemp = $user.SurName + ', ' + $user.GivenName + ' ' + $user.Initials + '.'
            $user | Add-Member -MemberType NoteProperty -Name DisplayName -Value ''
            $user.DisplayName = $displayNameTemp
        }

        # Event handler for Enter key press in the search box
        # Allows the user to key Enter inseted of clicking search
        $searchBox_A1.Add_KeyDown({
            param($sender, $e)
            if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                # Call the search button click event when Enter key is pressed
                $searchButton_A1.PerformClick()
            }
        })

        # Clicks the search button when the form loads so it automatically shows the users
        $form.Add_Shown({$searchButton_A1.PerformClick()})

        $searchButton_A1.Add_Click({

            # Gets text from the search box
            $searchTerm = $searchBox_A1.Text 

            try {
                # Clear existing radio buttons
                $panel_A1.Controls.Clear()

                # Filter AD users based on the search term
                $filteredUsers = $fileUsers | Where-Object { $_.Name -like "*$searchTerm*" -or $_.DisplayName -like "*$searchTerm*" }

                # Add radio buttons for each user to panel_A1
                $yPos = 5
                foreach ($user in $filteredUsers) {
                    $tempText = "$($user.Name) ($($user.SamAccountName))"
                    $radioButton = New-RadioButton -Text $tempText -Name $user.SamAccountName -xPos 20 -yPos $yPos -xSize 450 -ySize 25 -addTo $panel_A1
                    $yPos += 35
                }
            } catch {
                # Handle errors, if any
                [System.Windows.Forms.MessageBox]::Show("Error: $_")
            }
        })
########################### Left off here ############################

        $button_A1.Add_Click({
            # Finds the selected radio button in panel_A1
            $selectedRadioButton = $panel_A1.Controls | Where-Object { $_ -is [System.Windows.Forms.RadioButton] -and $_.Checked }

            if ($selectedRadioButton -ne $null) {
                # Grabs user object from selected Name
                $selectedUser = $fileUsers | Where-Object { $_.SamAccountName -eq $selectedRadioButton.Name }
                $userDataTypes2 = $tabPage_B1.controls | where-object { $_ -is [system.windows.forms.textbox] }

                $buttonCounter_A1 = 0
                foreach ($data in $userDataTypes2){
                    if ($buttonCounter_A1 -lt 6) {
                        $name = $userDataTypes[$buttonCounter_A1]
                        $buttonCounter_A1++
                        $data.Text = $selectedUser.$name
                    }
                }

            } else {
                [System.Windows.Forms.MessageBox]::Show("Please select a user.")
            }
        })


    # Tab for new user info
    $tabPage_A2 = New-TabPage -Text 'AD Users' -addTo $tabControl_A
    # Makes Search Box
        $yPos = 20 

        # Search box for searching users
        $searchBox_A2 = New-TextBox -xSize 220 -ySize 30 -xPos 20 -yPos $yPos -addTo $tabPage_A2

        # Makes button to be used wtih $searchBox_A1
        $searchButton_A2 = New-Button -Text 'Search' -xSize 220 -ySize 30 -xPos 250 -yPos $yPos -addTo $tabPage_A2; $yPos += 35

        # Makes Button that adds user info the the User Info section of pane B
        $button_A2 = New-Button -Text 'Add Account Data to User Info' -xSize 450 -ySize 30 -xPos 20 -yPos $yPos -addTo $tabPage_A2; $yPos += 35

        # Create a scrollable panel_A1 to hold the checkboxes
        $panel_A2 = New-Panel -xSize 490 -xPos 0 -yPos $yPos -addTo $tabPage_A2

            # Event handler for Enter key press in the search box
            $searchBox_A2.Add_KeyDown({
                param($sender, $e)
                if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                    # Call the search button click event when Enter key is pressed
                    $searchButton_A2.PerformClick()
                }
            })

            # Add event handler for the search button
            $searchButton_A2.Add_Click({
                # Get the search term from the search box
                $searchTerm = $searchBox_A2.Text
                try {
                    # Clear existing radio buttons
                    $panel_A2.Controls.Clear()

                    # Filter AD users based on the search term
                    $filteredUsers = $adUsers | Where-Object { $_.Name -like "*$searchTerm*" -or $_.SamAccountName -like "*$searchTerm*" }

                    # Add radio buttons for each user to the panel
                    $yPos = 5
                    foreach ($user in $filteredUsers) {
                        $Text = "$($user.GivenName) $($user.Surname) ($($user.SamAccountName))"
                        $radioButton = New-RadioButton -text $Text -Name $user.SamAccountName -xSize 450 -ySize 25 -xPos 20 -yPos $yPos -addTo $panel_A2
                        $yPos += 35
                    }
                } catch {
                    # Handle errors, if any
                    [System.Windows.Forms.MessageBox]::Show("Error: $_")
                }
            })

            $button_A2.Add_Click({

                $groupDataTypes = @() # Clears the array to there are no orphan objects ##### Not needed? ####
                $panel_b3.controls.clear() # fyi .controls.clear() cannot work on tab pages but works on panels

                # find the selected radio button into panel_a1
                $selectedradiobutton = $panel_a2.controls | where-object { $_ -is [system.windows.forms.radiobutton] -and $_.checked }

                if ($selectedradiobutton -ne $null) { 
                    # Grabs user object from selected Name
                    $selecteduser = $adusers | where-object { $_.samaccountname -eq $selectedradiobutton.name }

                    $userDataTypes2 = $tabPage_B1.controls | where-object { $_ -is [system.windows.forms.textbox] }
                    $buttoncounter_A2 = 0 # counter is used to only loop though the textboxes we want to change (after the first 6)
                    foreach ($data in $userdatatypes2){
                        if ($buttoncounter_a2 -ge 6) {
                            $name = $userDataTypes[$buttoncounter_A2]
                            $data.text = $selecteduser.$name
                            $buttoncounter_A2++
                        } else {
                            $buttoncounter_A2++
                        }
                    }

                    # gets all the text box objects from tabPage_B2. You would not need the where-object because there are only textboxes on that panel, but it future proofs the code.
                    $moreDataTypes2 = $tabPage_B2.controls | where-object { $_ -is [system.windows.forms.checkbox] }
                    # Checks the box if the selected user has that option, Unselects it if it does not
                    foreach ($data in $moreDataTypes2){
                        $name = $data.Text
                        if ($selecteduser.$name) {
                            $data.checked = $true
                        } else {
                            $data.checked = $false
                        }
                    }

                    # Lists all the groups that user has to tabPage_B3
                    $yPosTabPage_B3 = 20
                    foreach ($group in $selecteduser.memberof){
                        # Finds group object form group name
                        $groupname = $adgroups | where-object { $_.distinguishedname -eq $group }

                        # Makes a check box and adds it to an array to be used later
                        $tempbox = new-checkbox -text $groupname.cn -xsize 460 -ysize 30 -xpos 20 -ypos $yPosTabPage_B3 -addto $panel_B3
                        $tempBox.Checked = $true
                        $tempBox.Name = $group
                        $groupDataTypes = $groupDataTypes + $tempBox 

                        $yPosTabPage_B3 += 35
                    }

                    foreach ($url in $defaultWebsites.TextBoxObject) { # Clears the text boxes so new data can be added.
                        $url.Text = ''
                    }

                    $tempCounter = 0 
                    foreach ($url in $selectedUser.url) { # Adds user data to the url list
                        $defaultWebsites[$tempCounter].TextBoxObject.Text = $url
                        $tempCounter++
                    }

                #[System.Windows.Forms.MessageBox]::Show($textBox_B1.Text)

                } else {
                    [System.Windows.Forms.MessageBox]::Show("Please select a user.")
                }
            })


    # Tab for Groups 
    $tabPage_A3= New-TabPage -Text 'AD Groups' -addTo $tabControl_A
        $yPos = 20 

        # Search box for searching users
        $searchBox_A3= New-TextBox -xSize 220 -ySize 30 -xPos 20 -yPos $yPos -addTo $tabPage_A3

        # Makes button to be used wtih $searchBox_A1
        $searchButton_A3 = New-Button -Text 'Search' -xSize 220 -ySize 30 -xPos 250 -yPos $yPos -addTo $tabPage_A3; $yPos += 35

        # Makes button that adds selected group to the group tabControl_B3
        # Use the same buttion the "add selected" to tabPage_B1 (Selected User) or tabPage_B2 (Selected Group)?
        $addSelectedButton_A3 = New-Button -Text 'Add Selected' -xSize 220 -ySize 30 -xPos 20 -yPos $yPos -addTo $tabPage_A3

        # Makes Button that lists users from seleced group
        $getUsersButton_A3 = New-Button -Text 'Get Users' -xSize 220 -ySize 30 -xPos 250 -yPos $yPos -addTo $tabPage_A3; $yPos += 35

        # Create a scrollable panel_A1 to hold the checkboxes
        $panel_A3 = New-Panel -xSize 490 -xPos 0 -yPos $yPos -addTo $tabPage_A3


        # Event handler for Enter key press in the search box
        $searchBox_A3.Add_KeyDown({
            param($sender, $e)
            if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
                # Call the search button click event when Enter key is pressed
                $searchButton_A3.PerformClick()
            }
        })
    
        # Add event handler for the search button
        $searchButton_A3.Add_Click({
            # Get the search term from the search box
            $searchTerm = $searchBox_A3.Text
            try {
                # Clear existing radio buttons
                $panel_A3.Controls.Clear()

                # Filter AD users based on the search term
                $filteredGroups = $adGroups | Where-Object { $_.distinguishedname -like "*$searchTerm*"}

                # Adds list of groups to the screen/panel_A3
                $selectedRadioButtons_A3 = New-Object System.Collections.ArrayList # makes an array for the radio button objecs to be held, like a mother holds her child.

                $yPos = 5
                foreach ($group in $filteredGroups) {
                    $radioButton = New-RadioButton -Text $group.Name -Name $group.Name -xSize 450 -ySize 35 -xPos 20 -yPos $yPos -addTo $panel_A3
                    $selectedRadioButtons_A3.Add($radioButton)
                    $yPos += 35
                }
            } catch {
                # Handle errors, if any
                [System.Windows.Forms.MessageBox]::Show("Error: $_")
            }
        })

        $addSelectedButton_A3.Add_Click({ #adds selected group to tabPage_B3
        
            $yPosTabPage_B3 = $panel_B3.Controls.Count * 35 + 20 # grabs that count of objects on panel_b3 to determin how far down to make each object

            $selectedRadioButton_A3  = $panel_A3.controls | where-object { $_ -is [system.windows.forms.radiobutton] -and $_.checked }
            foreach ($selectedButton in $selectedRadioButton_A3){
                if ($selectedButton.Checked){
                    $group = Get-ADGroup -Filter { Name -eq $selectedButton.Text}

                    $tempbox = new-checkbox -text $group.Name -xsize 460 -ysize 30 -xpos 20 -ypos $yPosTabPage_B3 -addto $panel_B3
                    $tempBox.Checked = $true
                    $tempBox.Name = $group.Name
                    $groupDataTypes = $groupDataTypes + $tempBox # Adds textBox to an array to be used later

                    # Clear the array to it does not have groups we dont want.
                    $yPosTabPage_B3 += 35
                }
            }
        })

    $tabPage_A4 = New-TabPage -Text 'Templates' -addTo $tabControl_A


# Create tab control B
$tabControl_B = New-TabControl -xSize 500 -ySize 550 -xPos 506 -yPos 10 -addTo $form

    # Create the 1 tab
    $tabPage_B1 = New-TabPage -Text 'User Info' -addTo $tabControl_B
        # Creates the display of user information under the "User Info" tab
        $yPosLabel_B1 = 22.5; $yPosTextBox_B1 = 20

        # Loops through each object in $userDataTypes and creates a Label and TextBox for each to display them to the scteen
        foreach ($data in $userDataTypes) {
            $tempText = $data + ':'
            $LabelVariable = New-Label -Text $tempText -xSize 220 -ySize 30 -xPos 20 -yPos $yPosLabel_B1 -Location $tabPage_B1
            $TextBoxVariable = New-TextBox -xSize 220 -ySize 30 -xPos 250 -yPos $yPosTextBox_B1 -addTo $tabPage_B1
            $yPosLabel_B1 += 35; $yPosTextBox_B1 += 35
        }

    # Create the 1 tab
    $tabPage_B2 = New-TabPage -Text 'More Info' -addTo $tabControl_B
        $yPosTabPage_B2 = 20
        foreach ($item in $moreDataTypes){
            $CheckBoxObject = New-CheckBox -Text $item -xSize 460 -ySize 30 -xPos 20 -yPos $yPosTabPage_B2 -addTo $tabPage_B2
            $yPosTabPage_B2 += 35
        }

    $tabPage_B3= New-TabPage -Text 'Groups' -addTo $tabControl_B
        $panel_B3 = New-Panel -xSize 490 -xPos 0 -yPos 0 -addTo $tabPage_B3 # Used to house Groups

    $tabPage_B4 = New-TabPage -Text 'Web Pages' -addTo $tabControl_B

        $yPosTabPage_B4 = 20
        foreach ($site in $defaultWebsites) { # Change this so the amount of combo boxes are not tied to the csv
            # Make a seperate file with 10 spots, use a second csv to append the corret information?
            $site.ComboBoxObject = New-ComboBox -xSize 220 -ySize 30 -xPos 20 -yPos $yPosTabPage_B4 -addTo $tabPage_B4
            $site.TextBoxObject = New-TextBox -xSize 220 -ySize 30 -xPos 250 -yPos $yPosTabPage_B4 -addTo $tabPage_B4 
            $yPosTabPage_B4 += 35
        }

# Makes button for the reload feature
$reloadAllData_Form = New-Button -Text "Reload All Data" -xSize 500 -ySize 30 -xPos 3 -yPos 565 -AddTo $form


# Makes button for the Add User to Active Directory feature
$addADUserButton_Form = New-Button -Text "Add User to Active Directory" -xSize 500 -ySize 30 -xPos 506 -yPos 565 -AddTo $form

    $addADUserButton_Form.Add_Click({

        # Grabs user info from tabpage_b1
        $newUserTextBoxes = $tabPage_B1.controls | where-object { $_ -is [system.windows.forms.textbox] }

        New-ADUser -GivenName $newUserTextBoxes[0].Text -Initials $newUserTextBoxes[1].Text -SurName $newUserTextBoxes[2].Text -Name $newUserTextBoxes[3].Text -DisplayName $newUserTextBoxes[3].Text -SamAccountName $newUserTextBoxes[4].Text -Description $newUserTextBoxes[5].Text -Title $newUserTextBoxes[6].Text -Office $newUserTextBoxes[7].Text -Department $newUserTextBoxes[8].Text -Company $newUserTextBoxes[9].Text #-Organization -OfficePhone $selectedUser.OfficePhone -CannotChangePassword $cannotChangePassword -PasswordNeverExpires $passwordNeverExpires -Enabled $enabled -AccountPassword "Apple21*"

        foreach ($site in $defaultWebsites) { 
            if ($site.ComboBoxObject.SelectedItem) {
                $tempText = $site.ComboBoxObject.SelectedItem + $site.TextBoxObject.Text
                write-host $tempText
            }
        }
    })


# Makes button and text box for the Save As Template feature
$addTemplateButton_Form = New-Button -Text "Save As Template" -xSize 245 -ySize 30 -xPos 761 -yPos 600 -addTo $form
$addTemplateTextBox_Form = New-TextBox -xSize 245 -ySize 30 -xPos 506 -yPos 600 -addTo $form



# Display the form
[void]$form.ShowDialog()