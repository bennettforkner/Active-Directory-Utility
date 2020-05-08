
. D:\Scripts\ActiveDirectory\data_views.ps1
. D:\Scripts\ActiveDirectory\AD_functions.ps1

<# show-form_1:
    Function to display the first form to the user.
    @param $error_message: the message to be displayed to the user on this window. 
    @return an array of the entered values. #>
function show-form_1($error_message,$error_color) {

    <# $script:function: the script-scoped variable to store the selected function of the user. #>
    $script:function = ""
    
    <# Type Imports. #>
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    <# The form panel to be displayed. #>
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'AD Group Search'
    $form.Size = New-Object System.Drawing.Size(300,180)
    $form.StartPosition = 'CenterScreen'

    <# The 'Add Member' button. #>
    $aButton = New-Object System.Windows.Forms.Button
    $aButton.Location = New-Object System.Drawing.Point(20,70)
    $aButton.Size = New-Object System.Drawing.Size(75,23)
    $aButton.Text = 'ADD'
    $aButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    
    <# Functionality to set the $script:function to add when the button is clicked. #>
    $aButton.Add_Click({
        $script:function = "A"
        $form.Close()
    })

    $form.AcceptButton = $aButton
    $form.Controls.Add($aButton)

    <# The 'Remove Member' button. #>
    $rButton = New-Object System.Windows.Forms.Button
    $rButton.Location = New-Object System.Drawing.Point(100,70)
    $rButton.Size = New-Object System.Drawing.Size(75,23)
    $rButton.Text = 'REMOVE'
    $rButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    
    <# Functionality to set the $script:function to remove when the button is clicked. #>
    $rButton.Add_Click({
        $script:function = "R"
        $form.Close()
    })

    $form.AcceptButton = $rButton
    $form.Controls.Add($rButton)

    <# The 'Search Group' button. #>
    $sButton = New-Object System.Windows.Forms.Button
    $sButton.Location = New-Object System.Drawing.Point(180,70)
    $sButton.Size = New-Object System.Drawing.Size(75,23)
    $sButton.Text = 'SEARCH'
    $sButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    
    <# Functionality to set the $script:function to add when the button is clicked. #>
    $sButton.Add_Click({6
        $script:function = "S"
        $form.Close()
    })

    $form.AcceptButton = $sButton
    $form.Controls.Add($sButton)

    <# The prompt. #>
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10,50)
    $label1.Size = New-Object System.Drawing.Size(280,20)
    $label1.Text = 'Please choose the function you would like to perform'
    $form.Controls.Add($label1)
    
    <# Display the passed error message if applicable. #>
    if ($error_message) {

        $log = ("[" + (get-date) + "]: [ERROR] $error_message")
        Add-content $global:file_path $log


        $label4 = New-Object System.Windows.Forms.Label
        $label4.Location = New-Object System.Drawing.Point(10,0)
        $label4.Size = New-Object System.Drawing.Size(280,40)
        $label4.Text = $error_message
        if ($error_color) {
            $label4.ForeColor = $error_color
        } else {
            $label4.ForeColor = 'red'
        }
        $form.Controls.Add($label4)

    }

    <# Display the $form and set to to the frontmost window (Function pauses here until submit). #>  
    $form.Topmost = $true
    $result = $form.ShowDialog()

    <# Function to send control to the proper channels depending on the user-selected function. #>
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($script:function -eq "S") {
            show-form_search
        } if ($script:function -eq "A") {
            show-form_Add
        } if ($script:function -eq "R" ) {
            show-form_remove
        }
    } else {
        Add-content $global:file_path ("[" + (get-date) + "]: End of Log File`n")
        $log = get-content $global:file_path
        invoke-sqlcmd -serverinstance $script:ServerInstance ("INSERT INTO [adops].[dbo].[AD_Group_Update_Log] (Jobname,Username,Timestamp,Notes) VALUES ('Powershell AD Group Utility','$env:username','" + (get-date) + "','$log')")
        exit
    }
}



<# show-form_search:
    Function to display the form for the user to fill out about the group information that they would like to query. 
    @param $error_message: the error message to be displayed if passed from the last function.
    @call: show-search_results with the received data if query processed correctly, else show-form_search (recurse). If user exits, show-form_1.#>
    function show-form_search($error_message) {
    
    <# Import types. #>
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    <# The form panel to be displayed. #>
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Data Entry Form'
    $form.Size = New-Object System.Drawing.Size(300,250)
    $form.StartPosition = 'CenterScreen'

    <# The 'search' button. #>
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(107,175)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'SEARCH'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    <# The department prompt. #>
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10,20)
    $label1.Size = New-Object System.Drawing.Size(280,20)
    $label1.Text = 'What department would you like to search in?:'
    $form.Controls.Add($label1)

    <# The department text box. #>
    $group = New-Object System.Windows.Forms.TextBox
    $group.Location = New-Object System.Drawing.Point(10,40)
    $group.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($group)

    <# The fields prompt. #>
    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(10,65)
    $label2.Size = New-Object System.Drawing.Size(280,20)
    $label2.Text = 'What fields would you like to view?:'
    $form.Controls.Add($label2)

    <# Create a CheckedListBox. #>
    $CheckedListBox = New-Object -TypeName System.Windows.Forms.CheckedListBox;

    <# Add the CheckedListBox to the Form. #>
    $Form.Controls.Add($CheckedListBox);

    <# Widen the CheckedListBox. #>
    $CheckedListBox.Width = 260;
    $CheckedListBox.Height = 70;
    $CheckedListBox.Location = New-Object System.Drawing.Point(10,85)

    <# Add 10 items to the CheckedListBox. #>
    $CheckedListBox.Items.Add("Name") | out-null
    $CheckedListBox.Items.Add("Email") | out-null
    $CheckedListBox.Items.Add("SamAccountName") | out-null
    $CheckedListBox.Items.Add("SID") | out-null

    <# Clear all existing selections. #>
    $CheckedListBox.ClearSelected();

    <# Set the default checked items in the CheckedListBox. #>
    $CheckedListBox.SetItemChecked($CheckedListBox.Items.IndexOf("Name"), $true);

    <# display error message if applicable. #>
    if ($error_message) {

        <# The error message. #>
        $label4 = New-Object System.Windows.Forms.Label
        $label4.Location = New-Object System.Drawing.Point(10,0)
        $label4.Size = New-Object System.Drawing.Size(280,20)
        $label4.Text = $error_message
        $label4.ForeColor = 'red'
        $form.Controls.Add($label4)
    }

    <# Displaye the form and set it to the frontmost window. (Pause until form submitted. #>
    $form.Topmost = $true
    $result = $form.ShowDialog()

    <# On form submit; if exited normally, if the group was filled out, show the search results, else show-form_search, else show start form. #>
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {

        if ($group.text -like 'end-user-list') {
            $log = ("[" + (get-date) + "]: Attempted to view End-User-List")
            Add-content $global:file_path $log
            show-form_1 "This group may not be viewed."
        }

        if ($group.Text) {
            show-search_results (get-members $group.text) $CheckedListBox.CheckedItems "" "" $group.text
        } else {
            show-form_search "Please enter a valid group name."
        }
    } else {
        show-form_1
    }
}

<# show-form_add:
    Function to prompt the user for an aduser and group that they would like to link.
    @param $error_message: the message from the last control point that needs to be displayed to the user.
    @param $passed_group: the group to be placed in the textbox on recusive correction.
    @param $passed_name: the name to be placed in the textbox on recursive correction.
    @call: add-members with entered information, show-form_add on error, show-form_1 on close. #>
function show-form_add($error_message, $passed_group, $passed_name) {
    
    <# Import types. #>
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    <# The form panel to be displayed. #>
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Add Active Directory Group Members'
    $form.Size = New-Object System.Drawing.Size(300,200)
    $form.StartPosition = 'CenterScreen'

    <# The 'submit' button. #>
    $addButton = New-Object System.Windows.Forms.Button
    $addButton.Location = New-Object System.Drawing.Point(100,128)
    $addButton.Size = New-Object System.Drawing.Size(75,23)
    $addButton.Text = 'ADD'
    $addButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $addButton
    $form.Controls.Add($addButton)

    <# The group prompt. #>
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10,20)
    $label1.Size = New-Object System.Drawing.Size(280,20)
    $label1.Text = 'Active Directory Group:'
    $form.Controls.Add($label1)

    <# The group text box. #>
    $group = New-Object System.Windows.Forms.TextBox
    $group.Location = New-Object System.Drawing.Point(10,40)
    $group.Size = New-Object System.Drawing.Size(260,20)
    $group.text = $passed_group
    $form.Controls.Add($group)

    <# The name/s prompt. #>
    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(10,80)
    $label2.Size = New-Object System.Drawing.Size(280,20)
    $label2.Text = 'Comma-separated Gordon Login/s (first.last):'
    $form.Controls.Add($label2)

    <# The name/s text box. #>
    $name = New-Object System.Windows.Forms.TextBox
    $name.Location = New-Object System.Drawing.Point(10,100)
    $name.Size = New-Object System.Drawing.Size(260,80)
    $name.text = $passed_name
    $form.Controls.Add($name)

    <# Display the error message if passed. #>
    if ($error_message) {

        <# The error message. #>
        $label3 = New-Object System.Windows.Forms.Label
        $label3.Location = New-Object System.Drawing.Point(10,0)
        $label3.Size = New-Object System.Drawing.Size(280,20)
        $label3.Text = $error_message
        $label3.ForeColor = 'red'
        $form.Controls.Add($label3)
    }

    <# Show form and set to frontmost window. Pause here until form submit. #>
    $form.Topmost = $true
    $result = $form.ShowDialog()

    <# If proper submit and no errors, call add-members with passed info, else recurse and pass information back. #>
    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        if ($group.text -AND $name.text) {

            if ($group.text -like 'end-user-list') {
                $log = ("[" + (get-date) + "]: Attempted to add to End-User-List")
                Add-content $global:file_path $log
                show-form_1 "This group may not be edited."
            }

            try {
                add-members $group.text $name.text
            } catch {
                show-form_add "Invalid entry. Please confirm entered values." $group.text $name.text
            }
        } else {
            show-form_add "Please make sure all fields are filled." $group.text $name.text
        }
    } else {
        show-form_1
    }
}


<# show-form_remove:
    Function to display the form for removing users from an active directory group to the user.
    @param $error_message: the message from the last control point that needs to be displayed to the user.
    @param $passed_group: the group to be placed in the textbox on recusive correction.
    @param $passed_name: the name to be placed in the textbox on recursive correction.
    @call: remove-members with entered information, show-form_remove on error, show-form_1 on close. #>
function show-form_remove($error_message, $passed_group, $passed_name) {
    
    <# Type Imports. #>
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    <# The form panel to be displayed. #>
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Remove Group Members'
    $form.Size = New-Object System.Drawing.Size(300,200)
    $form.StartPosition = 'CenterScreen'

    <# The 'submit' button. #>
    $addButton = New-Object System.Windows.Forms.Button
    $addButton.Location = New-Object System.Drawing.Point(100,128)
    $addButton.Size = New-Object System.Drawing.Size(75,23)
    $addButton.Text = 'REMOVE'
    $addButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $addButton
    $form.Controls.Add($addButton)

    <# The group prompt. #>
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10,20)
    $label1.Size = New-Object System.Drawing.Size(280,20)
    $label1.Text = 'Active Directory Group:'
    $form.Controls.Add($label1)

    <# The group text box. #>
    $group = New-Object System.Windows.Forms.TextBox
    $group.Location = New-Object System.Drawing.Point(10,40)
    $group.Size = New-Object System.Drawing.Size(260,20)
    $group.text = $passed_group
    $form.Controls.Add($group)

    <# The aduser prompt. #>
    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(10,80)
    $label2.Size = New-Object System.Drawing.Size(280,20)
    $label2.Text = 'Comma-separated Gordon Login/s (first.last):'
    $form.Controls.Add($label2)

    <# The aduser text box. #>
    $name = New-Object System.Windows.Forms.TextBox
    $name.Location = New-Object System.Drawing.Point(10,100)
    $name.Size = New-Object System.Drawing.Size(260,80)
    $name.text = $passed_name
    $form.Controls.Add($name)

    <# Display error message if applicable. #>
    if ($error_message) {

        <# The error message. #>
        $label3 = New-Object System.Windows.Forms.Label
        $label3.Location = New-Object System.Drawing.Point(10,0)
        $label3.Size = New-Object System.Drawing.Size(280,20)
        $label3.Text = $error_message
        $label3.ForeColor = 'red'
        $form.Controls.Add($label3)
    }

    <# Display form and set it to topmost window. Pause until form submit. #>
    $form.Topmost = $true
    $result = $form.ShowDialog()

    <# If proper submit and no errors, call remove-members with passed info, else recurse and pass information back. #>
    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        if ($group.text -AND $name.text) {

            if ($group.text -like 'end-user-list') {
                $log = ("[" + (get-date) + "]: Attempted to remove from End-User-List")
                Add-content $global:file_path $log
                show-form_1 "This group may not be edited."
            }

            try {
                remove-members $group.text $name.text
            } catch {
                show-form_remove "Invalid entry. Please confirm entered values." $group.text $name.text
            }
        } else {
            show-form_remove "Please make sure all fields are filled." $group.text $name.text
        }
    } else {
        show-form_1
    }
}

<# show-form_export:
    Function to display the form for the user to enter the location where they would like to save a .csv export fo search data. 
    @param $data: the objects to be exported to a .csv file.
    @param $error_message: the error message from the last call to be displayed in this form.
    @call (Export-CSV and show-form_1 on proper submit | show-form_export on improper submit | show-form_1 on close). #>
function show-form_export($data, $error_message) {
    
    <# Type imports. #>
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    <# The form panel to be displayed. #>
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Data Entry Form'
    $form.Size = New-Object System.Drawing.Size(300,300)
    $form.StartPosition = 'CenterScreen'

    <# The 'submit' button. #>
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(107,100)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = 'EXPORT'
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $okButton
    $form.Controls.Add($okButton)

    <# The path prompt. #>
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10,20)
    $label1.Size = New-Object System.Drawing.Size(280,20)
    $label1.Text = 'What would you like the file to be saved as?:'
    $form.Controls.Add($label1)

    <# The path text box. #>
    $path = New-Object System.Windows.Forms.TextBox
    $path.Location = New-Object System.Drawing.Point(10,40)
    $path.Size = New-Object System.Drawing.Size(260,20)
    $path.text = "D:\Users\$env:username\Documents\ActiveDirectory\export.csv"
    $form.Controls.Add($path)

    <# Display $error_message if applicable. #>
    if ($error_message) {

        <# The error message. #>
        $label4 = New-Object System.Windows.Forms.Label
        $label4.Location = New-Object System.Drawing.Point(10,0)
        $label4.Size = New-Object System.Drawing.Size(280,20)
        $label4.Text = $error_message
        $label4.ForeColor = 'red'
        $form.Controls.Add($label4)
    }

    <# Display form and set to frontmost window. pause until form close. #>
    $form.Topmost = $true
    $result = $form.ShowDialog()

    <# If path entered, build .csv and return to start form, else recurse and pass information back. #>
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($path.text) {
            Write-Host "Building csv..."
            remove-item $path.text -ErrAction Ignore
            foreach ($d in $data) {
                Add-Member -InputObject $d -TypeName "gordonEmail"
                $d.gordonEmail = ($d.SamAccountName + "@gordon.edu")

                $entry = @{
                        name = $d.name
                        gordonEmail = (get-aduser $d.SamAccountName).userprincipalname
                        SamAccountName = $d.SamAccountName
                    }
                $d = $entry
                (New-Object psobject -Property $entry) | Select-Object "name", "gordonEmail", "SamAccountName" | export-csv -Path $path.text -Append -NoTypeInformation
                Write-Host "Exported line: ["$d.name"]."
            }
            
             $log = ("[" + (get-date) + "]: Exported above group")
             Add-content $global:file_path $log

            show-form_1 ("Data exported successfully at " + $path.text) 'green'

        } else {
            return show-form_export
        }
    } else {
        show-form_1
    }
}


<# Initializer code. #>
if (!(get-adgroupmember End-User-List).samaccountname.contains($env:username)) {
    $popup = new-object -comobject wscript.shell
    $popup.popup("You are not authorized to use this utility.`nIf you believe this is a mistake, please contact CTS.", 0,"Active Directory Utility",0)
    return;
}

if (!(Get-NetConnectionProfile).name.contains("gordon.edu")) {
    $popup = new-object -comobject wscript.shell
    $popup.popup("You need to be connected to the 'gordon.edu' domain (internet or vpn) to use this tool.`nIf you believe there is a problem, please contact CTS.", 0,"Active Directory Utility",0)
    return;
}

$script:ServerInstance = "AdminProdSQL"

<# $index: the number of the log file to be created. #>
$index = 0

<# Create directory in documents to store logs and csv files. #>
mkdir D:\Users\$env:username\Documents\ActiveDirectory -ErrorAction 'ignore'
$global:file_path = "D:\Users\$env:username\Documents\ActiveDirectory\ADlog_0.txt"
while ((Test-Path $global:file_path)) {
    $index += 1
    $global:file_path = "D:\Users\$env:username\Documents\ActiveDirectory\ADlog_" + $index + ".txt"
}

New-Item -Path $global:file_path -ItemType File | out-null
Add-content $global:file_path ("[" + (get-date) + "]: Starting log of Active Directory Utility.")
Add-content $global:file_path ("[" + (get-date) + "]: Host: $env:computername")
Write-Host "Log file stored as '$global:file_path'"

show-form_1