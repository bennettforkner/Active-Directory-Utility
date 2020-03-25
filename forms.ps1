
. D:\Scripts\ActiveDirectory\data_views.ps1
. D:\Scripts\ActiveDirectory\AD_functions.ps1

<# show-form_1:
    Function to display the first form to the user.
    @param count: the index of this loop in the locator
    @return an array of the entered values #>
function show-form_1($error_message) {

    $script:function = ""

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    <# The form panel to be displayed. #>
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Active Directory Group Search'
    $form.Size = New-Object System.Drawing.Size(300,150)
    $form.StartPosition = 'CenterScreen'

    <# The 'Add Member' button. #>
    $aButton = New-Object System.Windows.Forms.Button
    $aButton.Location = New-Object System.Drawing.Point(20,50)
    $aButton.Size = New-Object System.Drawing.Size(75,23)
    $aButton.Text = 'ADD'
    $aButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $aButton.Add_Click({
        $script:function = "A"
        $form.Close()
    })
    $form.AcceptButton = $aButton
    $form.Controls.Add($aButton)

    <# The 'Remove Member' button. #>
    $rButton = New-Object System.Windows.Forms.Button
    $rButton.Location = New-Object System.Drawing.Point(100,50)
    $rButton.Size = New-Object System.Drawing.Size(75,23)
    $rButton.Text = 'REMOVE'
    $rButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $rButton.Add_Click({
        $script:function = "R"
        $form.Close()
    })
    $form.AcceptButton = $rButton
    $form.Controls.Add($rButton)

    <# The 'Search Group' button. #>
    $sButton = New-Object System.Windows.Forms.Button
    $sButton.Location = New-Object System.Drawing.Point(180,50)
    $sButton.Size = New-Object System.Drawing.Size(75,23)
    $sButton.Text = 'SEARCH'
    $sButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $sButton.Add_Click({
        $script:function = "S"
        $form.Close()
    })
    $form.AcceptButton = $sButton
    $form.Controls.Add($sButton)

    <# The building prompt. #>
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10,20)
    $label1.Size = New-Object System.Drawing.Size(280,20)
    $label1.Text = 'Please choose the function you would like to perform'
    $form.Controls.Add($label1)

    if ($error_message) {
        $label4 = New-Object System.Windows.Forms.Label
        $label4.Location = New-Object System.Drawing.Point(10,0)
        $label4.Size = New-Object System.Drawing.Size(280,20)
        $label4.Text = $error_message
        $label4.ForeColor = 'red'
        $form.Controls.Add($label4)
    }

    $form.Topmost = $true
    $result = $form.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($script:function -eq "S") {
            show-form_search
        } if ($script:function -eq "A") {
            show-form_Add
        } if ($script:function -eq "R" ) {
            show-form_remove
        }
    } else {
        return
    }
}




function show-form_search($error_message) {

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

    <# The building prompt. #>
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10,20)
    $label1.Size = New-Object System.Drawing.Size(280,20)
    $label1.Text = 'What department would you like to search in?:'
    $form.Controls.Add($label1)

    <# The building text box. #>
    $group = New-Object System.Windows.Forms.TextBox
    $group.Location = New-Object System.Drawing.Point(10,40)
    $group.Size = New-Object System.Drawing.Size(260,20)
    $form.Controls.Add($group)

    <# The building prompt. #>
    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(10,65)
    $label2.Size = New-Object System.Drawing.Size(280,20)
    $label2.Text = 'What fields would you like to view?:'
    $form.Controls.Add($label2)

    # Create a CheckedListBox
    $CheckedListBox = New-Object -TypeName System.Windows.Forms.CheckedListBox;

    # Add the CheckedListBox to the Form
    $Form.Controls.Add($CheckedListBox);

    # Widen the CheckedListBox
    $CheckedListBox.Width = 260;
    $CheckedListBox.Height = 70;
    $CheckedListBox.Location = New-Object System.Drawing.Point(10,85)

    # Add 10 items to the CheckedListBox
    $CheckedListBox.Items.Add("Name") | out-null
    $CheckedListBox.Items.Add("Email") | out-null
    $CheckedListBox.Items.Add("SamAccountName") | out-null
    $CheckedListBox.Items.Add("SID") | out-null

    # Clear all existing selections
    $CheckedListBox.ClearSelected();

    $CheckedListBox.SetItemChecked($CheckedListBox.Items.IndexOf("Name"), $true);

    if ($error_message) {

        <# The error message. #>
        $label4 = New-Object System.Windows.Forms.Label
        $label4.Location = New-Object System.Drawing.Point(10,0)
        $label4.Size = New-Object System.Drawing.Size(280,20)
        $label4.Text = $error_message
        $label4.ForeColor = 'red'
        $form.Controls.Add($label4)
    }

    $form.Topmost = $true
    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($group.Text) {
            show-search_results (get-members $group.text) $CheckedListBox.CheckedItems
        } else {
            show-form_search "Please enter a valid group name."
        }
    } else {
        show-form_1
    }
}

function show-form_add($error_message, $passed_group, $passed_name) {

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

    <# The room number prompt. #>
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10,20)
    $label1.Size = New-Object System.Drawing.Size(280,20)
    $label1.Text = 'Active Directory Group:'
    $form.Controls.Add($label1)

    <# The room number text box. #>
    $group = New-Object System.Windows.Forms.TextBox
    $group.Location = New-Object System.Drawing.Point(10,40)
    $group.Size = New-Object System.Drawing.Size(260,20)
    $group.text = $passed_group
    $form.Controls.Add($group)

    <# The notes prompt. #>
    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(10,80)
    $label2.Size = New-Object System.Drawing.Size(280,20)
    $label2.Text = 'Comma-separated Gordon Login/s (first.last):'
    $form.Controls.Add($label2)

    <# The notes text box. #>
    $name = New-Object System.Windows.Forms.TextBox
    $name.Location = New-Object System.Drawing.Point(10,100)
    $name.Size = New-Object System.Drawing.Size(260,80)
    $name.text = $passed_name
    $form.Controls.Add($name)

    if ($error_message) {

        <# The error message. #>
        $label3 = New-Object System.Windows.Forms.Label
        $label3.Location = New-Object System.Drawing.Point(10,0)
        $label3.Size = New-Object System.Drawing.Size(280,20)
        $label3.Text = $error_message
        $label3.ForeColor = 'red'
        $form.Controls.Add($label3)
    }

    $form.Topmost = $true
    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        if ($group.text -AND $name.text) {
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

function show-form_remove($error_message, $passed_group, $passed_name) {
    
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    <# The form panel to be displayed. #>
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Remove Active Directory Group Members'
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

    <# The room number prompt. #>
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10,20)
    $label1.Size = New-Object System.Drawing.Size(280,20)
    $label1.Text = 'Active Directory Group:'
    $form.Controls.Add($label1)

    <# The room number text box. #>
    $group = New-Object System.Windows.Forms.TextBox
    $group.Location = New-Object System.Drawing.Point(10,40)
    $group.Size = New-Object System.Drawing.Size(260,20)
    $group.text = $passed_group
    $form.Controls.Add($group)

    <# The notes prompt. #>
    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(10,80)
    $label2.Size = New-Object System.Drawing.Size(280,20)
    $label2.Text = 'Comma-separated Gordon Login/s (first.last):'
    $form.Controls.Add($label2)

    <# The notes text box. #>
    $name = New-Object System.Windows.Forms.TextBox
    $name.Location = New-Object System.Drawing.Point(10,100)
    $name.Size = New-Object System.Drawing.Size(260,80)
    $name.text = $passed_name
    $form.Controls.Add($name)

    if ($error_message) {

        <# The error message. #>
        $label3 = New-Object System.Windows.Forms.Label
        $label3.Location = New-Object System.Drawing.Point(10,0)
        $label3.Size = New-Object System.Drawing.Size(280,20)
        $label3.Text = $error_message
        $label3.ForeColor = 'red'
        $form.Controls.Add($label3)
    }

    $form.Topmost = $true
    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {
        if ($group.text -AND $name.text) {
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

function show-form_export($data, $error_message) {

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

    <# The building prompt. #>
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10,20)
    $label1.Size = New-Object System.Drawing.Size(280,20)
    $label1.Text = 'What would you like the file to be saved as?:'
    $form.Controls.Add($label1)

    <# The building text box. #>
    $path = New-Object System.Windows.Forms.TextBox
    $path.Location = New-Object System.Drawing.Point(10,40)
    $path.Size = New-Object System.Drawing.Size(260,20)
    $path.text = "D:\Scripts\ActiveDirectory\export.csv"
    $form.Controls.Add($path)

    if ($error_message) {

        <# The error message. #>
        $label4 = New-Object System.Windows.Forms.Label
        $label4.Location = New-Object System.Drawing.Point(10,0)
        $label4.Size = New-Object System.Drawing.Size(280,20)
        $label4.Text = $error_message
        $label4.ForeColor = 'red'
        $form.Controls.Add($label4)
    }

    $form.Topmost = $true
    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($path.text) {
            Write-Host "Building csv..."
            remove-item $path.text
            foreach ($d in $data) {
                export-csv -Path $path.text -Append -InputObject $d
                Write-Host "Exported line: ["$d.name"]."
            }
            show-form_1 "Data exported successfully."

        } else {
            return show-form_export
        }
    } else {
        show-form_1
    }
}

$index = 0
mkdir D:\Users\$env:username\Documents\ActiveDirectory -ErrorAction 'ignore'
$global:file_path = "D:\Users\$env:username\Documents\ActiveDirectory\ADlog_0.txt"
while ((Test-Path $global:file_path)) {
    $index += 1
    $global:file_path = "D:\Users\$env:username\Documents\ActiveDirectory\ADlog_" + $index + ".txt"
}

New-Item -Path $global:file_path -ItemType File | out-null
Add-content $global:file_path ("[" + (get-date) + "]: Starting log of Active Directory Utility.")
Write-Host "Log file stored as '$global:file_path'"

show-form_1

Add-content $global:file_path ("[" + (get-date) + "]: End of Log File`n")