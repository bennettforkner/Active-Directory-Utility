<# data_views.ps1:
    file to collect all of the forms that display Active Directory data to the user. 
    
    Author: Bennett Forkner (CTS Intern, Gordon College)
    Date: 4/6/20 #>

<# show-search_results:
    function to display the results of a group search or query to the user. 
    @param $results: the ad users to be displayed as the result of a query to the Active Directory.
    @param $fields: the user-chosen fields to be displayed about the ad users.
    @param $error_message: the message, if not null, to be displayed on the form as the result of a previous error.
    @param $title: the header of the form to be shown to the user. ex: 'Search Results'.
    @return to (show-form_1 on normal exit | show-user_view on view button click | show-form_export on 'export' button click). #>
function show-search_results($results, $fields, $error_message, $title, $passed_group) {
    
    <# Check if a custom $title was entered. If not, set $title to default. #>
    if (!$title) {
        $title = "Search Results"
    }

    <# Type imports. #>
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    <# $num_lines: the number of fields that will be displayed for each user. #>
    $num_lines = $fields.count

    <# $size_y: the size displacement for each user-block. #>
    $size_y = ($num_lines * 10) + 20
    if ($size_y -lt 40) {
        $size_y = 40
    }

    <# $form_height: the total height of the form. #>
    $form_height = (($size_y + 10) * $results.count) + 140

    <# Limit the form to 800px so as to not extend too far. #>
    if ($form_height -gt 800) {
        $form_height = 800
    }

    <# The form panel to be displayed. #>
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Group View"
    $form.Size = New-Object System.Drawing.Size(810,($form_height))
    $form.StartPosition = 'CenterScreen'
    $form.AutoScroll = $true

    <# The heading of the form display. #>
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10,20)
    $label1.Size = New-Object System.Drawing.Size(150,20)
    $label1.Text = $title
    $label1.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10)
    $form.Controls.Add($label1)

    $script:add_user = 0

    <# The add-member prompt. #>
    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(160,20)
    $label2.Size = New-Object System.Drawing.Size(100,20)
    $label2.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10)
    $label2.Text = 'Add Member/s:'
    $form.Controls.Add($label2)

    <# The 'user to be added' text box. #>
    $user = New-Object System.Windows.Forms.TextBox
    $user.Location = New-Object System.Drawing.Point(260,20)
    $user.Size = New-Object System.Drawing.Size(200,20)
    $user.text = ""
    $form.Controls.Add($user)

    <# The 'add-member' button. On click, close this form and open the show-form_export form (logic at end of function). #>
    $addButton = New-Object System.Windows.Forms.Button
    $addButton.Location = New-Object System.Drawing.Point(470,20)
    $addButton.Size = New-Object System.Drawing.Size(75,23)
    $addButton.Text = "ADD USER"
    $addButton.add_click({$script:add_user = 1; $form.close()})
    $form.Controls.Add($addButton)


    <# $export: variable to store if the export button has been pressed as form submit. #>
    $script:export = $false

    <# The 'export' button. On click, close this form and open the show-form_export form (logic at end of function). #>
    $exportButton = New-Object System.Windows.Forms.Button
    $exportButton.Location = New-Object System.Drawing.Point(700,20)
    $exportButton.Size = New-Object System.Drawing.Size(75,23)
    $exportButton.Text = "EXPORT"
    $exportButton.add_click({$script:export = 1; $form.close()})
    $form.Controls.Add($exportButton)

    <# The 'close' button. #>
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(615,20)
    $closeButton.Size = New-Object System.Drawing.Size(75,23)
    $closeButton.Text = "CLOSE"

    <# Logic to close the form when this button is pressed. #>
    $closeButton.Add_Click({
        $form.Close()
    })

    $form.Controls.Add($closeButton)
    
    <# $location_y: the vertical location of the current user-block (incremented at each iteration fo the loop). #>
    $location_y = 50
    
    <# $script:user_to_view: the script-wide variable that stores which user view will be opened when a 'view' button is pressed. #>
    $script:user_to_view = ""

    <# $script:user_to_view: the script-wide variable that stores which user view will be opened when a 'view' button is pressed. #>
    $script:user_to_remove = ""

    <# loop:
        @condition: for all users in $results.
        @exit: all of the entries in $results have been displayed to the user. #>
    foreach ($value in $results) {
        
        <# The 'view' button. Later calls show-user_view to display further details about the respective user. #>
        $viewButton = New-Object System.Windows.Forms.Button
        $viewButton.Location = New-Object System.Drawing.Point(650,$location_y)
        $viewButton.Size = New-Object System.Drawing.Size(125,($size_y))
        $viewButton.Text = "VIEW `n" + $value.SamAccountName

        <# Functionality to set the user that will be viewed if the button is clicked. Show-form_user called at end of funtion. #>
        $viewButton.Add_Click({$script:user_to_view = $this.text.Substring(6); $form.close()})

        $form.Controls.Add($viewButton)


        
        <# The 'remove' button. Later calls show-user_view to display further details about the respective user. #>
        $removeButton = New-Object System.Windows.Forms.Button
        $removeButton.Location = New-Object System.Drawing.Point(525,$location_y)
        $removeButton.Size = New-Object System.Drawing.Size(125,($size_y))
        $removeButton.Text = "REMOVE `n" + $value.SamAccountName

        <# Functionality to set the user that will be viewed if the button is clicked. Show-form_user called at end of funtion. #>
        $removeButton.Add_Click({$script:user_to_remove = $this.text.Substring(8); $form.close()})

        $form.Controls.Add($removeButton)



        <# The field title. Filled in loop. #>
        $label1 = New-Object System.Windows.Forms.Label
        $label1.Location = New-Object System.Drawing.Point(15,$location_y)
        $label1.Size = New-Object System.Drawing.Size(130,$size_y)
        $label1.TextAlign = "MiddleRight"
        $label1.BackColor = "#D0D0D0"
        $label1.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)
        $label1.Padding = 0
        $label1.Padding.Right = 10

        $label1.text = ""

        <# The field value. Filled in loop. #>
        $label2 = New-Object System.Windows.Forms.Label
        $label2.Location = New-Object System.Drawing.Point(150,$location_y)
        $label2.Size = New-Object System.Drawing.Size(600,$size_y)
        $label2.TextAlign = "MiddleLeft"
        $label2.BackColor = "#D0D0D0"
        $label2.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 8)

        $label2.text = ""
        
        <# loop:
            @condition: for all fields in $fields.
            @entrance: label1 & label 2 are empty
            @invariant: the current field title and value are added to label1 and label2 respectively.
            @exit: $label1 & $label2 are filled with the values of all of the fields for the current user ($value). #> 
        foreach ($field in $fields) {

            <# $prop: the value of the current field. #>
            $prop = ""

            <# Call separate command to get users' email in the case that is a field. Else get the value from the given $field. #>
            try {
                if ($field -like "Email") {
                    $prop = (get-aduser $value.SamAccountName).userprincipalname
                } else {
                    $prop = $value."$field"
                }
                
                <# Concatenate values and field names into their respective labels. #>
                $field = $field.ToUpper()
                $label1.text += $field + ": `n"
                $label2.text += "$prop `n"
            } catch [Exception] {
                show-form_1 $_.Exception.message
            }
        }

        <# Add the labels into the form. #>
        $form.Controls.Add($label1)
        $form.Controls.Add($label2)
        
        <# Increment the $location_y for the next iteration. #>
        $location_y += $size_y + 10
    }

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

    <# Display form in the front of the screen. #>
    $form.AcceptButton = $addButton
    $form.Topmost = $true
    $result = $form.ShowDialog()

    <# Call the proper forms based on the user's button clicks. #>
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {} else {
        if ($script:user_to_view) {
            show-user_view $script:user_to_view
        } elseif ($script:user_to_remove) {
            remove-members $passed_group $script:user_to_remove
        } elseif ($script:add_user) {
            add-members $passed_group $user.text
        } elseif ($script:export) {
            show-form_export $results
        } else {
            show-form_1
        }
    }

}







<# show-user_view:
    function to show the data stored in the Active directory about a specific ad user.
    @param $user: the Sam Account name of the user to be displayed. #>
function show-user_view($user) {

    <# $user: the user object with the information to be displayed. #>
    $user = get-aduser $user
    <# $membership: the list of groups that this user si a member of. #>
    $membership = (get-adprincipalgroupmembership $user.SamAccountName).name

    <# Type imports. #>
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    <# $num_lines: the number of properties to be displayed. #>
    $num_lines = $user.PropertyCount

    <# $location_y: the height location of the current label set. #> 
    $location_y = 50

    <# $size_y: the size of the entire entry. #>
    $size_y = ($num_lines * 15) + 20

    <# $form_height: the height of the form, based upon the number of lines. #> 
    $form_height = (($size_y + 10) * $results.count) + 125

    <# limit the $form_height to 800px max. #>
    if ($form_height -gt 800) {
        $form_height = 800
    }

    <# The form panel to be displayed. #>
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $user.name + ' User View'
    $form.Size = New-Object System.Drawing.Size(850,500)
    $form.StartPosition = 'CenterScreen'
    $form.AutoScroll = $true

    <# The header of the form that will be shown at the top. #>
    $title = New-Object System.Windows.Forms.Label
    $title.Location = New-Object System.Drawing.Point(10,15)
    $title.Size = New-Object System.Drawing.Size(815,35)
    $title.Text = $user.name
    $title.TextAlign = "MiddleCenter"
    $title.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 20)
    $form.Controls.Add($title)


    <# The 'close' button #>
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(680,268)
    $closeButton.Size = New-Object System.Drawing.Size(75,23)
    $closeButton.Text = "CLOSE"

    <# Functionality to close the form when the 'close' button is pressed. #>
    $closeButton.Add_Click({
        $form.Close()
    })

    $form.Controls.Add($closeButton)
   
    <# The field type label. #>
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10,$location_y)
    $label1.Size = New-Object System.Drawing.Size(170,$size_y)
    $label1.TextAlign = "MiddleRight"
    $label1.BackColor = "#D0D0D0"
    $label1.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10.5)

    $label1.text = ""

    <# The field value label. #>
    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(185,$location_y)
    $label2.Size = New-Object System.Drawing.Size(630,$size_y)
    $label2.TextAlign = "MiddleLeft"
    $label2.BackColor = "#D0D0D0"
    $label2.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10.5)

    $label2.text = ""
    
    <# $fields: the names of the fields to be displayed. #>
    $fields = $user.propertynames

    <# $script:group_to_view: the group to pass to show_search_results if a group button is pressed. #>
    $script:group_to_view = $false






    <# loop:
        @condition: for all values in $fields.
        @entrance: $fields is initialized with the user-entered list of fields to display.
        @exit: the fields and their values have been set up. #>
    foreach ($field in $fields) {
        <# Separate logic to get the aduser's email, else get the field value. #>
        if ($field -like "Email") {
            $prop = (get-aduser $user.SamAccountName).userprincipalname
        } else {
            $prop = $user."$field"
        }
        <# Uppercase &  concatenate labels. #>
        $field = $field.ToUpper()
        $label1.text += $field + ": `n"
        $label2.text += "$prop `n"
    }
    
    <# Add both labels to the form. #>
    $form.Controls.Add($label1)
    $form.Controls.Add($label2)


    <# Commented-out to hide group-membership [for security's sake].


    # 'Group Membership' title.
    $title = New-Object System.Windows.Forms.Label
    $title.Location = New-Object System.Drawing.Point(10,230)
    $title.Size = New-Object System.Drawing.Size(640,35)
    $title.Text = "Group Membership"
    $title.TextAlign = "MiddleCenter"
    $title.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 15)
    $form.Controls.Add($title)

    # $index: the current number of the adgroup button out of the list.
    $index = 0

    # $location_x: the x-index of the current column of buttons. Incremented every 1/3 of the total buttons.
    $location_x = 20

    # $location_y: the current height of the button to be placed. Incremented at every iteration.
    $location_y = 270

    # loop:
        @condition: for all values in $membership.
        @entrance: $membership is initialized with the groups that this aduser is a part of.
        @exit: all of the buttons have been initialized and added to the $form.
    foreach ($group in $membership) {
        # The view group button.
        $viewButton = New-Object System.Windows.Forms.Button
        $viewButton.Location = New-Object System.Drawing.Point($location_x,$location_y)
        $viewButton.Size = New-Object System.Drawing.Size(200,20)
        $viewButton.text = $group
        $viewButton.Add_Click({$script:group_to_view = $this.text; $form.close()})

        $form.Controls.Add($viewButton)
        
        # Increment $index.
        $index += 1

        # Increment $location_y.
        $location_y += 25

        # Increment $location_x and set $location_y back to initial location if the end of the column has been reached.
        if ($index -eq ([math]::ceiling($membership.count / 3)) -OR $index -eq ([math]::ceiling($membership.count / 3) * 2)) {
            $location_x += 210
            $location_y = 270
        }
    }



    #>

    
    <# Display error message if passed. #>
    if ($error_message) {

        <# The error message. #>
        $label3 = New-Object System.Windows.Forms.Label
        $label3.Location = New-Object System.Drawing.Point(10,0)
        $label3.Size = New-Object System.Drawing.Size(280,20)
        $label3.Text = $error_message
        $label3.ForeColor = 'red'
        $form.Controls.Add($label3)
    }

    <# Display window and set it to the frontmost (topmost) window. #>
    $form.Topmost = $true
    $result = $form.ShowDialog()

    <# If form exits normally, show start form, else if a group button was clicked, view that group with show-search_results, else show start form. #>
    if ($result -eq [System.Windows.Forms.DialogResult]::OK)
    {show-form_1} else {
        if ($script:group_to_view) {
            show-search_results (get-members $script:group_to_view) @("name","email") "" $script:group_to_view
        } else {
            show-form_1
        }
    }
}