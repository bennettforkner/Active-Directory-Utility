function show-search_results($results, $fields, $error_message, $title) {

    if (!$title) {
        $title = "Search Results"
    }

    $script:user_to_view = ""

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $location_y = 50
    $num_lines = $fields.count
    $size_y = ($num_lines * 10) + 20
    $form_height = (($size_y + 10) * $results.count) + 140
    if ($form_height -gt 800) {
        $form_height = 800
    }

    <# The form panel to be displayed. #>
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Group View"
    $form.Size = New-Object System.Drawing.Size(800,($form_height))
    $form.StartPosition = 'CenterScreen'
    $form.AutoScroll = $true

    <# The building prompt. #>
    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10,20)
    $label1.Size = New-Object System.Drawing.Size(280,20)
    $label1.Text = $title
    $label1.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10)
    $form.Controls.Add($label1)

    <# The 'export' button.#>
    $exportButton = New-Object System.Windows.Forms.Button
    $exportButton.Location = New-Object System.Drawing.Point(700,10)
    $exportButton.Size = New-Object System.Drawing.Size(75,23)
    $exportButton.Text = "EXPORT"
    $exportButton.add_click({$script:export = 1; $form.close()})

    $form.AcceptButton = $exportButton
    $form.Controls.Add($exportButton)

    <# The 'close' button #>
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(615,10)
    $closeButton.Size = New-Object System.Drawing.Size(75,23)
    $closeButton.Text = "CLOSE"

    $closeButton.Add_Click({
        $form.Close()
    })

    $form.Controls.Add($closeButton)
    
    <# The SlotPort value.. #>
    $script:user_to_view = ""
    $script:export = $false
    foreach ($value in $results) {
        
        <# The 'ok' button.#>
        $viewButton = New-Object System.Windows.Forms.Button
        $viewButton.Location = New-Object System.Drawing.Point(650,$location_y)
        $viewButton.Size = New-Object System.Drawing.Size(125,($size_y))
        $viewButton.Text = "VIEW `n" + $value.SamAccountName
        $viewButton.Add_Click({$script:user_to_view = $this.text.Substring(6); $form.close()})

        $form.Controls.Add($viewButton)

        $label1 = New-Object System.Windows.Forms.Label
        $label1.Location = New-Object System.Drawing.Point(15,$location_y)
        $label1.Size = New-Object System.Drawing.Size(130,$size_y)
        $label1.TextAlign = "MiddleRight"
        $label1.BackColor = "#D0D0D0"
        $label1.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 8, [System.Drawing.FontStyle]::Bold)

        $label1.text = ""

        $label2 = New-Object System.Windows.Forms.Label
        $label2.Location = New-Object System.Drawing.Point(150,$location_y)
        $label2.Size = New-Object System.Drawing.Size(600,$size_y)
        $label2.TextAlign = "MiddleLeft"
        $label2.BackColor = "#D0D0D0"
        $label1.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 8)

        $label2.text = ""
        

        foreach ($field in $fields) {
            $prop = ""
            try {
            if ($field -like "Email") {
                $prop = (get-aduser $value.SamAccountName).userprincipalname
            } else {
                $prop = $value."$field"
            }

            $field = $field.ToUpper()
            $label1.text += $field + ": `n"
            $label2.text += "$prop `n"
            } catch {
            }
        }
    
        $form.Controls.Add($label1)
        $form.Controls.Add($label2)
        $location_y += $size_y + 10
    }

    
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

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {} else {
        if ($script:user_to_view) {
            show-user_view $script:user_to_view
        } elseif ($script:export) {
            show-form_export $results
        } else {
            show-form_1
        }
    }

}









function show-user_view($user) {

    $user = get-aduser $user
    $membership = (get-adprincipalgroupmembership $user.SamAccountName).name

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    $num_lines = $user.PropertyCount
    $location_y = 50
    $size_y = ($num_lines * 15) + 20
    $form_height = (($size_y + 10) * $results.count) + 125
    if ($form_height -gt 800) {
        $form_height = 800
    }

    <# The form panel to be displayed. #>
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $user.name + ' User View'
    $form.Size = New-Object System.Drawing.Size(850,500)
    $form.StartPosition = 'CenterScreen'
    $form.AutoScroll = $true

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

    $closeButton.Add_Click({
        $form.Close()
    })

    $form.Controls.Add($closeButton)
   

    $label1 = New-Object System.Windows.Forms.Label
    $label1.Location = New-Object System.Drawing.Point(10,$location_y)
    $label1.Size = New-Object System.Drawing.Size(170,$size_y)
    $label1.TextAlign = "MiddleRight"
    $label1.BackColor = "#D0D0D0"
    $label1.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10.5)

    $label1.text = ""

    $label2 = New-Object System.Windows.Forms.Label
    $label2.Location = New-Object System.Drawing.Point(185,$location_y)
    $label2.Size = New-Object System.Drawing.Size(630,$size_y)
    $label2.TextAlign = "MiddleLeft"
    $label2.BackColor = "#D0D0D0"
    $label2.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 10.5)

    $label2.text = ""
        
    $fields = $user.propertynames
    $script:group_to_view = $false
    foreach ($field in $fields) {

        if ($field -like "Email") {
            $prop = (get-aduser $user.SamAccountName).userprincipalname
        } else {
            $prop = $user."$field"
        }
        $field = $field.ToUpper()
        $label1.text += $field + ": `n"
        $label2.text += "$prop `n"
    }
    
    $form.Controls.Add($label1)
    $form.Controls.Add($label2)



    $title = New-Object System.Windows.Forms.Label
    $title.Location = New-Object System.Drawing.Point(10,230)
    $title.Size = New-Object System.Drawing.Size(640,35)
    $title.Text = "Group Membership"
    $title.TextAlign = "MiddleCenter"
    $title.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 15)
    $form.Controls.Add($title)

    $index = 0
    $location_x = 20
    $location_y = 270
    foreach ($group in $membership) {
        $viewButton = New-Object System.Windows.Forms.Button
        $viewButton.Location = New-Object System.Drawing.Point($location_x,$location_y)
        $viewButton.Size = New-Object System.Drawing.Size(200,20)
        $viewButton.text = $group
        $viewButton.Add_Click({$script:group_to_view = $this.text; $form.close()})

        $form.AcceptButton = $addButton
        $form.Controls.Add($viewButton)
        $index += 1
        $location_y += 25
        if ($index -eq ([math]::ceiling($membership.count / 3)) -OR $index -eq ([math]::ceiling($membership.count / 3) * 2)) {
            $location_x += 210
            $location_y = 270
        }
    }
    
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

    if ($result -eq [System.Windows.Forms.DialogResult]::OK -AND !$delete)
    {} else {
        if ($script:group_to_view) {
            show-search_results (get-members $script:group_to_view) @("name","email") "" $script:group_to_view
        } else {
            show-form_1
        }
    }
}