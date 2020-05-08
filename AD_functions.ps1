<# AD_functions.ps1:
    File to compile all of the functionality that interacts with the Active Directory utility commands.
    
    Author: Bennett Forkner (CTS Intern, Gordon College)
    Date: 4/6/20 #>

<# Imports to include other files in this file's scope. #>
. D:\Scripts\ActiveDirectory\data_views.ps1

<# get-members:
    function to query the active directory to get members in a particular group.
    @param $group: the group that will be queried.
    @return the list of member objects in group $group. #>
function get-members($group) {

    <# Update logs. #>
    $log = ("[" + (get-date) + "]: Getting Active Directory members from group: $group.")
    Add-content $global:file_path $log

    <# return surrounded by error catching, resulting in return to start form with error. #>
    try {return get-adgroupmember $group} catch [Exception] {show-form_1 $_.Exception.message}
}

<# addmembers:
    function to link adusers to a specific group as a resemblance of membership.
    @param $group: the group that the user/s will be added to.
    @param $member: the user/s that will be added to group $group. #>
function add-members($group, $member) {

    <# Split $members into an array if it is a list of adusers (comma delimited). #>
    if ($member.indexof(',')) {
        $member = $member.split(',')
    }

    <# action surrounded by error catching, resulting in return to start form with error. #>
    try {add-adgroupmember $group $member} catch [Exception] {show-form_1 $_.Exception.message}
    $member = $member.toupper()

    <# Update logs. #>
    $log = ("[" + (get-date) + "]: Added members: $member to Active Directory group: $group.")
    Add-content $global:file_path $log

    <# Show the updated list of members in the referenced $group. #>
    show-search_results (get-members $group) @("name","email") "" "Updated Group:" $group
}

<# remove-members:
    function to remove the link between ad users and the specified $group.
    @param $group: the group that the $member will no longer be tied to.
    @param $member: the aduser name that will be removerd from the $group. #>
function remove-members($group, $member) {

    <# Split $members into an array if it is a list of adusers (comma delimited). #>
    if ($member.indexof(',')) {
        $member = $member.split(',')
    }

    <# action surrounded by error catching, resulting in return to start form with error. #>
    try {remove-adgroupmember $group $member -confirm:0} catch [Exception] {show-form_1 $_.Exception.message}
    $member = $member.toupper()

    <# Update logs. #>
    $log =  ("[" + (get-date) + "]: Removed members: $member from Active Directory group: $group.")
    Add-content $global:file_path $log

    <# Show the updated list of members in the referenced $group. #>
    show-search_results (get-members $group) @("name","email") "" "Updated Group:" $group
}