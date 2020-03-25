
. D:\Scripts\ActiveDirectory\data_views.ps1

function get-members($group) {
    Add-content $global:file_path ("[" + (get-date) + "]: Getting Active Directory members from group: $group.")
    try {return get-adgroupmember $group} catch [Exception] {show-form_1 $_.Exception.message}
}

function add-members($group, $member) {
    if ($member.indexof(',')) {
        $member = $member.split(',')
    }
    Add-ADGroupMember $group $member
    $member = $member.toupper()
    Add-content $global:file_path ("[" + (get-date) + "]: Added members: $member to Active Directory group: $group.")
    show-search_results (get-members $group) @("name","email") "" "Updated Group:"
}

function remove-members($group, $member) {
    if ($member.indexof(',')) {
        $member = $member.split(',')
    }
    remove-adgroupmember $group $member -confirm:0
    $member = $member.toupper()
    Add-content $global:file_path ("[" + (get-date) + "]: Removed members: $member from Active Directory group: $group.")
    show-search_results (get-members $group) @("name","email") "" "Updated Group:"
}