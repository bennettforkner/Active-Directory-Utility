# Active-Directory-Utility

This program is made to allow end-users on Gordon College's campus add AD members to certain groups, particularly for emailing purposes.
<br><br><b>Access Prerequisites:</b>
<ul><li>The user must be using a machine that is running a Windows operating system.
<li>The user must be connected to the 'gordon.edu' domain (via VPN or direct internet connection).
<li>The user must be authorized to use this tool (i.e. they must be a member of the AD group 'enduser-list' which is not editable using this utility).
  </ul>
<br><b>Documentation:</b>
<br><br><b>User Flow:</b>
<ul>
  <li>
    <b>Action Choice Window:</b> If the prerequisites mentioned above are met, the user will be presented with a dialogue window asking them to select an action. The choice is presented as three horizontally-aligned buttons: ADD, REMOVE, and SEARCH. When a button is pressed, the action choice window is closed and the respective functionality detailed below is enacted.
<ul>
  <li>
    <b>ADD:</b> The user is then presented with a new dialogue window with prompts for an Active Directory group name and the 'sAMAccountName' (username) of the 'AD user' or users to be added to the specified group. When the user has entered the information to fill both of these fields, they may press the add button to enact that button's functionality or the 'x' window button to close the window and re-open the Action Choice Window.
<ul><li>
  <b>AD Group Name:</b> This field must be filled with a valid and complete group name from the AD domain. This provides a light layer of security and closure in that the group names are not displayed or searchable to the user. The user should already know the name of the group that they are editing, o/w they should inquire CTS.
  </li><li>
  <b>AD User Name:</b> This field can be filled with one valid and complete username or multiple valid and complete usernames delimited by a sole comma (,). This field must not be empty.
  </li><li>
<b>ADD Button:</b> This button will close the window and call another function to add the specified user to the group. That function will then display the SEARCH window with the updated member list from the specified group.
  </li>
    </ul>
</li>
  <li> <b>REMOVE:</b> The user is then presented with a new dialogue window with prompts for an Active Directory group name and the 'sAMAccountName' () of the 'AD user' or users to be added to the specified group. When the user has entered the information to fill both of these fields, they may press the add button to enact that button's functionality or the 'x' window button to close the window and re-open the Action Choice Window.
<ul><li>
  <b>AD Group Name:</b> This field must be filled with a valid and complete group name from the AD domain. This provides a light layer of security and closure in that the group names are not displayed or searchable to the user. The user should already know the name of the group that they are editing, o/w they should inquire CTS.
  </li><li>
  <b>AD User Name:</b> This field can be filled with one valid and complete username or multiple valid and complete usernames delimited by a sole comma (,). This field must not be empty.
  </li><li>
  <b>REMOVE Button:</b> This button will close the window and call another function to remove the specified user from the specified group. That function will then displa the SEARCH window with the updated member list from the specified group.
  </li>
    </ul>
</li><li>
<b>SEARCH:</b> The user is then presented with a new dialogue window with a prompt for an Active Directory group name to be displayed as well as a select option box with a list of possible user fields to be displayed. On click of the SEARCH button, the functionality will be executed, o/w the user can click the 'x' button and return to the Action Choice Window.
<ul><li>
  <b>AD Group Name:</b> This field must be filled with a valid and complete group name from the AD domain. This provides a light layer of security and closure in that the group names are not displayed or searchable to the user. The user should already know the name of the group that they are editing, o/w they should inquire CTS.
  </li><li>
  <b>Field Select Option Box:</b> This box provides a multiple-select choice option for the user to select the fields that they want to see about the users in the specified group. The options include: name, email, sAMAccountName, and SID.
  </li><li>
  <b>SEARCH Button:</b> This button will close the window and call another function to display the members of the specified group with the specified fields in a formatted table. More information about this display is found below.
<ul><li>
  <b>Search Results Window:</b> This window displays specified field information about the users who are members of a particular group.
  <ul><li>
  <b>Add Member Input Field:</b> This field allows the user to quickly specify a user to be added to the currently displayed group.
  </li><li>
  <b>ADD Button:</b> This button adds the user to the currently displayed group by calling the ADD functionality.
  </li><li>
  <b>EXPORT Button:</b> This button closes this window and allows the user to export the data from a group list in a csv format while specifying the file path.
  </li><li>
  <b>REMOVE *USER* Button:</b> These buttons (per member listed) allow the user to quickly remove a user from the AD group by calling the REMOVE functionality.
  </li><li>
  <b>VIEW *USER* Button:</b> This button opens a new window with the user detailed information.
  </li>
  </ul>
  </li>
  </ul>
  </li>
  </li>
    </li>
</ul>
