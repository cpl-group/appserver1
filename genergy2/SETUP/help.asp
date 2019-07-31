<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim notoolbar
if not(allowGroups("Genergy Users,clientOperations")) then
notoolbar = 1
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN"
        "http://www.w3.org/TR/1999/REC-html401-19991224/strict.dtd">
<html>
<head>
	<title>Utility Manager Help</title>
	<style type="text/css">
	dt { text-indent:12px;font-weight:bold;margin-top:12px; }
	</style>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>
<body>
<table border=0 cellpadding="10" cellspacing="0">
<tr>
  <td>
  <p>


  <% if secureRequest("page") = "portfolioedit" then %>
  <b>Add/Edit Portfolio</b><br>
  Portfolio numbers are not editable once they have been saved to the system, so please be careful to enter a meaningful and brief designation the first time around. Portfolio names, on the other hand, may be changed at any time.
  <br><br>
  
  <span class="notetext">Note: Do not use single quote marks (') in any text entry field in Utility Manager.</span>
  <br><br>
  
  <hr size="1" noshade>
  <b>Manage Portfolio Groups</b><br>
  In general, groups are used to generate invoices for certain tenants, show load profiles for a selection of meters, or bundle items together for reports. 
  <br><br>
  
  <b>Manage Portfolio Contacts</b><br>
  Enter contact information for the portfolio holders. Building and account contacts are entered separately. 
  <br><br>
  
  <hr size="1" noshade>
  <b>Buildings</b><br>
  <p class="helptext" style="margin:10px;">
  HOW TO<br>
  Click &quot;Add Building&quot; to set up buildings under this portfolio.
  <br><br>
  Click a row in the building list to view/edit that building and assign it accounts (ie, tenants).
  </p>
  
   <% elseif secureRequest("page") = "bsrate" then %>
   <p><b>Building Specific Invoice Amount</b>
   <p>The purpose of this screen is to enter Invoice amounts that apply to all tenants and meters in a building for a particular utility. Below are some examples of how this screen could be used.
   <p> </p>
   <p>Example One &#8211; Hot Water Rates<br>
	Hot water rates vary per building based on how the building heats the water.  
   <p class="helptext" style="margin:10px;">
HOW TO<br>
  To enter a building specific hot water rate:
   </p>
   <p class="helptext" style="margin:10px;">Enter the Utility. If applicable, choose a season, otherwise choose &#8220;N/A.&#8221; Enter the &#8220;Date From&#8221; and &#8220;Date To&#8221; range that the rate covers. Choose the appropriate rate. If applicable choose a bill year and a bill period, otherwise choose &#8220;N/A&#8221; for both*. Enter the rate amount, entering percentages as decimals. Enter a note describing the rate entry <br>
   </p>
   <p>Example Two - Building adjustments or credits<br>
	An amount should be entered in this screen if it applies to all tenants and meters in the building for a particular utility. </p>
      
   <p class="helptext" style="margin:10px;">
HOW TO<br>
  To enter a building adjustment or credit:
   </p>
   <p class="helptext" style="margin:10px;">Enter the Utility. If applicable, choose a season, otherwise choose N/A . Enter the &#8220;Date From&#8221; and &#8220;Date To&#8221; range that the rate covers. Choose the appropriate rate. If applicable choose a bill year and a bill period, otherwise choose &#8220;N/A&#8221; for both*. Enter the rate amount, entering percentages as decimals. Enter a note describing the rate entry <br>
   </p>
   <p>Example Three &#8211; Utility Rate<br>
	An amount that, much like the Hot Water rate in example one, applies to a general rate but has specific numbers depending on the building. These amounts would be entered following the instructions above.</p>
   <p> </p>
   <p> </p>
   <p> </p>
   <p>--------------------------------------------------------------------------------</p>
   <p>* If a bill year/bill period combination in entered, you cannot enter a season or a date range<br>
   </p>
   <p><br>
  
    
	
	<% elseif secureRequest("page") = "1970rate" then %>
   <p><b>1970 Rate Parameters</b><br>
  This interface can be used to define the parameters used in the 1970 Rate. The fields used in this interface are:</p>
   <dt>Utility:<dd>The utility that the rate amount applies to.  This field is required.
   <dt>Season:<dd>The season that this rate will apply to.  Only the seasons set up for this region are available to chose from.  This field is required.
   <dt>Bill Year and Bill Period:<dd>The bill year and bill period that the 1970 rate increase is based on.  This field is required.
   <dt>Amount:<dd>The Amount of the rate increase.  This should be a percentage, and should be entered as a decimal. (i.e. 15% should be entered as .15)  This field is required.
   <dt>Note:
   <dd>A brief description of the invoice amount.  Please be descriptive.  This field is required.</p>

   <br>
    <% elseif secureRequest("page") = "buildingedit" then %>
  <b>Add/Edit Building</b><br>
  Building numbers are not editable once they have been saved to the system, so please be careful to enter a meaningful and brief designation the first time around. Building names may be changed at any time.
  <br>
	<br>
    <%if allowGroups("Genergy Users,clientOperations") then %>
  Regions are created in Rate Setup. <br>
  
  </p>
    <p class="helptext" style="margin:10px;">
HOW TO<br>
  To add re  gions, click "Set Up Rates" in the black navigation bar at the top of the page or "Rate Setup" in the Intranet navigation bar at left.
  </p>
	<% end if %>

  <b>Hide Info</b><br>
  Users with limited screen space might want to click this link to hide building information so that more accounts can be viewed on one screen. The link toggles the info display when clicked again.
  <br>
	<br>
  
    <hr size="1" noshade>
  <b>Manage Bill Periods For This Building</b><br>
  Bill periods are set up individually for each utility in a building.<br>
    <p class="helptext" style="margin:10px;">
  HOW TO<br>
  To add a bill period, click "Manage bill periods for this building", then click the button labeled "Add Bill Period". Enter info and save.
  <br>
	 <br>
  To review existing bill periods, click "Manage bill periods for this building", then select a utility to load its bill periods. 
  <br>
	 <br>
  To edit existing bill periods, follow the steps for reviewing, above, and click a bill period in the list.
  </p>
	<b>Manage Building Groups</b><br>
  In general, groups are used to generate invoices for certain tenants, show load profiles for a selection of meters, or bundle items together for reports. 
  <br>
	<br>
  
  <b>Manage Building Contacts</b><br>
  Enter contact information for the portfolio holders. Building and account contacts are entered separately. 
  <br>
	<br>
  
    <hr size="1" noshade>
  <b>Accounts</b><br>
    <p class="helptext" style="margin:10px;">
  HOW TO<br>
  Click &quot;Add Account&quot; to set up accounts (i.e., tenants) to this building.
  <br>
	 <br>
  Click an account in the Accounts list to review and edit it.
  </p>
	(Note: Accounts can't be added to a building until the building has been saved.)
  <br>
	<br>

    <% elseif secureRequest("page") = "groupview" then %>
  <b>Manage Groups</b><br>
    <ul>
  	 <li>Groups at the portfolio level can include meters from all buildings in the portfolio.
  	 <li>Groups at the building level can include meters from all accounts in the building.
  	 <li>Groups at the account level can include only meters from the same account.
    </ul>
	<p class="helptext" style="margin:10px;">
  HOW TO<br>
  Click "Add group" or click an existing group in the list to edit
  </p>
	<% elseif secureRequest("page") = "groupedit" then %>
  <b>Group Edit</b><br>

    <p class="helptext" style="margin:10px;">
  HOW TO<br>
  Select a category from the first pulldown to indicate how the group will be used
  <br>
	 <br>
  Enter a name for the group in the text field
  <br>
	 <br>
  Check the meters you want in the group
  <br>
	 <br>
  Click "Save"
  </p>
	<% elseif secureRequest("page") = "billperiodview" then %>
  <b>Manage Bill Periods</b><br>
  Bill periods are set up individually for each utility in a building.<br>
    <p class="helptext" style="margin:10px;">
  HOW TO<br>
  To add a bill period, click "Manage bill periods for this building", then click the button labeled "Add Bill Period". Enter info and save.
  <br>
	 <br>
  To review existing bill periods, click "Manage bill periods for this building", then select a utility to load its bill periods. 
  <br>
	 <br>
  To edit existing bill periods, follow the steps for reviewing, above, and click a bill period in the list.
  </p>
	<p>When finished setting up bill periods, return to the building page using the "breadcrumb trail" navigation at the top of the page. The breadcrumb trail provides links back to portfolio, building and account as available.</p>
	<% elseif secureRequest("page") = "tenantedit" then %>
  <b>Hide Info</b><br>
  This link in the upper right corner of the page toggles the display of account-related information on and off for better use of screen space on small monitors.
  <br>
	<br>
  
    <hr size="1" noshade>
  <b>Edit Account Information</b><br>
  You may change but not delete account information. 
    <%if allowGroups("Genergy Users,clientOperations") then %>
  To delete an account, please open a trouble ticket in the Trouble Tracker under IT Services on the intranet.
    <% else %>
  To delete an account, please contact your Genergy account manager.
    <%end if%>
  <br>
	<br>
  
    <hr size="1" noshade>
  <b>Transfer Info To New Account</b><br>
  When one tenant moves out and another moves in, it may be helpful to transfer the lease utility and meter information from the previous tenant to the new one. This button initiates a three-step process:
  </p>
  
    <ol>
  	 <li><b>Enter Account Details</b><br>
	  Create a new account and take the old one offline by marking "Lease Expired"
  	 <li><b>Assign New Lease Utilities</b><br>
	  Information on the previous account's lease utilities is displayed on the same screen where you enter the lease utilities for the new account
  	 <li><b>Transfer Meters</b><br>
	  The previous account's meters are listed on the screen so you know exactly what to assign to the new tenant. If you need to set up a new meter for the new account, you can do so after the transfer process is complete.
    </ol>
	<hr size="1" noshade>
  <b>Manage account groups</b><br>
	Select meters from this account to assign to a group. In general, groups are used to generate invoices for certain tenants, show load profiles for a selection of meters, or bundle items together for reports.
  
    <p><b>Custom links</b><br>
	 Quick Help for custom components of gEnergyOne may be found on the screens generated by these links. Custom component links are displayed to the right of any core component links, such as "Manage account groups".</p>
	<hr size="1" noshade>
  <b>Lease Utilities</b><br>
  Lease utilities define certain rates and fees associated with a given account. Because meters track the usage of a particular utility (electricity, gas, etc.) for a given account (lease), a &quot;lease utility&quot; must be defined before any meters can be set up.<br>
  
    <p class="helptext" style="margin:10px;">
  HOW TO<br>
  Click the &quot;Add Lease Utility&quot; button to the right of the Lease Utilities section to assign a new utility.
  <br>
	 <br>

  The links in the &quot;Jump to&quot; area quickly scroll the page down to the relevant lease utility when long lists of meters make scrolling inconvenient.
  <br>
	 <br>
  
  Click &quot;Show Info&quot; to see the rates and fees of an existing lease utility
  </p>
	<hr size="1" noshade>
  <b>Meters</b><br>
  Meters are listed by lease utility. Click on a listed meter to edit it.


    <% elseif secureRequest("page") = "tenanttransfer1" then %>
  <b>Transfer Account</b>
  When one tenant moves out and another moves in, it may be helpful to transfer the lease utility and meter information from the previous tenant to the new one. This is a three-step process.
  </p>
  
  <b>Enter Account Details</b><br>
    <p class="helptext" style="margin:10px;">
  HOW TO<br>
  Take the previous account offline by marking "Lease Expired" in the first column.
  <br>
	 <br>
  Enter details for the new account in the second column. Hit "Continue" to save.
  </p>
	<% elseif secureRequest("page") = "tenanttransfer2" then %>
  <b>Assign New Lease Utilities</b><br>
  Lease utilities are not copied over from the previous account, but they are displayed for reference when assigning lease utilities to the new account. Otherwise, creating lease utilities for a transfer is no different from adding lease utilities to a new account.
  <br>
	<br>
    <p class="helptext" style="margin:10px;">
  When you have finished setting up lease utilities, click "Transfer Meter" in a lease utility row to move on to the next step in the process: assigning meters to the lease utility.
  </p>
	<% elseif secureRequest("page") = "tenanttransfer3" then %>
  <b>Transfer Meters</b><br>
  Meters are not copied over from the previous account, but they are displayed for reference when assigning meters to the new lease utilities for this account.
  <br>
	<br>
  When a meter is "transferred", it is marked off-line in its former account and is assigned to the new one. You can choose to transfer historical data along with the meter by selecting a bill period from the pulldown that appears once you have chosen a meter to transfer.
  <br>
	<br>
    <p class="helptext" style="margin:10px;">
  When you have finished transferring meters, use the breadcrumb trail navigation at the top of the page to navigate back to the portfolio, building or account you want to set up next.
  </p>
	<% elseif secureRequest("page") = "leaseutilityedit" then %>
  <b>Add/Update Lease Utility</b><br>
  Lease utilities define rates and fees to apply to meters associated with a given account. Because meters track the usage of a particular utility for a given account (lease), a &quot;lease utility&quot; must be defined before any meters can be set up.<br>
  <br>
    <dl>
  	 <dt>Account Rate
     <dd>
      <%if allowGroups("Genergy Users,clientOperations") then %>
    Rates created for this account in Rate Setup appear in the account rate pulldown. Use the "Edit Rates" button to jump to the rate setup screen for this account.
      <%else%>
    If you don't see an applicable rate in the account rate pulldown, please contact your Genergy representative.
      <%end if%>
  	 <dt>Rate Function
	 <dd>How the rate applies to usage in order to calculate cost
    </dl>
	<% elseif secureRequest("page") = "meteradd" or secureRequest("page") = "meteredit" then %>
    <% if secureRequest("page") = "meteradd" then %>
  <b>Transfer An Existing Meter</b><br>
  Meters can be assigned to a different account within a single building. When a meter is "transferred", it is marked off-line in its former account and is assigned to the new one. You can choose to transfer historical data along with the meter by selecting a bill period from the pulldown that appears once you have chosen a meter to transfer.
  <br>
	<br>
  Meters can't be transferred between different buildings. A message indicating that no meters are available for transfer means that no meters have been set up yet.
  <br>
	<br>
  
    <hr size="1" noshade>
  <b>Add New Meter</b><br>
    <%else%>
  <b>Update Meter</b><br>
    <% end if %>
  A meter name is required to create a new meter. All other fields depend on the particulars of the meter's use.
    <dl>
  	 <dt>Meter Name
	 <dd>Required.
  	 <dt>On Line
	 <dd>The meter is or is not collecting data.
  	 <dt>No Billing
	 <dd>This meter will not be used for invoicing.
  	 <dt>Start Date
	 <dd>The date the meter begins collecting data for this account.
  	 <dt>Date Off
	 <dd>The date the meter stops or will stop collecting data for this account.
  	 <dt>Date Last Read
	 <dd>This value is updated by the system whenever a meter reading has been entered. If there are no readings for a given meter, the system will put in a placeholder date (1/1/1900).
  	 <dt>Factor
	 <dd>
  	 <dt>Usage.x
	 <dd>Consumption multiplier. For electricity, usage would be kWh.
  	 <dt>Capacity.x
	 <dd>Demand multiplier. For electricity, capacity would be kW.
  	 <dt>Variance
	 <dd>Adjustment factor for system readings.
  	 <dt>Meter Reference, Power Feed
	 <dd>Pulldowns appear when a virtual meter has been set up for custom monitoring.
  	 <dt>Location
	 <dd>Something descriptive, for example, "Electrical closet".
  	 <dt>Floor
	 <dd>Can be multiple floors.
  	 <dt>Riser
	 <dd>Optional.
  	 <dt>Datasource
	 <dd>
  	 <dt>Define Data Fields
	 <dd>
    </dl>
	<b>Meter LMP</b><br>
  Pops up the meter's current monthly load profile as available.
    <% end if %>
    <hr size="1" noshade>
    <p align="center">
	 <input type="button" name="close" value="Close Window" onclick="self.close();" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
	</p>
  </td>
</tr>
</table>
</body>
</html>
