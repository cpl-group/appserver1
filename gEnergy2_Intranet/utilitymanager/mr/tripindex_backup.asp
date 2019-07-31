<%@Language="VBScript"%>
<% '1/8/2007 N.Ambo changed this page from an htm page to an asp page and added functionality for 
'the user to enter the bill year together with the trip code and billperiod which already existed. 
'Prior there was no billyear filed and the system always assumed the billyear was the current year.
'However there are some buildings with two of the same periods (e.g., 12/2007, 12/2008) adn the system 
'will only show the 12/2008 period. Hence, in order to allow the users to view readings for 12/2007 for example, 
'an input field has now been put in place for the bill year
'In addition, a validity check has been added to return a message to the user if there were no values entered 
'in the fields.
%>
<HTML>
	<HEAD>
		<title>Trip Code Setup</title> 
		<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
		<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
		
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
			<style type="text/css"> .tblunderline { border-bottom:1px solid #cccccc; }
	</style>
	
<%
dim curryear 
curryear = Year(now())
%>
<script>

function loadtrips(tripcode, billperiod){
document.frames.tripsheet.location  = "tripcodes.asp?tripcode="+tripcode+"&billperiod="+billperiod
}
function opentripsheet(tripcode, billyear,billperiod,extended){
	var w = 800;
	var h = 600;
	var page = "tripsheet.asp?tripcode="+tripcode+"&billyear="+billyear+"&billperiod="+billperiod+"&extended="+extended
    winprops = 'height='+h+',width='+w+',status=yes, scrollbars=yes'
    if (tripcode=="" || billperiod=="" || billyear == "")
    {
		alert("You must enter a value for all three fields.");
	}
	else
	{
		window.open(page,'tripsheetprint',winprops);
	}
	
}
function exec_umsync(){
	var w = 200;
	var h = 100;
	var page = "umsync.asp"
    winprops = 'height='+h+',width='+w+',status=yes, scrollbars=yes'
	window.open(page,'umsync',winprops)
}
			</script>
	</HEAD>
	<body bgcolor="#eeeeee" leftmargin="0" topmargin="0">
		<form name="form1" method="post" action="">
			<table width="100%" border="0" cellspacing="0" cellpadding="2">
				<tr>
					<td bgcolor="#6699cc" class="standardheader">Trip Sheet Access</td>
				</tr>
				<tr>
					<td><strong>Enter Trip</strong></td>
				</tr>
				<tr>
					<td width="50%" valign="middle">Trip Code : <input name="tripcode" type="text" size="3" maxlength="2" title="Trip code">
						Bill Period: <input name="billperiod" type="text" size="3" maxlength="2" title="Trip Date">
						Bill Year: <input name="billyear" type="text" size="5" maxlength="4" title="Trip Year" value= <%=curryear%>>
					</td>
				</tr>
				<tr>
					<td>
						<input type="button" name="Button22" value="Print Tripsheet" onclick="opentripsheet(tripcode.value, billyear.value, billperiod.value, 'false')">
						<input type="button" name="Button222" value="Print Extended Tripsheet" onclick="opentripsheet(tripcode.value, billyear.value, billperiod.value,'true')">
					</td>
				</tr>
				<tr>
					<td><iframe name="tripsheet" src="./tripcodes.asp" frameborder="0" width="100%" height="450"></iframe>
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
