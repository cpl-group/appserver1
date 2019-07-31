<%option explicit%>
<html>
<head>
<title>Entry</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>
function acctselect(building,utility){
	var temp = "acctlist.asp?building=" + building+ "&utility="+utility
	window.open(temp,"selectaccount", "scrollbars=yes,width=1000, height=300, status=no" );
//	window.selectaccount.focus();
}
function ypid(id1,building,utility){
	var temp = "ypid.asp?acctid=" + id1 + "&building=" + building+ "&utility=" + utility
	window.open(temp,"ypid", "scrollbars=yes,width=500, height=300, status=no" );
}
function setup(building,utility){
	clearselections();
	var temp = "save.asp?building="+building+ "&utility=" + utility
	document.frames.entry.location=temp;
	
}

function editacct(acctid)
{	if(acctid.length>0)
	{	var temp = "editacct.asp?acctid=" +acctid
		document.frames.entry.location=temp;
	}else
	{	alert('Select an Account');
	}
}

function clearselections()
{	document.frames.entry.location='about:blank';
	document.form1.acctid.value=''
	document.all['accountdisplay'].innerText='No Account Selected'
	document.all['enterbillbutton'].style.visibility='hidden'
}

function loadportfolio()
{	var frm = document.forms['form1'];
	var newhref = "entry.asp?pid="+frm.pid.value;
	document.location.href=newhref;
}

function loadbuilding()
{	var frm = document.forms['form1'];
	var newhref = "entry.asp?pid="+frm.pid.value+"&building="+frm.building.value;
	document.location.href=newhref;
}

</script>
<body bgcolor="#FFFFFF" text="#000000">

<%
Dim cnn1, rst1, rst2, str1, str
Set cnn1 = Server.CreateObject("ADODB.connection")

cnn1.Open application("cnnstr_genergy1")

dim acctid, pid, building
acctid=Request.querystring("acctnum")
pid=Request.Querystring("pid")
building=Request.Querystring("building")


'response.write pid
'response.end
%>

<table width="100%" border="0">
<tr><td>
	<table width="100%" border="0" height="33">
    <tr bgcolor="#3399CC"><td><b><i><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Utility Bill</font></i></b></td></tr>
	</table>
<form name="form1" method="post" action="">           
    <table width="100%" border="0">
<%Set rst1 = Server.CreateObject("ADODB.recordset")
if not(trim(pid)="" and trim(building)<>"") then
	response.write "<tr><td width=""25%"" height=""27""><font face=""Arial, Helvetica, sans-serif"">Portfolio:</font></td><td width=""75%"" height=""27"">" 
	rst1.open "SELECT distinct portfolioid FROM buildings ORDER BY portfolioid", cnn1
	response.write "<select name=""pid"" onchange=""loadportfolio()""><option value="""">Select Portfolio</option>"
	do until rst1.eof
		%><option value="<%=trim(rst1("portfolioid"))%>"<%if trim(rst1("portfolioid"))=trim(pid) then response.write " SELECTED"%>><%=rst1("portfolioid")%></option><%
		rst1.movenext
	loop
	rst1.close
	response.write "</select></td></tr>"
	response.write "<tr><td width=""25%"" height=""27""><font face=""Arial, Helvetica, sans-serif"">Building:</font></td><td width=""75%"" height=""27"">"
	if trim(pid)<>"" then
		response.write "<select name=""building"" onchange=""loadbuilding();"">"
		rst1.open "SELECT BldgNum, strt FROM buildings WHERE portfolioid='"&pid&"' ORDER BY strt", cnn1
		do until rst1.eof
			%><option value="<%=trim(rst1("Bldgnum"))%>"<%if trim(rst1("Bldgnum"))=trim(building) then response.write " SELECTED"%>><%=rst1("strt")%> (<%=rst1("Bldgnum")%>)</option><%
			rst1.movenext
		loop
		rst1.close
		response.write "</select></td></tr>"
	else%>
		<font face="Arial, Helvetica, sans-serif"><input type="hidden" name="building" value="">No Building Selected</font>
	<%end if
else
	response.write "<tr><td width=""25%"" height=""27""><font face=""Arial, Helvetica, sans-serif"">Building:</font></td><td width=""75%"" height=""27"">"
	rst1.open "SELECT BldgNum, strt FROM buildings WHERE bldgnum='"&building&"' ORDER BY strt", cnn1
	if not rst1.eof then response.write "<font face=""Arial, Helvetica, sans-serif""><input type=""hidden"" name=""building"" value="""&building&""">"&rst1("strt")&"</font>"
	response.write "</td></tr>"
end if



'			if trim(building)<>"" and trim(pid)<>"" then
'				str1="select * from buildings where portfolioid='"&pid&"' and bldgnum='"&trim(building)&"'"
'				rst1.Open str1, cnn1, 0, 1, 1
'				response.write rst1("strt")&", "&rst1("bldgnum")
'				response.write "<input type=""hidden"" name=""building"" value="""&rst1("bldgnum")&""">"
'				rst1.close
'			elseif trim(pid)="all" then
'				response.write "<select name=""building"" onchange=""clearselections();"">"
'				str1="select * from buildings order by strt"
'				rst1.Open str1, cnn1, 0, 1, 1
'				do until rst1.eof
'				%
'				<option value="<%=rst1("bldgnum")%" selected><%=rst1("strt")%, <% %<%=rst1("bldgnum")%</option>
'				<%
'				rst1.movenext
'			  loop
'			  rst1.close
'			elseif trim(pid)<>"" then
'				response.write "<select name=""building"" onchange=""clearselections();"">"
'				str1="select * from buildings where portfolioid='"&pid&"' order by strt"
'				rst1.Open str1, cnn1, 0, 1, 1
'				do until rst1.eof
'				%
'				<option value="<%=rst1("bldgnum")%" selected><%=rst1("strt")%, <% %<%=rst1("bldgnum")%</option>
'				<%
'				rst1.movenext
'			  loop
'			  rst1.close
'		  end if
			  %>
			  </select>
            </td>
            </tr>
			 <tr><td width="25%" height="26"><font face="Arial, Helvetica, sans-serif">Utility :</font></td>
			  <td width="75%">
              <select name="utility" onchange="clearselections();">
              <% 
			  Set rst2 = Server.CreateObject("ADODB.recordset")
			  str="select * from tblutility order by utility"
			  rst2.Open str, cnn1, 0, 1, 1
			  do until rst2.eof
			  %>
			  <option value="<%=rst2("utility")%>" <%if lcase(rst2("utility"))="electricity" then response.write "SELECTED"%>><%=rst2("utilitydisplay")%></option>
			  <%
			  rst2.movenext
			  loop
			  rst2.close
			  %>
			  </select>
			</td>
            </tr> 
			
			<tr>
				<td width="25%" height="26" valign="top"><font face="Arial, Helvetica, sans-serif">Account Number :</font></td>
				<td width="75%"><font face="Arial, Helvetica, sans-serif">
					<input type="hidden" name="id1">
					<input type="hidden" name="acctid">
					<span id="accountdisplay">No Account Selected</span><br>
					<input type="button" name="Submit" value="Select Account" onClick="acctselect(building.value,utility.value)">
					<input type="button" name="this" value="Setup New Account" onClick="setup(building.value,utility.value)">
					<span id="enterbillbutton" style="visibility:hidden"><input type="button" name="acctinfo" value="Enter Bill" onClick="ypid(acctid.value,building.value,utility.value)"></span>
					</font>
				</td>
			</tr>
           <tr>    
              
            <td width="25%"><font face="Arial, Helvetica, sans-serif"> 
              <input type="hidden" name="bldg">
              </font></td>
		      
            <td width="75%"> <font face="Arial, Helvetica, sans-serif"> </font> 
            </td>
            </tr>
           
			</table>
	</form>		
</table>
<%
set cnn1=nothing%>
<IFRAME name="entry" width="100%" height="385" src="null.htm" scrolling="no" marginwidth="0" marginheight="0" ></IFRAME></body>
</html>