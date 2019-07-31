<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
sub closewindow()
	%>
	<script>
		window.close();
	</script>
	<%
	response.end
end sub

if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, tid, customsrc, action, premise, cisid, pointname, bmsid
pid = request("pid")
bldg = request("bldg")
tid = request("tid")
premise = request("premise")
cisid = request("cisid")
bmsid = request("bmsid")
pointname = request("pointname")
action = request("action")
customsrc = request("customsrc")

dim cnn1, rst1, sql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

if trim(premise)<>"" then
	if trim(action)="Save" then
		sql = "INSERT INTO "&customsrc&" (billingid, premiseid) VALUES ('"&tid&"', '"&premise&"')"
	elseif trim(action)="Update" then
		sql =  "UPDATE "&customsrc&" SET premiseid='"&premise&"' WHERE billingid='"&tid&"'"
	elseif trim(action)="Add BMS" then
		sql =  "INSERT INTO premiseAssoc (CIS,BMS) VALUES ('"&cisid&"', '"&pointname&"')"
	elseif trim(action)="Delete BMS" then
		sql =  "DELETE FROM premiseAssoc WHERE id='"&bmsid&"'"
	elseif trim(action)="Update BMS" then
		sql =  "UPDATE premiseAssoc SET BMS='"&pointname&"' WHERE id='"&bmsid&"'"
	end if

  'Logging Update
  logger(sql)
  'end Log
	if sql<>"" then cnn1.execute sql

end if
'if trim(action)<>"" then closewindow()

dim bldgname, address
if trim(bldg)<>"" then
	rst1.Open "SELECT bldgname, strt, customsrc FROM buildings WHERE bldgnum='"&bldg&"'", cnn1
	if not rst1.EOF then
		bldgname = rst1("bldgname")
		address = rst1("strt")
		customsrc = "custom_oucAccount"
	end if
	rst1.close
  rst1.open "SELECT premiseid, id FROM custom_oucAccount WHERE billingid='"&tid&"'", cnn1
	if not rst1.eof then 
    premise = rst1("premiseid")
    cisid = rst1("id")
  end if
	rst1.close
end if
%>
<html>
<head>
<title>Building View</title>
<link rel="Stylesheet" href="../setup.css" type="text/css">
<script>
function openbmspoint(idbms, bms)
{	frm = document.form2;
	if(idbms!="")
	{	frm.bmsid.value=idbms;
		frm.pointname.value=bms;
		document.all['editbms'].style.display='inline'
	}
}

function clearrows(currentrow)
{	var nodes = document.getElementsByName('bmsrow');
	for(i=0;i<nodes.length;i++)
	{	nodes[i].basecolor='white';
		nodes[i].style.backgroundColor = 'white'
	}
	currentrow.basecolor='lightgreen';
	currentrow.style.backgroundColor = 'lightgreen'
}
</script>
</head>
<body bgcolor="#ffffff" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form name="form2" method="post" action="oucbuildingpremise.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#3399cc"><td>
	<table border=0 cellpadding="0" cellspacing="0" width="100%">
	<tr><td><span class="standardheader">
		<%if trim(cisid)<>"" then%>
			Update Premise | <span style="font-weight:normal;"><a href="portfolioedit.asp?pid=<%=pid%>" style="color:#ffffff;"><%'=portfolioname%></a> &gt; <%=bldgname%></span>
		<%else%>
			Add New Premise | <span style="font-weight:normal;"><a href="portfolioedit.asp?pid=<%=pid%>" style="color:#ffffff;"><%'=portfolioname%></a> 
		<%end if%></span></td></tr>
	</table></td>
</tr>
<tr bgcolor="#eeeeee">
	<td style="border-bottom:1px solid #999999">
		<table border=0 cellpadding="3" cellspacing="0"><tr>
    <td align="right"><span class="standard">Building Name</span></td> 
    <td><span class="standard"><%=bldgname%></span></td>
    <td width="30">&nbsp;</td>
    <td  align="right"><span class="standard">Building&nbsp;#</span></td>
    <td><span class="standard"><%=bldg%></span></td>  
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">OUC Building Premise</span></td>
    <td><input type="text" name="premise" value="<%=premise%>"></td>
    <td width="30">&nbsp;</td>
    <td align="right"></td>
    <td>
    </td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right">
	<%if not(isBuildingOff(bldg)) then%>
      <%if trim(premise)<>"" then%>
        <input type="submit" name="action" value="Update" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
      <%else%>
        <input type="submit" name="action" value="Save" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
      <%end if%>
    <%end if%>
	</td>
    <td></td>
    <td width="30">&nbsp;</td>
    <td align="right"></td>
    <td>
    </td>
  </tr></table>
	</td>
</tr>
</table>
<%
if trim(premise)<>"" then
	rst1.Open "SELECT * FROM premiseAssoc WHERE CIS='"&cisid&"' ORDER BY BMS", cnn1
	'response.write "SELECT * FROM premiseAssoc WHERE CIS='"&cisid&"' ORDER BY BMS"
	'response.end
	if not rst1.EOF then%>
		<table width="100%" border="0" cellpadding="3" cellspacing="0">
		<tr bgcolor="#dddddd">
			<td><span class="standard"><b>BMS Point Name</b></span></td>
		</tr>
		</table>
		<div style="height:70;overflow:auto">
		<table width="100%" border="0" cellpadding="3" cellspacing="0">
		<%do until rst1.EOF%>
		<tr bgcolor="#ffffff" id="bmsrow" basecolor="white" onmouseover="this.style.backgroundColor='lightgreen'" onmouseout="this.style.backgroundColor = this.basecolor" onclick="clearrows(this);openbmspoint('<%=rst1("id")%>', '<%=rst1("bms")%>');">
			<td><span class="standard"><%=rst1("BMS")%>&nbsp;</span></td>
		</tr>
		<%rst1.movenext
		loop%>
		</table>
		</div>
	</table>
	<%end if
end if%>

&nbsp;BMS&nbsp;<input type="text" name="pointname">
<div id="editbms" style="display:none">
<%if not(isBuildingOff(bldg)) then%>
<input type="submit" name="action" value="Update BMS" style="border:1px outset #ddffdd;background-color:ccf3cc;">
<input type="submit" name="action" value="Delete BMS" style="border:1px outset #ddffdd;background-color:ccf3cc;">
<%end if%>
</div>
<%if not(isBuildingOff(bldg)) then%>
<input type="submit" name="action" value="Add BMS" style="border:1px outset #ddffdd;background-color:ccf3cc;">
<%end if%>


<input type="hidden" name="cisid" value="<%=cisid%>">
<input type="hidden" name="bmsid">
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="tid" value="<%=tid%>">
<input type="hidden" name="customsrc" value="<%=customsrc%>">
</form>
</body>
</html>