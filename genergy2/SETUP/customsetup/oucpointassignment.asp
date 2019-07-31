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

if 	not( _
	checkgroup("Genergy Users")<>0 _
	or checkgroup("clientOperations")<>0 _
	) then%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim meterid, id, action, premise, bms, pointtype, pointname, bldg
id = trim(request("id"))
bldg = trim(request("bldg"))
meterid = trim(request("meterid"))
action = trim(request("action"))
premise = trim(request("premise"))
bms = trim(request("bms"))
pointtype = trim(request("pointtype"))
pointname = trim(request("pointname"))

dim cnn1, rst1, sql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

if trim(action)<>"" then
	if trim(action)="Add Point" then
		sql = "INSERT INTO oucdatakey (premise, meterPName, type, meterid, fpointname) VALUES ('"&premise&"', '"&bms&"', '"&pointtype&"', "&meterid&", '"&pointname&"')"
	elseif trim(action)="Update Point" then
		sql =  "UPDATE oucdatakey SET premise='"&premise&"', meterPName='"&bms&"', type='"&pointtype&"', meterid="&meterid&", fpointname='"&pointname&"' WHERE id="&id
	elseif trim(action)="Delete Point" then
		sql =  "DELETE FROM oucdatakey WHERE id='"&id&"'"
	end if
  'Logging Update
  logger(sql)
  'end Log
	if sql<>"" then cnn1.execute sql
    premise = ""
    bms = ""
    pointtype = ""
    pointname = ""
    id = ""
end if

dim billingid, address
if trim(meterid)<>"" then
  rst1.open "SELECT * FROM meters m, buildings b , tblleasesutilityprices lup, tblleases l WHERE m.bldgnum=b.bldgnum and m.leaseutilityid=lup.leaseutilityid and lup.billingid=l.billingid and m.meterid="&meterid, cnn1
  if not rst1.eof then 
    address = rst1("strt") & " &gt; " & rst1("billingname") & " &gt; " & rst1("meternum")
    billingid = rst1("billingid")
  end if
  rst1.close
end if
%>
<html>
<head>
<title>Building View</title>
<link rel="Stylesheet" href="../setup.css" type="text/css">
</head>
<body bgcolor="#eeeeee" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<form name="form2" method="post" action="oucpointassignment.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#6699cc">
  <td><span class="standardheader"> Data Assignments | <span style="font-weight:normal;"><%=address%></span></td>
</tr>
</table>
<%
dim rowcolor
if trim(meterid)<>"" then
  rst1.open "SELECT o.id as oid, * FROM oucdatakey o, dbo.pointdefs p WHERE o.type=p.id and meterid="&meterid, getLocalConnect(bldg)
	if not rst1.EOF then%>
		<table width="100%" border="0" cellpadding="3" cellspacing="0">
		</table>
		<div style="height:165;overflow:auto;border-bottom:1px solid #cccccc;">
		<table width="100%" border="0" cellpadding="3" cellspacing="0">
		<tr bgcolor="#dddddd">
			<td><span class="standard"><b>Premise</b></span></td>
			<td><span class="standard"><b>BMS</b></span></td>
			<td><span class="standard"><b>datapoint</b></span></td>
			<td><span class="standard"><b>Type</b></span></td>
			<td>&nbsp;</td>
		</tr>
		<%do until rst1.EOF
        if id=trim(rst1("oid")) then
          rowcolor = "#ccffcc"
          premise = rst1("premise")
          bms = rst1("meterPName")
          pointname = rst1("fpointname")
          pointtype = rst1("type")
        else
          rowcolor = "white"
        end if
        %>
    		<tr bgcolor="<%=rowcolor%>" onmouseover="this.style.backgroundColor='lightgreen'" onmouseout="this.style.backgroundColor='<%=rowcolor%>'" onclick="document.location='oucpointassignment.asp?meterid=<%=meterid%>&id=<%=rst1("oid")%>&bldg=<%=bldg%>';">
    			<td width="5%"><span class="standard"><%=rst1("premise")%>&nbsp;</span></td>
    			<td width="5%"><span class="standard"><%=rst1("meterPName")%>&nbsp;</span></td>
    			<td width="5%"><span class="standard"><%=rst1("fpointname")%>&nbsp;</span></td>
    			<td width="15%"><span class="standard"><nobr><%=rst1("pointdesc")%> (<%=rst1("dpointname")%>)</nobr></span></td>
    			<td width="70%">&nbsp;</td>
    		</tr>
    		<%
        rst1.movenext
		loop
    %>
		</table>
		</div>
	</table>
	<%end if
  rst1.close
end if%>
<table border=0 cellpadding="3" cellspacing="0" width="100%" style="border-top:1px solid #ffffff;">
<tr>
  <td>
  <table border=0 cellpadding="3" cellspacing="0">
  <tr><td align="right">BMS&nbsp;</td>
      <td>
      <select name="bms">
      <%if trim(pointtype)="" then%><option value="">Select BMS</option><%end if%>
      <%
        rst1.open "SELECT c.premiseid, substring(BMS,CHARINDEX('.',BMS)+1, len(bms)) as BMS FROM custom_oucAccount c, PremiseAssoc p WHERE c.id=p.CIS and billingid='"&billingid&"' ORDER BY bms", cnn1
        if trim(premise)="" then if not rst1.eof then premise = rst1("premiseid")
        do until rst1.EOF
        %>
          <Option value="<%=rst1("BMS")%>"<%if trim(rst1("BMS"))=trim(bms) then response.write " SELECTED"%>><%=rst1("bms")%></Option>
        <%
        rst1.movenext
        loop
        rst1.close
      %>
      </select>
      </td></tr>
  <tr><td align="right">Point&nbsp;Name&nbsp;</td>
      <td><input type="text" name="pointname" size="10" value="<%=pointname%>"></td></tr>
  <tr><td align="right">Type&nbsp;</td>
      <td><select name="pointtype">
      <%if trim(pointtype)="" then%><option value="">Select Point Definition</option><%end if%>
      <%rst1.open "SELECT * FROM pointdefs", getConnect(0,bldg,"billing")
      do until rst1.EOF
      %>
        <Option value="<%=rst1("id")%>"<%if trim(rst1("id"))=trim(pointtype) then response.write " SELECTED"%>><%=rst1("pointdesc")%> (<%=rst1("dpointname")%>)</Option>
      <%
      rst1.movenext
      loop
      rst1.close
      %>
      </select></td></tr>
  <tr>
    <td>&nbsp;</td>
    <td>
	<%if not(isBuildingOff(bldg)) then%>
    <%if id<>"" then%>
    <input type="submit" name="action" value="Update Point" style="border:1px outset #ddffdd;background-color:ccf3cc;">
    <input type="submit" name="action" value="Delete Point" style="border:1px outset #ddffdd;background-color:ccf3cc;">
    <input type="button" value="Cancel" style="border:1px outset #ddffdd;background-color:ccf3cc;" onclick="document.location='oucpointassignment.asp?meterid=<%=meterid%>';">
    <%else%>
    <input type="submit" name="action" value="Add Point" style="border:1px outset #ddffdd;background-color:ccf3cc;">
    <%end if%>
    <%end if%>
    <input type="hidden" name="meterid" value="<%=meterid%>">
    
    <input type="hidden" name="billingid" value="<%=billingid%>">
    <input type="hidden" name="premise" value="<%=premise%>">
    <input type="hidden" name="bldg" value="<%=bldg%>">
    <input type="hidden" name="id" value="<%=id%>">
    </td>
  </tr>
  </table>
  </td>
</tr>
</table>

</form>
</body>
</html>