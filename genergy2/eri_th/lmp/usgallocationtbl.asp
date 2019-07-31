<%
Dim varpercent, usgunit,totalBuildingUsg,totalBuildingBill,tenantname,datacolors(14)

if by = "" then by = 0
if bp = "" then bp = 0

varpercent 	= ".00"

Dim tempcolor, tempR, tempG, tempB, count

        tempR = 0.3
        tempG = 0.61
        tempB = 0.79

        For count = 0 To 14 'static color generation
            tempR = (1999 / tempR) Mod 255
            tempG = (1777 / tempG) Mod 255
            tempB = (2003 / tempB) Mod 255
            datacolors(count) = RGB(CInt(tempR), CInt(tempG), CInt(tempB))
            tempR = Abs(Sin(tempR))
            tempG = Abs(Cos(tempG))
            tempB = Abs(Sin(tempB))
        Next 


Dim sqlstr, cmd3

'set cnn 	= server.createobject("ADODB.Connection")
set cmd3 	= server.createobject("ADODB.Command")
'set rs 		= server.createobject("ADODB.Recordset")

'cnn.Open getLocalConnect(bldgnum)
cmd3.CommandType 	= adcmdStoredProc
cnn.CursorLocation 	= adUseClient
Set cmd3.ActiveConnection = cnn

Set cmd3.ActiveConnection = cnn
    cmd3.CommandText = "sp_tenant_usage_allocation"
    ' set parameter type and append for tenant contribution pie
    Set prm = cmd3.CreateParameter("building", adVarChar, adParamInput, 5)
    cmd3.Parameters.Append prm
    Set prm = cmd3.CreateParameter("by", adinteger, adParamInput)
    cmd3.Parameters.Append prm
    Set prm = cmd3.CreateParameter("bp", adinteger, adParamInput)
    cmd3.Parameters.Append prm
    Set prm = cmd3.CreateParameter("percent", advarchar, adParamInput,10)
    cmd3.Parameters.Append prm
    Set prm = cmd3.CreateParameter("buildingusg", adinteger, adParamOutput)
    cmd3.Parameters.Append prm
    Set prm = cmd3.CreateParameter("buildingamt", adinteger, adParamOutput)
    cmd3.Parameters.Append prm
	
x=30
do until x <= 15 	
	cmd3.Parameters("building") 	= bldgnum
    cmd3.Parameters("by") 		= by
    cmd3.Parameters("bp") 		= bp
    cmd3.Parameters("percent") 	= 0
	'response.write "exec sp_tenant_usage_allocation '"&cmd3.Parameters("building")&"', "&cmd3.Parameters("by")&", "&cmd3.Parameters("bp")&", "&cmd3.Parameters("percent")
	'response.end
    set rs = cmd3.execute
	totalBuildingUsg 			= cmd3.Parameters("buildingusg")
	totalBuildingBill			= cmd3.Parameters("buildingamt")
	usgunit 					= "KWH"
	if rs.recordcount > 15 then 
		x = rs.recordcount
		varpercent = varpercent + .01
		rs.close
	else 
		x = rs.recordcount
	end if 	
loop
%>

<html>
<head>
<title></title>
</head><style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>
<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000">
<% if totalBuildingUsg <> "" and totalbuildingbill <> "" then %>
<table width="650" border="0" cellspacing="3" cellpadding="3" align="center">
  <tr>
    <td height="22"><font face="Arial, Helvetica, sans-serif" size="1">This buildings total usage for Period <%=bp%>, <%=by%> was <%=formatnumber(totalBuildingUsg)%>&nbsp;&nbsp;<%=usgunit%> and had a total electric bill of <%=formatcurrency(totalbuildingBill,2)%>. <br>Below is the tenant contribution analysis of usage and costs for this period.</font></td>
</tr></table>
<table width="650" height="1" cellspacing="0" cellpadding="0" align="center">
  <tr bgcolor="#0099FF"> 
    <td width="21%" style="border:1px solid black" align="center" colspan=2><b><font size="1" face="Arial">Tenant</font></b></td>
    <td width="20%" style="border:1px solid black" align="center"><b><font face="Arial" size="1">Usage</font></b></td>
    <td width="20%" style="border:1px solid black" align="center"><b><font size="1" face="Arial">Usage<br>Percentage</font></b></td>
    <td width="20%" style="border:1px solid black" align="center"><b><font face="Arial" size="1">Total 
      Cost</font></b></td>
    <td width="20%" style="border:1px solid black" align="center"><b><font size="1" face="Arial">Total 
      Cost<br>
      Percentage</font></b></td>
  </tr>
</table>
<div style="overflow:auto;width:675">
<table cellpadding="0" cellspacing="0" border="0" width="650" align="center">
<%
if rs.state = adstateopen then 
	x = 0
    Do Until rs.EOF
	if trim(cdbl(rs("totalamt"))) > 0 then 
	if lcase(rs("tenantname")) = "small tenants" then tenantname = "Tenants under " & Formatpercent(varpercent,0) &" usage*" else tenantname = rs("tenantname") end if
	%>
      <tr style="cursor:hand">
	   <td width="1%" bgcolor="<%=getColor(datacolors(x))%>">&nbsp;&nbsp;&nbsp;&nbsp;</td>
       <td width="20%" style="border-bottom:1px solid black;"><b><font size="1" face="Arial"><%=tenantname%></font></b></td>
       <td width="20%" style="border-bottom:1px solid black;" align="right"><b><font size="1" face="Arial"><%=formatnumber(rs("used"),1)%>&nbsp;<%=usgunit%></font></b></td>
       <td width="20%" style="border-bottom:1px solid black;" align="right"><b><font size="1" face="Arial"><%=formatpercent(cdbl(rs("usgpercent")),1)%></font></b></td>
       <td width="20%" style="border-bottom:1px solid black;" align="right"><b><font size="1" face="Arial"><%=formatcurrency(rs("totalamt"),1)%></font></b></td>
       <td width="20%" style="border-bottom:1px solid black;" align="right"><b><font size="1" face="Arial"><%=formatpercent(cdbl(rs("amtpercent")),1)%></font></b></td>
       </tr>
       <% 
	 end if
	 rs.MoveNext
	 x = x + 1
    Loop
%></table>
  <%
end if
Set cnn = Nothing
%>
  <div align="left"><br>
    <font size="1" face="Arial, Helvetica, sans-serif"><strong>*Only accounts 
    for sub-metered tenants </strong></font> </div>
</div>
  
  <% 
 else
 %>
<div align="center"><font size="4" face="Arial, Helvetica, sans-serif"><strong>NO 
  DATA AVAILBLE FOR PERIOD <%=bp%>, <%=by%> </strong></font> </div>
  <%	
 end if %>

</body></html>
<%
Public Function getColor(color)
    Dim r, g, b, tempcolor

    tempcolor = color
    r = tempcolor And &HFF&
    g = (tempcolor And &HFF00&) \ &HFF&
    b = (tempcolor And &HFF0000) \ &HFFFF&
    getColor = Right("0" & Hex(r), 2) & Right("0" & Hex(g), 2) & Right("0" & Hex(b), 2)
End Function
%>