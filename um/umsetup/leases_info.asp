<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
			Response.Redirect "http://www.genergyonline.com"
		end if		
%>
<html>

<head>
<!--#include file="../adovbs.inc" -->
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Tenant Selection</title>
</head>

<body bgcolor="#FFFFFF">
<%
bldg = Request("bldg")
tenant = request("ten")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.8;uid=genergy1;pwd=g1appg1;database=genergy1;"

strsql = "SELECT * FROM tblleases WHERE bldgnum ='" & bldg &"' AND tenantnum ='" & tenant & "'"
rst1.Open strsql, cnn1, 0, 1, 1



if not rst1.eof then

Set cmd = Server.CreateObject("ADODB.command")
Set rst2 = Server.CreateObject("ADODB.recordset")

cmd.ActiveConnection = cnn1
cmd.CommandText = "sp_lease_display"
cmd.CommandType = 4
Set prm1 = cmd.CreateParameter("billid", adChar, adParamInput, 10, rst1("billingid"))
cmd.Parameters.Append prm1

Set rst2 = cmd.Execute



tmpMoveFrame =  "parent.frames.meter.location = " & Chr(34) & _
                  "meter_info.asp?lui=" & rst2("leaseutilityid") & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf

Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf



%>
<table border="1" width="100%" bordercolor="#000000">
  <tr bgcolor="#66CCFF"> 
    <td width="50%" align="center"> 
      <p align="left"><font size="2" face="Arial, Helvetica, sans-serif">Billing 
        Name</font>
    </td>
    <td width="20%" align="center"> 
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif">Floor</font>
    </td>
    <td width="10%" align="center"> 
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif">Sqft</font>
    </td>
    <td width="10%" align="center"> 
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif">Tax 
        Exempt</font>
    </td>
    <td width="10%" align="center"> 
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif">Off 
        Line</font>
    </td>
  </tr>
  <%
 while not rst1.eof
 %>
  <tr> 
    <td width="50%">
      <p align="left"><font size="2"><a href=lease_display.asp?B=<%=rst1("billingid")%> target="_self"><%=rst1("billingname")%></a>
    </td>
    <td width="20%">
      <p align="center"><font size="2"><%=rst1("flr")%></font>
    </td>
    <td width="10%">
      <p align="center"><font size="2"><%=rst1("sqft")%></font>
    </td>
    <% if rst1("taxexempt") then %>
    <td width="10%" align="center" bgcolor="#00FF00">
      <p><font size="2">On</font></p>
    </td>
    <% else %>
    <td width="10%" align="center" bgcolor="#FF0000">
      <p><font size="2">Off</font></p>
    </td>
    <% end if %>
    <% if rst1("LeaseExpired") then %>
    <td width="10%" align="center" bgcolor="#FF0000">
      <p><font size="2">Off</font></p>
    </td>
    <% else %>
    <td width="10%" align="center" bgcolor="#00FF00">
      <p><font size="2">On</font></p>
    </td>
    <% end if %>
  </tr>
  <%
rst1.movenext
wend
rst1.close
set cnn1 = nothing
end if
%>
</table>
</form>
</body>

</html>




















