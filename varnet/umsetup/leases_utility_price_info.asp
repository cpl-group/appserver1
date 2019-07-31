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

<body>

<%
bill = Request("billid")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.8;uid=genergy1;pwd=g1appg1;database=genergy1;"

strsql = "SELECT * FROM tblLeasesUtilityPrices where billingid = " & bill
rst1.Open strsql, cnn1, 0, 1, 1
if not rst1.eof then

Response.Write "<script>" & vbCrLf

Response.Write "parent.frames.meter.location = " & Chr(34) & _
                  "meter_info.asp" & _
                  "?lui=" & rst1("leaseutilityid") & _
                  Chr(34) & vbCrLf
                  
 Response.Write "</script>" & vbCrLf 





%>


<table border="1" width="100%">
  <tr>
    <td width="10%" bgcolor="#C0C0C0"><p align="center"><font size="2">Admin Fee</font></td>
    <td width="10%" bgcolor="#C0C0C0"><p align="center"><font size="2">Tenant Rate</font></td>
    <td width="10%" bgcolor="#C0C0C0"><p align="center"><font size="2">Add On Fee</font></td>
    <td width="10%" bgcolor="#C0C0C0"><p align="center"><font size="2">Mod. Rate</font></td>
    <td width="10%" bgcolor="#C0C0C0"><p align="center"><font size="2">Coincident</font></td>
    <td width="10%" bgcolor="#C0C0C0"><p align="center"><font size="2">Profile</font></td>
    <td width="10%" bgcolor="#C0C0C0"><p align="center"><font size="2">Graph</font></td>
  </tr>
  
 <%
 while not rst1.eof
 %>
  <tr>
    <td width="10%"><p align="left"><font size="2"><%=FormatPercent(rst1("adminfee"))%></font></td>
    <td width="10%"><p align="left"><font size="2"><%=rst1("ratetenant")%></font></td>
    <td width="10%"><p align="left"><font size="2"><%=rst1("addonfee")%></font></td>
    <td width="10%"><p align="left"><font size="2"><%=rst1("ratemodify")%></font></td>
    
    <% if rst1("coincident") then %>   
    <td width="10%" align="center"><p><font size="2"><input type="checkbox" name="C1" value="ON" checked></font></p></td>
    <% else %>
    <td width="10%" align="center"><p><font size="2"><input type="checkbox" name="C1" value="ON" ></font></p></td>
    <% end if %>   
    
    <% if rst1("loadprofile") then %>   
    <td width="10%" align="center"><p><font size="2"><input type="checkbox" name="C1" value="ON" checked></font></p></td>
    <% else %>
    <td width="10%" align="center"><p><font size="2"><input type="checkbox" name="C1" value="ON" ></font></p></td>
    <% end if %>
    
    <% if rst1("prtgraph") then %>   
    <td width="10%" align="center"><p><font size="2"><input type="checkbox" name="C1" value="ON" checked></font></p></td>
    <% else %>
    <td width="10%" align="center"><p><font size="2"><input type="checkbox" name="C1" value="ON" ></font></p></td>
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








