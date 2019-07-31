<html>

<head>
<!--#include file="../adovbs.inc" -->
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Tenant Selection</title>
</head>

<body bgcolor="#FFFFFF">
<%
lui = Request("lui")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.8;uid=genergy1;pwd=g1appg1;database=genergy1;"

strsql = "SELECT * FROM meters where leaseutilityid = " & lui 
rst1.Open strsql, cnn1, 0, 1, 1

if not rst1.eof then
%>
<table border="1" width="100%" bordercolor="#000000">
  <tr bgcolor="#66CCFF"> 
    <td width="10%"> 
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif">Meter</font> 
    </td>
    <td width="10%"> 
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif">Date 
        Off-line</font> 
    </td>
    <td width="10%"> 
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif">Date 
        Last read</font> 
    </td>
    <td width="10%"> 
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif">Time 
        Read</font> 
    </td>
    <td width="10%"> 
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif">KWH 
        X</font> 
    </td>
    <td width="10%"> 
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif">KW 
        X</font> 
    </td>
    <td width="10%"> 
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif">Location</font> 
    </td>
    <td width="10%"> 
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif">Riser</font> 
    </td>
    <td width="5%"> 
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif">On-Line</font> 
    </td>
    <td width="5%"> 
      <p align="center"><font size="2" face="Arial, Helvetica, sans-serif">Note</font> 
    </td>
  </tr>
  <%
 while not rst1.eof
 %>
  <tr> 
    <td width="10%"> 
      <p align="left"><a href=addmember.asp?M=<%=rst1("MeterID")%> target="_self"><font size="2"><%=rst1("MeterNum")%></font></a> 
    </td>
    <td width="10%"> 
      <p align="left"><font size="2"><%=rst1("dateoffline")%></font> 
    </td>
    <td width="10%"> 
      <p align="left"><font size="2"><%=rst1("datelastread")%></font> 
    </td>
    <td width="10%"> 
      <p align="left"><font size="2"><%=rst1("TimeLastRead")%></font> 
    </td>
    <td width="10%"> 
      <p align="left"><font size="2"><%=rst1("Manualmultiplier")%></font> 
    </td>
    <td width="10%"> 
      <p align="left"><font size="2"><%=rst1("demandmultiplier")%></font> 
    </td>
    <td width="10%"> 
      <p align="left"><font size="2"><%=rst1("location")%></font> 
    </td>
    <td width="10%"> 
      <p align="left"><font size="2"><%=rst1("riser")%></font> 
    </td>
    <% if rst1("online") then %>
    <td width="5%" align="center" bgcolor="#00FF00"> 
      <p><font size="2">On</font></p>
    </td>
    <% else %>
    <td width="5%" align="center" bgcolor="#FF0000"> 
      <p><font size="2" color="#FFFF00">Off</font></p>
    </td>
    <% end if %>
    <% if not isnull(rst1("metercomments")) then %>
    <td width="5%" align="center" bgcolor="#00FF00"> 
      <p>&nbsp;</p>
    </td>
    <% else %>
    <td width="5%" align="center"> 
      <p>&nbsp;</p>
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

</body>

</html>





















