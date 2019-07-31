<%option explicit%>
<!--#include file="adovbs.inc"-->
<%
'http params
dim crdate,c,total,j,strsql,tot1,sql,atable,prefix


if request("ji")="" then 
	j = request ("jg")
	aTable = "GY_MASTER_JOB"
	prefix = "jg"
else
	j = request ("ji")
	aTable = "IL_MASTER_JOB"
	prefix = "ji"
end if

c = request("c")


'adodb vars
dim cnn, cmd, rs, prm
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")


' open connection
cnn.Open "driver={SQL Server};server=10.0.7.20;uid=sa;pwd=!general!;database=main;"
cnn.CursorLocation = adUseClient

rs.open "SELECT crdate FROM sysobjects WHERE name = 'GY_master_po'",cnn
crdate=rs(0)
rs.close

' specify stored procedure to run based on company
if c="IL" then
strsql = "Select * from ilite.dbo.times where jobno='" & right(j,4) & "' order by date desc"
else
strsql = "Select * from times where jobno='" & right(j,4) & "' order by date desc"
end if

rs.open strsql,cnn

%>

<html>
<head>

<script language="JavaScript1.2">
function po_item(c,p,j) {
	theURL="https://appserver1.genergy.com/um/war/po/po_item.asp?c="+c+"&p="+p+"&j=" +j
	openwin(theURL,800,400)
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, scroll=on, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}
</script>
<title>Genergy War Room - Time sheet Detail</title>
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

<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#FFFFFF">

<tr style="font-family:arial; font-size:12; color:black" bgcolor="#0099FF">
    
  <td style="font-size:10" rowspan="2" align="center">
  

    <table width="99%" border="1">
      <tr bgcolor="#99FFCC"> 
        <td width="9%">Date</td>
        <td width="50%"> 
          <div align="center">Description</div>
        </td>
        <td width="11%"> 
          <div align="center">Hours</div>
        </td>
        <td width="11%"> 
          <div align="center">Over Time</div>
        </td>
        <td width="19%"> 
          <div align="center">User</div>
        </td>
      </tr>
    </table>
	 
    <table width="99%" border="0">
      <%
total = 0

while not rs.EOF 
total = total + cdbl(rs("hours"))
tot1 = tot1 + cdbl(rs("overt"))

%>
      
    </table>

    <table width="99%" border="0">
      <tr> 
        <td width="9%" height="20"> <font size="2"><%=rs("date")%> </font></td>
        <td width="50%" height="20"> <font size="2"><%=left(rs("description"),55)%> 
          </font></td>
        <td align="right" width="11%" height="20"> 
          <div align="right"> <font size="2"><%=formatnumber(rs("hours"),2)%> 
            </font></div>
        </td>
        <td align="right" width="9%" height="20"><font size="2"><%=formatnumber(rs("overt"),2)%></font></td>
        <td align="right" width="2%" height="20"><font size="2">&nbsp;</font></td>
        <td align="left" width="19%" height="20"><font size="2"><%=rs("matricola")%></font></td>
        <%
 rs.movenext


wend
%>
    </table>
      
        
	<p>&nbsp;</p>
    <table width="99%" border="0" dwcopytype="CopyTableRow">
      <tr> 
        <td height="23" width="16%"> 
          <div align="left"><font size="2">update as of now !</font></div>
        </td>
        <td width="43%"> 
          <div align="center"><b><%=j%></b></div>
        </td>
        <td height="23" width="11%"> 
          <div align="right"><b><%=formatnumber(total,2)%></b></div>
        </td>
        <td height="23" width="11%">
<div align="right"><b><%=formatnumber(tot1,2)%></b></div></td>
        <td height="23" width="19%">&nbsp;</td>
      </tr>
    </table>
	
	
      <tr> 
        <td>
          <div align="right"></div>
  </td>
      </tr> 
	  <%
	  set cnn = nothing %>
	  
</body>
</html>
