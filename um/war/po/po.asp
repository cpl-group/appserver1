<%option explicit%>
<!--#include file="adovbs.inc"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'http params
dim crdate,c,total,j,strsql

'd = request("d")
c = request("c")
if request("ji")="" then 
	j = request ("jg")
else
	j = request ("ji")
end if
'o = request ("o")


'adodb vars
dim cnn, cmd, rs, prm
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")


' open connection
cnn.Open getConnect(0,0,"Intranet")
cnn.CursorLocation = adUseClient

rs.open "SELECT crdate FROM sysobjects WHERE name = 'GY_master_po'",cnn
crdate=rs(0)
rs.close

' specify stored procedure to run based on company
if c="IL" then
strsql= "select *,purchase_order as po_num from IL_master_PO where job ='" & j & "'"
else
strsql= "select *,commitment as po_num from GY_master_PO where job ='" & j & "'"
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
<title>Genergy War Room - Purchase Order</title>
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
      <tr> 
        <td width="10%">PO number</td>
        <td width="12%"> 
          <div align="center">Date</div>
        </td>
        <td width="23%"> 
          <div align="center">PO Description</div>
        </td>
        <td width="12%"> 
          <div align="center">Status</div>
        </td>
        <td width="13%"> 
          <div align="center">Amount</div>
        </td>
        <td width="14%"> 
          <div align="center">Amount Invoiced</div>
        </td>
        <td width="16%">Amount Paid</td>
      </tr>
    </table>
	 
    <table width="99%" border="0">
      <%
total = 0

while not rs.EOF 
total = total + rs("amount")


%>
      
    </table>

    <table width="99%" border="0">
	<%if rs("closed") then %>
	<tr bgcolor="#FFFFFF"> 
	<% else
	%>
      <tr bgcolor="#33ffff"> 
	  <%end if%>
	  
	  
        <td width="10%" height="20"> <a href=<%="javascript:po_item('" & c & "','" & rs("po_num") & "','" & j & "')"%>><%=rs("po_num")%></a> </td>
        <td width="12%" height="20"> <%=rs("date")%> </td>
        <td align="right" width="23%" height="20"> 
          <div align="left"> <%=rs("description")%> </div>
        </td>
        <td align="right" width="12%" height="20"><%=rs("closed")%></td>
        <td align="right" width="13%" height="20"> <%=rs("amount")%></td>
        <td align="right" width="14%" height="20"><%=formatcurrency(rs("amount_invoiced"),2)%></td>
        <td align="right" width="16%" height="20"><%=formatcurrency(rs("amount_paid"),2)%></td>
        <%
 rs.movenext


wend
%>
    </table>
      
        
	<p>&nbsp;</p><table width="99%" border="0" dwcopytype="CopyTableRow">
      <tr> 
        <td height="23" width="14%"> 
          <div align="left"><font size="2">update as of   <%=formatdatetime(crdate,0)%></font></div>
        </td>
        <td height="23" width="10%">&nbsp;</td>
        <td height="23" width="46%"> 
          <div align="right"><b><%=formatcurrency(total,2)%></b></div>
        </td>
        <td height="23" width="30%"> 
          <div align="right"></div>
        </td>
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