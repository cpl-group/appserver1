<%option explicit%>
<!--#include file="adovbs.inc"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'http params
dim crdate,total,p,j,c

'd = request("d")
j = request("j")
p = request("p")
c = request ("c")


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
cmd.CommandText = "IL_PO_ITEM"
else
cmd.CommandText = "GY_PO_ITEM"
end if
cmd.CommandType = adCmdStoredProc

Set prm = cmd.CreateParameter("po", adchar, adParamInput,9)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("job", adchar, adParamInput,9)
cmd.Parameters.Append prm

' assign internal name to stored procedure
cmd.Name = "test"
Set cmd.ActiveConnection = cnn

'return set to recordset rs
cnn.test  p,j, rs


%>
<html>
<head>
<script language="JavaScript1.2">
function po_item(c,p,j) {
	theURL="https://appserver1.genergy.com/um/war/po/po_item.asp?c="+c+"&p="+p+"&j=" +j
	openwin(theURL,800,400)
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, scrollbars=yes, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}
</script>
<title>Genergy War room - PO Details</title>
</head>
<style type="text/css">
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
        <td width="10%"> 
          <div align="center">Description</div>
        </td>
        <td width="10%"> 
          <div align="center">Units</div>
        </td>
        <td width="10%"> 
          <div align="center">Unit_cost</div>
        </td>
        <td width="10%"> 
          <div align="center">Unit_description</div>
        </td>
        <td width="10%"> 
          <div align="center">Amount</div>
        </td>
        <td width="10%">Amount Invoiced</td>
        <td width="10%">Amount Paid</td>
      </tr>
    </table>
	 
    <table width="99%" border="0">
      <%
total = 0

while not rs.EOF 
total = total + rs("amount_paid")


%>
      
    </table>

    <table width="99%" border="0" height="8">
      <tr bgcolor="#FFFFFF"> 
        <td width="10%" height="20"> <%=rs("po_num")%> </td>
        <td width="10%" height="20"> <%=rs("po_Desc")%></td>
        <td align="right" width="10%" height="20"> 
          <div align="left"> <%=rs("Units")%> </div>
        </td>
        <td align="right" width="10%" height="20"><%=rs("Unit_cost")%></td>
        <td align="right" width="10%" height="20"> <%=rs("Unit_description")%></td>
        <td align="right" width="10%" height="20"><%=formatcurrency(rs("amount"),2)%></td>
        <td align="right" width="10%" height="20"><%=formatcurrency(rs("amount_invoiced"),2)%></td>
        <td align="right" width="10%" height="20"><%=formatcurrency(rs("amount_paid"),2)%></td>
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
          <div align="right"></div>
        </td>
        <td height="23" width="30%"> 
          <div align="right"><b><%=formatcurrency(total,2)%></b></div>
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