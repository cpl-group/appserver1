<html>
<head>
<script language="JavaScript" type="text/javascript">
//<!--
function updateEntry(id,poid){
  parent.frames.podetail.location="podetail.asp?id="+id+"&poid="+poid
}

function highlight(tRow){
  tRow.style.backgroundColor = "lightgreen";
}

function unlight(tRow){
  tRow.style.backgroundColor = "white";
  
}
//-->
</script>

<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="../../gEnergy2_Intranet/styles.css" type="text/css">
</head>

<body bgcolor="#eeeeee" text="#000000">
<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")



sqlstr = "select * from po_item where poid='" & Request.querystring("poid") & "'"

rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.eof then
%>

<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr>
  <td>
  <p>No PO items found</p>
  <hr size="1" noshade>
  <p>New PO or closed job</p>
  </td>
  </tr>
</table>
<%
else
%>
<table border=0 cellpadding="3" cellspacing="1" width="100%">
<tr bgcolor="#dddddd" style="font-weight:bold;">
  <td>Quantity</td>
  <td>Unit</td>
  <td>Item #</td>
  <td>Unit  Price</td>
  <td>Description/<wbr>Comments</td>
</tr>
<% While not rst1.EOF %>
<form name="form1" method="post" action="">
<%if not Request.querystring("submitted") and not Request.querystring("accepted") then  %>
<tr onmouseover="highlight(this);" onmouseout="unlight(this);" onclick="updateEntry(key.value,poid.value)" bgcolor="#ffffff">
<% else %>
<tr>
<%end if %>
  <td><%=rst1("qty")%></td>
  <td><%=rst1("unit")%></td>
  <td><%=rst1("invnum")%></td>
  <td><%=FormatCurrency(rst1("unitprice"))%></td>
  <td><%=rst1("description")%></td>
</tr>
  <input type="hidden" name="key" value="<%=rst1("id")%>">
  <input type="hidden" name="poid" value="<%=Request.querystring("poid")%>">
</form>
<%
rst1.movenext
Wend
%>
</table>
<%
end if
rst1.close
%>
</body>
</html>
