<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="../../gEnergy2_Intranet/styles.css" type="text/css">
</head>

<body bgcolor="#eeeeee" text="#000000">
<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%

ID= Request.Querystring("id")
POID = Request.Querystring("poid")

if not Request.Querystring("submitted") and not Request.Querystring("accepted") then
if isempty(id) then
%>
<form name="form2" method="post" action="savepoitem.asp">

<table border=0 cellpadding="3" cellspacing="1">
<tr bgcolor="#dddddd" style="font-weight:bold;">
  <td width="8%">Qty</td>
  <td width="15%">Unit</td>
  <td width="15%">Item #</td>
  <td width="15%">Unit  Price</td>
  <td>Description/<wbr>Comments</td>
</tr>
  
<tr> 
  <td><input type="text" name="qty" size="4"></td>
  <td><input type="text" name="unit"></td>
  <td><input type="text" name="invnum"></td>
  <td>$&nbsp;<input type="text" name="unitprice" size="16"></td>
  <td>
  <input type="text" name="description" size="30">
  <input type="submit" name="choice2"  value="Save">
  <input type="hidden" name="poid" value="<%=POID%>">
  </td>
</tr>  
</table>
</form>
<%
else

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")


sqlstr = "select * from po_item where id="&id

rst1.Open sqlstr, cnn1, 0, 1, 1

  if not rst1.EOF then
  %>
  <form name="form1" method="post" action="poitemupdate.asp">
  <table border=0 cellpadding="3" cellspacing="1" width="100%" style="border:1px solid #cccccc;">
  <tr bgcolor="#dddddd" style="font-weight:bold;">
    <td width="8%">Qty</td>
    <td width="15%">Unit</td>
    <td width="15%">Item #</td>
    <td width="15%">Unit Price</td>
    <td>Description/<wbr>Comments</td>
  </tr>
    <tr> 
    <td><input type="text" name="qty" value="<%=rst1("qty")%>" size="4"></td>
    <td><input type="text" name="unit" value="<%=rst1("unit")%>"></td>
    <td><input type="text" name="invnum" value="<%=rst1("invnum")%>" ></td>
    <td>$<input type="text" name="unitprice" value="<%=rst1("unitprice")%>" size="16"></td>
    <td>
    <input type="text" name="description" value="<%=rst1("description")%>">
    <input type="submit" name="choice"  value="Update">
    <input type="hidden" name="key" value="<%=rst1("id")%>">
    <input type="hidden" name="poid" value="<%=POID%>">
    </td>
  </tr>  
  </table>
  </form>      
  <%
  end if
end if
end if

%>
</body>
</html>
