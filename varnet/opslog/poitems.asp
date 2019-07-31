<html>
<head>
<script>
function updateEntry(id,poid){
	parent.frames.podetail.location="podetail.asp?id="+id+"&poid="+poid
}
</script>

<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<%

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"



sqlstr = "select * from po_item where poid=" & Request.querystring("poid")

rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.eof then
%>

<table width="100%" border="0">
  <tr>
    <td>
      <div align="center"> 
        <p><font face="Arial, Helvetica, sans-serif"><i>No PO Items Found </i></font></p>
        <hr>
        <p><font face="Arial, Helvetica, sans-serif"><i>New PO Or Closed Job</i></font></p>
        </div>
    </td>
  </tr>
</table>
<%
else
%>
<table width="100%" border="0">
  <tr bgcolor="#CCCCCC"> 
    <td width="2%" height="2"> 
    <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Quanity</font></td>
    <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Unit</font></td>
    <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Item 
      #</font></td>
    <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Unit 
      Price</font></td>
    <td width="18%" height="10"><font face="Arial, Helvetica, sans-serif">Description 
      / Comments</font></td>
  </tr>
  <% While not rst1.EOF %>
  <form name="form1" method="post" action="">
    <tr> 
      <input type="hidden" name="key" value="<%=rst1("id")%>">
      <input type="hidden" name="poid" value="<%=Request.querystring("poid")%>">
      <td width=6%> 
        <%if not Request.querystring("submitted") and not Request.querystring("accepted") then  %>
        <input type="button" name="edit" value="edit" size="7" onClick="updateEntry(key.value,poid.value)">
        <%end if %>
      </td>
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <%=rst1("qty")%></font></td>
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <%=rst1("unit")%> </font></td>
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <%=rst1("invnum")%> </font></td>
      <td width="11%" height="19" > <font face="Arial, Helvetica, sans-serif"> 
        <%=FormatCurrency(rst1("unitprice"))%> </font></td>
      <td width="18%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <%=rst1("description")%> </font></td>
    </tr>
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
