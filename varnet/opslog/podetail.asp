<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<%

ID= Request.Querystring("id")
POID = Request.Querystring("poid")

if not Request.Querystring("submitted") and not Request.Querystring("accepted") then
if isempty(id) then
%>
<form name="form2" method="post" action="savepoitem.asp">

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
  
  <tr> 
    <td width=6%> 
      <input type="submit" name="choice2"  value="SAVE">
      <input type="hidden" name="poid" value="<%=POID%>">
    </td>
    <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
      <input type="text" name="qty" >
      </font></td>
    <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
      <input type="text" name="unit">
      </font></td>
    <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
      <input type="text" name="invnum" >
      </font></td>
      <td width="11%" height="19" > <font face="Arial, Helvetica, sans-serif"> $
<input type="text" name="unitprice">
      </font></td>
    <td width="18%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
      <input type="text" name="description">
      </font></td>
    
  </tr>
  
</table></form>
<%
else

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"


sqlstr = "select * from po_item where id="&id

rst1.Open sqlstr, cnn1, 0, 1, 1


%>
<form name="form1" method="post" action="poitemupdate.asp">

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
  
  <tr> 
    
    <td width=6%> 
      <input type="submit" name="choice"  value="UPDATE">
	  <input type="hidden" name="key" value="<%=rst1("id")%>">
	<input type="hidden" name="poid" value="<%=POID%>">
    </td>
    <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
      <input type="text" name="qty" value="<%=rst1("qty")%>" >
      </font></td>
    <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
      <input type="text" name="unit" value="<%=rst1("unit")%>">
      
      </font></td>
    <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
      <input type="text" name="invnum" value="<%=rst1("invnum")%>" >
      </font></td>
      <td width="11%" height="19" > <font face="Arial, Helvetica, sans-serif"> 
        $
<input type="text" name="unitprice" value="<%=rst1("unitprice")%>">
      </font></td>
    <td width="18%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
      <input type="text" name="description" value="<%=rst1("description")%>">
      </font></td>
    
  </tr>
  
</table>
 </form>      
 <%
end if
end if

%>        
</body>
</html>
