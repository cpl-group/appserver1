<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<%
po= Request.Querystring("id1")

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"
sqlstr = "select po.*,ltrim(str(po.jobnum))+'.'+ltrim(str(po.ponum)) as pono,p.* ,(p.qty*p.unitprice) as ptotal,p.description as d, case when j.[entry id] > 6283 then left(j.[entry type],2)+'-00'+convert(varchar(4),j.[entry id]) else '00-00'+convert(varchar(4),j.[entry id]) end as  tjob from po join po_item as p on PO.id=p.poid join [job log] j on po.jobnum=j.[entry id] where PO.id='"&po&"'"

'response.write sqlstr
rst1.Open sqlstr, cnn1, 0, 1, 1


if rst1.EOF then
%>
<table width="100%" border="0">
  <tr>
    <td>ERROR</td>
  </tr>
</table>
<%
else
pototal=rst1("po_total")
shipping=rst1("ship_amt")
%>
<table width="100%" border="0">
  <tr>
    <td bgcolor="#3399CC" height="45"> <font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF" size="4">Genergy 
      Purchase Order Number </font><font face="Arial, Helvetica, sans-serif"><b><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="4"><%=rst1("pono")%></font></b></font></b></font></b></font></td>
  </tr></table>
	  
<table width=370 border="0">
  <tr> 
    <td width="68%"  ><font face="Arial, Helvetica, sans-serif"><b><font size="4">353 
      West 48th Street</font></b></font></td>
  </tr>
  <tr>
    <td width="68%"  ><b><font face="Arial, Helvetica, sans-serif" size="4">New 
      York, NY 10036</font></b></td>
  </tr>
  <tr> 
    <td width="68%"  > 
      <p><b><font face="Arial, Helvetica, sans-serif" size="4"> Tel:(212)664-7600</font></b></p>
    </td>
  </tr>
</table>
  
<table width="100%" border="0">
  <tr> 
    <td height="29" width="18%"> 
      <div align="center"><b><font face="Arial, Helvetica, sans-serif">Job Name</font></b></div>
    </td>
    <td height="29" width="23%"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font face="Arial, Helvetica, sans-serif">Vendor 
        Name</font></b></font></div>
    </td>
    <td height="29" width="25%"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b>Job Address</b></font></div>
    </td>
    <td height="29" width="34%"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b>Ship Address</b></font></div>
    </td>
  </tr>
  <tr> 
    <td width="18%"> 
      <p align="center"><font face="Arial, Helvetica, sans-serif"><%=rst1("JobName")%></font></p>
    </td>
    <td width="23%"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><%=rst1("vendor")%></font></div>
    </td>
    <td width="25%"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><%=rst1("JobAddr")%></font></div>
    </td>
    <td width="34%"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><%=rst1("shipaddr")%></font></div>
    </td>
  </tr>
  <br>
</table>
<hr>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td> 
      <div align="center"><b><font face="Arial, Helvetica, sans-serif">PO Date</font></b></div>
    </td>
    <td> 
      <div align="center"><b><font face="Arial, Helvetica, sans-serif">Requistioner</font></b></div>
    </td>
    <td> 
      <div align="center"><b><font face="Arial, Helvetica, sans-serif">Job Number</font></b></div>
    </td>
    <td> 
      <div align="center"><b><font face="Arial, Helvetica, sans-serif">PO Number</font></b></div>
    </td>
   
  </tr>
  <tr> 
    <font face="Arial, Helvetica, sans-serif">
    <td><div align="center"><%=rst1("podate")%>
      </div>
    </td>
    <td><div align="center"><%=rst1("requistioner")%></div></td>
    <td><div align="center"><%=rst1("tjob")%></div></td>
    <td><div align="center"><%=rst1("JObnum")%>.<%=rst1("ponum")%></div></td>
    </font>
  </tr>
</table>
<hr>

<table width="100%" border="0" cellspacing="0" cellpadding="0" height="61">
  <tr> 
    <td width="9%" bgcolor="#CCCCCC" height="28"><b><font face="Arial, Helvetica, sans-serif">Quantitiy</font></b> 
    </td>
    <td width="14%" bgcolor="#CCCCCC" height="28"><b><font face="Arial, Helvetica, sans-serif">Unit</font></b></td>
	<td width="14%" bgcolor="#CCCCCC" height="28"><b><font face="Arial, Helvetica, sans-serif">Item #</font></b></td>
    <td width="49%" bgcolor="#CCCCCC" height="28"><b><font face="Arial, Helvetica, sans-serif">Description</font></b> 
    </td>
    <td width="10%" bgcolor="#CCCCCC" height="28"><b><font face="Arial, Helvetica, sans-serif">Unit 
      Price</font></b> </td>
    <td width="11%" bgcolor="#CCCCCC" height="28"><b><font face="Arial, Helvetica, sans-serif">Total</font></b> 
    </td>
  </tr>
  <% While not rst1.EOF %>
  <tr> 
    <td width="9%"> <font face="Arial, Helvetica, sans-serif"><%=rst1("qty")%></font></td>
    <td width="14%"> 
      <div align="left"><font face="Arial, Helvetica, sans-serif"><%=rst1("unit")%></font></div>
    </td>
	<td width="14%"> 
      <div align="left"><font face="Arial, Helvetica, sans-serif"><%=rst1("invnum")%></font></div>
    </td>
    <td width="49%"> 
      <div align="left"><font face="Arial, Helvetica, sans-serif"><%=rst1("d")%></font></div>
    </td>
    <td width="10%"> 
      <div align="left"><font face="Arial, Helvetica, sans-serif"><%=formatcurrency(rst1("unitprice"))%></font></div>
    </td>
    <td width="11%"> 
      <div align="left"><font face="Arial, Helvetica, sans-serif"><%=formatcurrency(rst1("ptotal"))%></font></div>
    </td>
  </tr>
  <% 
		rst1.movenext
		Wend
		end if
		%>
</table>
<p></p>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="88%">
      <div align="right"><font face="Arial, Helvetica, sans-serif"><b>Shipping:</b></font> 
      </div>
    </td>
    <td width="12%"><%=FormatCurrency(shipping)%></td>
  </tr>
  <tr> 
    <td width="88%"> 
      <div align="right"><font face="Arial, Helvetica, sans-serif"><b>Total: </b></font></div>
    </td>
    <td width="12%"><%=FormatCurrency(pototal)%></td>
  </tr>
</table>
<p></p>
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="27">
  <tr> 
    <td width="37%"><font face="Arial, Helvetica, sans-serif">Called in Order_______________________________________</font></td>
    <td width="63%"><font face="Arial, Helvetica, sans-serif">Approved By___________________________________________</font></td>
  </tr>
</table>
<br>
<i><font face="Arial, Helvetica, sans-serif">Please notify us immediately if there are any freight charges concerning this order.</font></i>
<hr>
<i><font face="Arial, Helvetica, sans-serif">Date: </font></i>
 <%
 Function MyCurrentDate()
    	 MyCurrentDate = CStr(FormatDateTime(Date, vbShortDate))
 End Function
 Response.write mycurrentdate
%>



</body>
</html>
