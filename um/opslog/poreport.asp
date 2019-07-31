<html>
<head>
<title>Genergy Purchase Order</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" onload="print();">
<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
po= Request.Querystring("id1")

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")
'sqlstr = "select po.*,ltrim(str(po.jobnum))+'.'+ltrim(str(po.ponum)) as pono,p.* ,(p.qty*p.unitprice) as ptotal,p.description as d, case when j.[entry id] > 6283 then left(j.[entry type],2)+'-00'+convert(varchar(4),j.[entry id]) else '00-00'+convert(varchar(4),j.[entry id]) end as  tjob from po join po_item as p on PO.id=p.poid join [job log] j on po.jobnum=j.[entry id] where PO.id='"&po&"' order by p.id"
sqlstr="select po.*, po.description as po_description, cast(po.tax as decimal(18,4)) as posttax,ltrim(str(po.jobnum))+'.'+ltrim(str(po.ponum)) as pono,p.* ,(p.qty*p.unitprice) as ptotal,p.description as d, job as  tjob from po join po_item as p on PO.id=p.poid join master_job j on po.jobnum=j.id where PO.id='"&po&"' order by p.id"
sqlstr1="select sum((p.qty*p.unitprice)) as psubtotal from po join po_item as p on PO.id=p.poid where PO.id='"&po&"'"

rst1.Open sqlstr, cnn1, 0, 1,1
rst2.Open sqlstr1, cnn1, 0, 1,1

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
tax = cdbl(rst2("psubtotal")) * cdbl(rst1("posttax"))
shipping=rst1("ship_amt")
po_description = rst1("po_description")

%>
<table border=0 cellpadding="3" cellspacing="0" width="550">
<tr>
  <td>
  <table border="0" cellpadding="3" cellspacing="0" width="100%">
  <tr>
    <td colspan="2" class="borderTRBL000"><b>Genergy Purchase Order Number <%=rst1("pono")%></b>
	<br>
	Purchase Order Description: <%=po_description%>

	</td>
  </tr>
  <tr><td colspan="2" height="8"></td></tr>
  <tr>
    <td>
    <table border=0 cellpadding="0" cellspacing="0">
    <tr>
      <td>
      353 West 48th Street<br>
      New York, NY 10036<br>
      Tel: (212) 664-7600<br>
      Fax: (212) 664-1549<br>
      </td>
    </tr>
    </table>
    </td>
  </tr>
  </table>
  </td>
</tr>
<tr><td height="8"></td></tr>
<tr>
  <td>
  <table border=0 cellpadding="3" cellspacing="0" class="borderTRBL000" width="100%">
  <tr>
    <td width="34%"><b>PO Date:</b>&nbsp;<%=rst1("podate")%></td>
    <td width="33%" class="borderLCCC"><b>PO Number:</b>&nbsp;<%=rst1("JObnum")%>.<%=rst1("ponum")%></td>
    <td width="33%" class="borderLCCC"><b>Requisitioner:</b>&nbsp;<%=rst1("requistioner")%></td>
  </tr>
  </table>  
  </td>
</tr>
<tr><td height="4"></td></tr>
<tr>
  <td>
  <table border=0 cellpadding="2" cellspacing="0">
  <tr>
    <td>Vendor Name:</td>
    <td><%=rst1("vendor")%> </td>
  </tr>
  <tr>
    <td>Job Number:</td>
    <td><%=rst1("tjob")%> </td>
  </tr>
  <tr>
    <td>Job Name:</td>
    <td><%=rst1("JobName")%> </td>
  </tr>
  <tr>
    <td>Job Address:</td>
    <td><%=rst1("JobAddr")%> </td>
  </tr>
  <tr> 
    <td>Ship Address:</td>
    <td><%=rst1("shipaddr")%> </td>
  </tr>
  </table>  
  </td>
</tr>
<tr><td height="8"></td></tr>
<tr>
  <td>
  <table border=0 cellpadding="3" cellspacing="0" width="100%" class="borderTRBL000">
  <tr style="font-weight:bold;">
    <td class="borderB000">Qty</td>
    <td class="borderB000">Unit</td>
    <td class="borderB000">Item #</td>
    <td class="borderB000">Description</td>
    <td class="borderB000" width="13%">Unit Price</td>
    <td class="borderB000" width="13%">Total</td>
  </tr>
  <% While not rst1.EOF %>
  <tr> 
    <td class="borderBCCC"><%=rst1("qty")%>&nbsp;</td>
    <td class="borderBLCCC"><%=left(rst1("unit"),20)%>&nbsp;</td>
    <td class="borderBLCCC"><%=rst1("invnum")%>&nbsp;</td>
    <td class="borderBLCCC"><%=left(rst1("d"),60)%>&nbsp;</td>
    <td class="borderBLCCC" align="right"><%=formatcurrency(rst1("unitprice"))%>&nbsp;</td>
    <td class="borderBLCCC" align="right"><%=formatcurrency(rst1("ptotal"))%>&nbsp;</td>
  </tr>
  <% 
		rst1.movenext
		Wend
		end if
		%>
  </table>
  </td>
</tr>
<tr>
  <td align="right">
  <table border=0 cellpadding="3" cellspacing="0" class="borderTRBL000" bgcolor="#eeeeee" width="26%">
  <tr>
    <td>Shipping:</td>
    <td align="right"><%=FormatCurrency(shipping)%></td>
  </tr>
  <tr> 
    <td>Tax:</td>
    <td align="right"><%=FormatCurrency(tax,2)%></td>
  </tr>
  <tr>
    <td><b>Total:</b></td>
    <td align="right"><b><%=FormatCurrency(pototal)%></b></td>
  </tr>
  </table>  
  </td>
</tr>
<tr>
  <td height="30">&nbsp;</td>
</tr>
<tr>
  <td>
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="10%">Called&nbsp;in&nbsp;order</td>
    <td width="35%" class="borderB000">&nbsp;</td>
    <td>&nbsp;</td>
    <td width="10%">Approved&nbsp;by</td>
    <td width="35%" class="borderB000">&nbsp;</td>
  </tr>
  </table>
  </td>
</tr>
<tr>
  <td height="30">&nbsp;</td>
</tr>
<tr>
  <td class="borderB000">
  <i>Please notify us immediately of any freight charges applicable to this order.</i>
  </td>
</tr>
<tr>
  <td>
  <span style="font-size:7pt;">Date: 
   <%
   Function MyCurrentDate()
         MyCurrentDate = CStr(FormatDateTime(Date, vbShortDate))
   End Function
   Response.write mycurrentdate
  %></span>
  </td>
</tr>
</table>	  
 
<%rst1.close
rst2.close%>


</body>
</html>
