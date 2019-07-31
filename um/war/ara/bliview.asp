<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc" -->
<%
'COMMENTS test
'1/22/2008 N.Ambo modified page to ensure that values reflected on the duplicate bill are infact values from the
'actual bill such as the billing address and amounts (sp_invoice_text also amended in accordance with this change)
'1/20/2008 Michelle T. Modify line 155 and 162 to match invoices in timberline with invoices in our system.  
  
dim cnn, c,cmd, rs,jid,invoiceid,prm1,prm2,invoicedate,invoicecustomer,nobilltext, nobill, bt_Name,bt_Company, bt_Address_1, bt_Address_2, bt_Address_3, bt_City, bt_State, bt_Zip, pdf,billtax, tot_invoice_amt

jid=trim(request("jid"))
c=trim(request("c"))
invoiceid=trim(request("invoiceid"))
if trim(request("pdf"))="yes" then pdf = true else pdf = false
  
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")

' open connection
cnn.Open getConnect(0,0,"Intranet")

' Set command properties
cmd.CommandType = adCmdStoredProc

select case c
case "GY"
cmd.CommandText = "sp_invoice_text"
case "IL"
cmd.CommandText = "sp_IL_invoice_text"
case "NY"
cmd.CommandText = "sp_NY_invoice_text"
case "GE"
cmd.CommandText = "sp_GE_invoice_text"
end select


' Specify connection
cmd.ActiveConnection = cnn

' Set up parameters
cmd.Parameters.Refresh
cmd(1)=jid
cmd(2)=invoiceid
'response.write cmd(1) & "<BR>" & cmd(2)
'response.end
'return set to recordset rs
set rs=cmd.execute
dim price,recidx
price=0
recidx=0
if not rs.eof then
'1/22/2008 N.Ambo - This section of variables modified because of change to Sp "sp_invoice_text"
 price=rs("billed")
  invoicedate=rs("InvoiceDate")
  invoicecustomer=trim(rs("customer"))
  bt_Name=trim(rs("contact_name"))
  bt_Address_1=trim(rs("address_1"))
  bt_Address_2=trim(rs("address_2")) 
  bt_Address_3=trim(rs("address_3"))
  'bt_city	= trim(rs("city"))
  'bt_state	= trim(rs("state"))
  'bt_zip	= trim(rs("zip_code")) 
  bt_Company= trim(rs("company")) 
  billtax = rs("tax")
  tot_Invoice_Amt = rs("invoice_Amount")
  'if cdbl(price) <> cdbl(tot_invoice_amt) then 
  		'billTax = tot_invoice_amt - price
 ' else
  		'billTax = 0 
		'tot_invoice_amt = price
  'end if 
else
  nobill = true
  nobilltext = "<tr><td colspan=""6"">Please Check if the Customer has a Billing Contact.<br><br><input type=""button"" value=""Back"" onClick=""history.back()"" style=""background-color:#eeeeee;border:1px outset #ffffff;color:336699;""><br></td></tr>"
end if

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Filed Invoice Report</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<body bgcolor="#FFFFFF">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr bgcolor="#6699cc"> 
    <td width="61%"><span class="standardheader">Invoice <%=invoiceid%>, Job <%=jid%></span></td>
    <td width="39%" align="right"><%if not pdf then%><input type=button value="Back" onClick="history.back()" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;"><%end if%></td>
	<td align="right"><%if not pdf then%><input type="button" value="Print PDF" onclick="window.open('http://pdfmaker.genergyonline.com/pdfmaker/pdfReport_v2.asp?devIP=<%=request.servervariables("server_name")%>&sn=<%=request.servervariables("script_name")%>&qs=<%=server.urlencode("jid="&jid&"&invoiceid="&invoiceid&"&invoicedate="&invoicedate&"&invoicecustomer="&invoicecustomer&"&c="&c)%>','','')" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;"><%end if%></td>
  </tr>
  <tr> 
    <td align=right><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="14%">SENT TO:</td>
          <td width="86%"><font size="2"><%=bt_Company%></font></td>
        </tr>
         <%if bt_Name <> "" then %>
        <tr>
          <td>&nbsp;</td>
          <td><font size="2"><%=bt_Name%></font></td>
        </tr>
         <%end if%>
        <tr>
          <td>&nbsp;</td>
          <td><font size="2"><%=bt_Address_1%></font></td>
        </tr>
		 <tr>
          <td>&nbsp;</td>
          <td><font size="2"><%=bt_Address_2%></font></td>
        </tr>	
		<%if bt_Address_3 <> "" then %>
        <tr>
          <td>&nbsp;</td>
          <td><font size="2"><%=bt_Address_3%></font></td>
        </tr>
        <%end if%>
      </table></td>
    <td align=right><table border=0 cellpadding="3" cellspacing="1" bgcolor="#cccccc">
        <tr bgcolor="dddddd"> 
          <td align="center">Invoice Date</td>
          <td align="center">Customer No.</td>
          <td align="center">Invoice No.</td>
          <td align="center">Job No.</td>
        </tr>
        <tr bgcolor="#ffffff"> 
          <td align="center"><%=invoicedate%></td>
          <td align="center"><%=invoicecustomer%></td>
          <td align="center"><%=invoiceid%></td>
          <td align="center">&nbsp;<%=jid%>&nbsp;</td>
        </tr>
      </table></td>
  </tr>
</table>

<table border=0 width=100% cellpadding="3" cellspacing="0" bgcolor="#ffffff">
  <tr bgcolor="dddddd" style="font-weight:bold;">
  <td width=10% valign=bottom style="border:1px solid #cccccc;">&nbsp;&nbsp;&nbsp;Item</td>
  <td valign=bottom width=45% style="border-top:1px solid #cccccc;border-bottom:1px solid #cccccc;">Description</td>
  <td width=10% valign=bottom style="border:1px solid #cccccc;">Units</td>
  <td width=10% style="border-top:1px solid #cccccc;border-bottom:1px solid #cccccc;">Unit of<br>Measure</td>
  <td width=10% valign=bottom style="border:1px solid #cccccc;">Unit Price</td>
  <td align=right width=15% valign=bottom style="border-top:1px solid #cccccc;border-bottom:1px solid #cccccc;border-right:1px solid #cccccc;">Amount&nbsp;&nbsp;&nbsp;</td>
  </tr>
<%
dim units
if nobill then 
  response.Write("<tr><td colspan=""6"">BLI table(s) not populated<br><br><input type=""button"" value=""Back"" onClick=""history.back()"" style=""background-color:#eeeeee;border:1px outset #ffffff;color:336699;""><br></td></tr>")
end if
while not rs.eof
  recidx=recidx+1
  'Added the following code to correct the units, on 6/23/2010, Xiufeng
  if rs("units")<>"0" then units=rs("units") else units=""
  
  ' Any long items will be wrapped anyway, so just print...
  if len(rs("description"))>100 then
    'response.Write("<tr bgcolor=""#ffffff""><td>&nbsp;</td><td>"&rs("description")&"&nbsp;</td><td colspan=""3"">&nbsp;</td>")
    '1/22/2008 N.Ambo replaced with line below
   ' response.Write("<tr bgcolor=""#ffffff""><td>&nbsp;</td><td>"&rs("description")&"&nbsp;</td><td>"&replace(rs("units"),"0","&nbsp;")&"</td><td>"&replace(rs("unit_of_measure")," ","&nbsp;")&"</td><td>"&replace(formatcurrency(rs("unit_price"),2),"$0.00","&nbsp;")&"</td>")
	'1/20/2008 Michelle T. commented line above and added line below to match invoices found in timberline matches invoices in our system.  
	response.Write("<tr bgcolor=""#ffffff""><td>&nbsp;</td><td>"&rs("description")&"&nbsp;</td><td>"&units&"</td><td>"&replace(rs("unit_of_measure")," ","&nbsp;")&"</td><td>"&replace(formatcurrency(rs("unit_price"),2),"$0.00","&nbsp;")&"</td><td>"&replace(formatcurrency(rs("amount"),2),"$0.00","&nbsp;")&"</td>") 
	  
	  else
   ' ... but short ones need tags for proper display 
    'response.Write("<tr bgcolor=""#ffffff""><td>&nbsp;</td><td>"&replace(rs("description")," ","&nbsp;")&"&nbsp;</td><td colspan=""3"">&nbsp;</td>")
  'response.Write("<tr bgcolor=""#ffffff""><td>&nbsp;</td><td>"&replace(rs("description")," ","&nbsp;")&"&nbsp;</td><td>"&replace(rs("units"),"0","&nbsp;")&"</td><td>"&replace(rs("unit_of_measure")," ","&nbsp;")&"</td><td>"&replace(formatcurrency(rs("unit_price"),2),"$0.00","&nbsp;")&"</td>")
  '1/20/2008 Michelle T. commented line above and added line below to match invoices found in timberline matches invoices in our system.  
   response.Write("<tr bgcolor=""#ffffff""><td>&nbsp;</td><td>"&replace(rs("description")," ","&nbsp;")&"&nbsp;</td><td>"&units&"</td><td>"&replace(rs("unit_of_measure")," ","&nbsp;")&"</td><td>"&replace(formatcurrency(rs("unit_price"),2),"$0.00","&nbsp;")&"</td><td>"&replace(formatcurrency(rs("amount"),2),"$0.00","&nbsp;")&"</td>")

  end if
  rs.movenext
  ' Last record has price in last column
  if rs.eof then
   '1/20/2008 Michelle T. comment line below b/c total price appeared twice
    'response.Write("<td align=right>"&formatcurrency(price,2)&"&nbsp;</td></tr>"&vbcrlf)
  else
  	'intermediate records
    response.Write("</tr>"&vbcrlf)
	' leave vertical space after first
	if recidx=1 then
      response.Write("<tr bgcolor=""#ffffff""><td height=20>&nbsp;</td></tr>"&vbcrlf)
	end if
  end if
wend
rs.close
set cnn=nothing
set cmd=nothing
set rs=nothing
%>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="91%" align="center">&nbsp;</td>
  </tr>
  <tr> 
    <td align="center">&nbsp;</td>
  </tr>
  <% if billTax <> 0 then %>
  <tr>
    <td align="right" colspan=6>&nbsp;Total Tax:&nbsp;</td>
	<td width="9%" align="right"><%=formatcurrency(billTax)%>&nbsp;</td>
  </tr>
  <tr>
    <td align="right" colspan=6>&nbsp;Total Due:&nbsp;</td>
	<td width="9%" align="right"><%=formatcurrency(tot_invoice_amt)%>&nbsp;</td>
  </tr>
  <tr> 
  <% end if %>
    <td align="center"><strong><font size="5">GENERGY INTRANET DUPLICATE</font></strong></td>
  </tr>
</table>

<br>
</body>
</html>