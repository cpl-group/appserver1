<%option explicit%>
<!--#include file="adovbs.inc"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'http params
dim crdate,d,c,j,total,currentvendor,total1,total2,gtotal,gtotal1,gtotal2, costcode

d = request("d")
c = request("c")

'Cost code for Materials cost is 005 and subcontractor cost 004. the value passed to this page as cost code should be
'the code you wish to EXCLUDE. So, if all you want is Materials cost, then costcode should=004 and vice versa. 
costcode = trim(request("costcode"))

if request("ji")="" then 
	j = request ("jg")
else
	j = request ("ji")
end if

gtotal=0
'adodb vars
dim cnn, cmd, rs, prm
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")

' open connection
cnn.open getConnect(0,0,"Intranet")
cnn.CursorLocation = adUseClient

rs.open "SELECT crdate FROM sysobjects WHERE name = 'GY_activity_ara_status'",cnn
crdate=rs(0)
rs.close


' specify stored procedure to run based on company

cmd.CommandText = "sp_material_cost"

cmd.CommandType = adCmdStoredProc

Set prm = cmd.CreateParameter("c", adchar, adParamInput,2)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("j", advarchar, adParamInput,9)
cmd.Parameters.Append prm

' assign internal name to stored procedure
cmd.Name = "test"
Set cmd.ActiveConnection = cnn

'return set to recordset rs
cnn.test  c,j, rs

Dim title

if costcode = "004" then title = "Material Cost" else title = "Subcontractor Cost" end if

%>
<html>
<head>
<title>Genergy War Room - <%=title%></title>

<script type="text/javascript">

function openWindow(jobno,company)
{

// Append jobno to http link

if (company=="IL") {
var urlLink     = "/um/war/jc/jc.asp?c=" + company + "&ji=" + jobno
}
else {
var urlLink     = "/um/war/jc/jc.asp?c=" + company + "&jg=" + jobno
}

// Open new window and customize window settings
window.open(urlLink,"window","scrollbars=no,width=900,height=600,resizeable")

}

</script>

<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>

<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#FFFFFF">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699cc">
  <td><span class="standardheader">Materials Cost</span></td>
</tr>
</table>
<table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#cccccc">
<tr bgcolor="#dddddd" style="font-weight:bold;">
 <td width="23%">Vendor</td>
 <td width="11%" style="border-left:1px solid #cccccc;">Invoice #</td>
 <td width="11%" style="border-left:1px solid #cccccc;">Invoice Date</td>
 <td width="11%" style="border-left:1px solid #cccccc;">PO (cost cat)</td>
 <td width="11%" style="border-left:1px solid #cccccc;">Original Amt.</td>
 <td width="11%" style="border-left:1px solid #cccccc;">Amount Paid</td>
 <td width="11%" style="border-left:1px solid #cccccc;">Amount Open</td>
</tr>
	 
      <%
if not rs.eof then
	  
total = 0
total1 = 0
total2 = 0
gtotal = 0
gtotal1 = 0
gtotal2 = 0

currentvendor = cstr(rs("vendor"))

while not rs.EOF 



if trim(cstr(rs("category"))) <> costcode and trim(cstr(rs("category"))) <> "003" and trim(cstr(rs("category"))) <> "002" and trim(cstr(rs("category"))) <> "001" and trim(cstr(rs("category"))) <> "000" then 

gtotal = gtotal + rs("amount")
gtotal1= gtotal1 + rs("amount_paid")
gtotal2 = gtotal2 + rs("amount_open")

	if rs("vendor") = currentvendor then
	'response.write  rs("amount") &"<BR>"
	total = total + rs("amount")
	total1= total1 + rs("amount_paid")
	total2 = total2 + rs("amount_open")
	
	else
	
	%>
	<tr bgcolor="#ffffff">
	  <td style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;">&nbsp;</td>
	  <td style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;">&nbsp;</td>
	  <td style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;">&nbsp;</td>
	  <td style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;">&nbsp;</td>
	  <td style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;" align="right"><b><%=formatcurrency(total,2)%></b></td>
	  <td style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;" align="right"><b><%=formatcurrency(total1,2)%></b></td>
	  <td style="border-bottom:1px solid #cccccc;border-left:1px solid #cccccc;" align="right"><b><%=formatcurrency(total2,2)%></b></td>
	</tr>
	<%if total=0 then%>
		<!-- /table -->
	<%end if%>
	
	<%
	'gtotal=gtotal + rs("amount_paid")
	total =  rs("amount")
	total1 =  rs("amount_paid")
	total2 =  rs("amount_open")
	
	currentvendor= rs("vendor")
	
	end if
	%>
	<tr bgcolor="#ffffff"> 
	  <td style="border-left:1px solid #cccccc;"><%=rs("name")%></td>
	  <td style="border-left:1px solid #cccccc;"><%=rs("invoice")%></td>
	  <td style="border-left:1px solid #cccccc;"><%=rs("date")%></td>
	  <td style="border-left:1px solid #cccccc;">
	  <% 
	  Dim objRegExp
	  Set objRegExp = New RegExp
	  objRegExp.Pattern = "^\d{4}\.\d+$"
	  if objRegExp.Test(rs("po")) then
	  %>
	  <a href="/um/opslog/poview.asp?po=<%=rs("po")%>"><%=rs("po")%></a>
	  <% else %>
	  <%=rs("po")%>
	  <% end if %>
	  (<%=rs("category")%>)
	  </td>
	  <td style="border-left:1px solid #cccccc;" align="right"><%=formatcurrency(rs("amount"),2)%></td>
	  <td style="border-left:1px solid #cccccc;" align="right"><%=formatcurrency(rs("amount_paid"),2)%></td>
	  <td style="border-left:1px solid #cccccc;" align="right"><%=formatcurrency(rs("amount_open"),2)%></td>
	</tr>
	<%
end if 
 rs.movenext
wend
%>

<tr bgcolor="#ffffff"> 
  <td style="border-left:1px solid #cccccc;">&nbsp;</td>
  <td style="border-left:1px solid #cccccc;">&nbsp;</td>
  <td style="border-left:1px solid #cccccc;">&nbsp;</td>
  <td style="border-left:1px solid #cccccc;">&nbsp;</td>
  <td style="border-left:1px solid #cccccc;" align="right"><b><%=formatcurrency(total,2)%></b></td>
  <td style="border-left:1px solid #cccccc;" align="right"><b><%=formatcurrency(total1,2)%></b></td>
  <td style="border-left:1px solid #cccccc;" align="right"><b><%=formatcurrency(total2,2)%></b></td>
</tr>
	  
<tr bgcolor="#eeeeee">
  <td style="border-left:1px solid #000000;border-top:1px solid #000000;border-bottom:1px solid #000000;">&nbsp;</td>
  <td style="border-left:1px solid #cccccc;border-top:1px solid #000000;border-bottom:1px solid #000000;">&nbsp;</td>
  <td style="border-left:1px solid #cccccc;border-top:1px solid #000000;border-bottom:1px solid #000000;">&nbsp;</td>
  <td style="border-left:1px solid #cccccc;border-top:1px solid #000000;border-bottom:1px solid #000000;">&nbsp;</td>
  <td style="border-left:1px solid #cccccc;border-top:1px solid #000000;border-bottom:1px solid #000000;" align="right"><b><%=formatcurrency(gtotal,2)%></b></td>
  <td style="border-left:1px solid #cccccc;border-top:1px solid #000000;border-bottom:1px solid #000000;" align="right"><b><%=formatcurrency(gtotal1,2)%></b></td>
  <td style="border-left:1px solid #cccccc;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000;" align="right"><b><%=formatcurrency(gtotal2,2)%></b></td>
</tr>
<tr bgcolor="#ffffff">
  <td colspan="7">Updated as of <%=crdate%></td> 
</tr>
<% else %>
<tr bgcolor="#ffffff"><td colspan=7>No records found</td></tr>
<%
set cnn = nothing 
end if
%>
<tr bgcolor="#ffffff">
  <td colspan="7"><input type="button" value="Close Window" onclick="window.close();"></td>
</tr>
</table>
</body>
</html>
