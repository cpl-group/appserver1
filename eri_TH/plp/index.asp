<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim bldgid,bldgname
dim rs, cnn, cmd,graphtype,strsql, printView
bldgid 		= trim(request("bldgid")) '"99"
printView = trim(request("prview"))
if printView = "" then
	printView = 0
end if

Set cnn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.recordset")

cnn.Open getConnect(0,0,"Engineering")

strsql = "select address from tlbldg where bldgnum = '" & bldgid &"'"

rs.Open strsql, cnn,0

if not rs.eof then 
	bldgname =rs("address")
else
	bldgname ="Unknown"
end if

dim link
link = "http://pdfmaker.genergyonline.com/pdfmaker/pdfReport_v2.asp?landscape=true&devIP=appserver1.genergy.com&sn=/eri_TH/plp/index.asp&qs="&server.urlencode("bldgid="&Request("bldgID")&"&prview=1"&"&graphtype=6" )
%>
<html>
<head>
<title>Power Availability Chart</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">		

</head>
<body bgcolor="#ffffff" text="#000000">
<%
if printView = 0 then ' this is the normal view
%>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0" >
  <tr bgcolor = "#6699cc"> 
    <td bgcolor><span class="standardheader"><font size="2">Power Availability 
      Chart for <%=bldgname%></font></span></td>
	<td align="right"><a href = "<%=link%>"><img src = "/images/print_pdf.gif" border=0></a></td>
  </tr>
  <tr>
    <td colspan="2" align="center"><img src="/eri_th/plp/charts.asp?bldgid=<%=request("bldgid")%>&graphtype=6" width="800" height="600"></td>
  </tr>
  <tr> 
    <td colspan = "2" bgcolor="#6699cc"><strong><font color="#FFFFFF" size="1" face="Arial, Helvetica, sans-serif">Power availability calculations are based 
    on power distribution system information gathered by Genergy or made available to Genergy by the property’s management.  
    Genergy normally applies a 15-20% safety factor reduction when calculating power availability figures.  
    For more information on the use of safety factors or to modify the safety factor calculation criteria we encourage you to 
    contact our office.</font></strong></td>
  </tr>
</table>


<%
	else		' this is the pdf view
	%>
	<table width="100%" border="0" cellpadding="0" cellspacing="0">
		<tr>
			<td bgcolor = "6699cc" align="center" colspan="2">
				<span class="standardheader"><font size="4">Power Availability 
				Chart for <%=bldgname%></font></span>
			</td>
		</tr>
		<tr>
			<td width="20%">&nbsp;&nbsp;&nbsp;&nbsp;</td>
			<td align="right" width="80%">
				<img src="/eri_th/plp/charts.asp?bldgid=<%=request("bldgid")%>&graphtype=6" width="1000" height="750">
			</td>
		</tr>
		<tr>
			<td width="20%">&nbsp;&nbsp;&nbsp;&nbsp;</td>
			<td width="80%">
				<strong><font size="2" face="Arial, Helvetica, sans-serif">Power availability calculations 
				are based on power distribution system information gathered by Genergy or made available to Genergy by the property’s management.  
				Genergy normally applies a 15-20% safety factor reduction when calculating power availability figures.  
				For more information on the use of safety factors or to modify the safety factor calculation criteria 
				we encourage you to contact our office. </font></strong>
			</td>
		</tr>
	</table>
	<%
end if %>
</body>
</html>
