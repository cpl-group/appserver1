 <%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'if not(allowGroups("Genergy Users")) then
'%><%
'end if


	dim ra_ed, ra_mac, ra_msc, re_msc, grt_ed, grt_s, grt_d, grt_r, grt_wp
	dim isql, rasql, grtsql, cnn1, rst1, madjid,  ra, grt, user
	set cnn1 = server.createobject("ADODB.connection")
	set rst1 = server.createobject("ADODB.recordset")
	set ra = server.createobject("ADODB.recordset")
	set grt = server.createobject("ADODB.recordset")
	user = getXmlUserName()
	cnn1.open getConnect(0,0,"dbcore")
	'if request.servervariables("ra_ed") and request.servervariables("ra_mac") and request.servervariables("ra_msc") then
		ra_ed  = request.form("ra_ed")
		ra_mac = request.form("ra_mac")
		ra_msc = request.form("ra_msc")
		re_msc = request.form("re_msc")
		if ra_ed <> "" then
			rasql = "select top 1 ra_ed, madjid from ratebuilderadjustments where ra_ed = '" &ra_ed& "'"
			ra.open rasql, cnn1
			if not ra.eof then
				madjid = ra("madjid")
				isql = "update ratebuilderadjustments set ra_mac = '" &ra_mac& "', ra_msc = '" &ra_msc& "', re_msc = '"&re_msc&"' output Inserted.madjid where madjid ='" &madjid& "'"
			else
				isql = "insert into ratebuilderadjustments ( ra_ed, ra_msc, ra_mac, re_msc, createdBy ) output Inserted.madjid values ('" &ra_ed& "','" &ra_msc& "','" &ra_mac& "','" &re_msc& "','"&user&"')"
			end if
			'response.write isql
			'response.end
			cnn1.execute isql
			ra.close
		ra_ed = null
		ra_mac = null
		ra_msc = null
		re_msc = null
		end if
	'end if
	'if request.servervariables("grt_ed") and request.servervariables("grt_s") and request.servervariables("grt_d") and request.servervariables("grt_r") and request.servervariables("grt_wp") then
		grt_ed = request.form("grt_ed")
		grt_s  = request.form("grt_s")
		grt_d  = request.form("grt_d")
		grt_r  = request.form("grt_r")
		grt_wp = request.form("grt_wp")
		if grt_ed <> "" then
			grtsql = "select top 1 grt_ed, madjid from ratebuilderadjustments where grt_ed = '" &grt_ed& "'"
			grt.open grtsql, cnn1
			if not grt.eof then
				madjid = grt("madjid")
				isql = "update ratebuilderadjustments set grt_s = '" &grt_s& "', grt_d = '" &grt_d& "', grt_r = '" &grt_r& "', grt_wp = '" &grt_wp& "' output Inserted.madjid where madjid ='" &madjid& "'"
			else
				isql = "insert into ratebuilderadjustments ( grt_ed, grt_s, grt_d, grt_r, grt_wp, createdBy ) output Inserted.madjid values ('" &grt_ed& "','" &grt_s&  "','" &grt_d& "','" &grt_r& "','" &grt_wp& "','"&user&"')"
			end if
			'response.write isql
			'response.end
			cnn1.execute isql
			grt.close
		grt_ed = null
		grt_s = null
		grt_d = null
		grt_r = null
		grt_wp = null
		end if
	'end if
	
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 15">
<link rel=File-List href="Monthly%20Adjustments_files/filelist.xml">
<style id="Monthly Adjustments_21795_Styles">
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
.xl1521795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6321795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6421795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border-top:1.0pt solid windowtext;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6521795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border-top:none;
	border-right:none;
	border-bottom:none;
	border-left:1.0pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6621795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6721795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:1.0pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6821795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border-top:none;
	border-right:1.0pt solid windowtext;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6921795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border-top:none;
	border-right:none;
	border-bottom:1.0pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7021795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7121795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border-top:1.0pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7221795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7321795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl7421795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:"Short Date";
	text-align:general;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7521795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:right;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7621795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:normal;}
.xl7721795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7821795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:middle;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:none;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7921795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:right;
	vertical-align:bottom;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl8021795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:bottom;
	border:1.0pt solid windowtext;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl8121795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:"Short Date";
	text-align:center;
	vertical-align:middle;
	background:#D9D9D9;
	mso-pattern:black none;
	white-space:nowrap;}
.xl8221795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	background:#D9D9D9;
	mso-pattern:black none;
	white-space:nowrap;}
.xl8321795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:bottom;
	background:#D9D9D9;
	mso-pattern:black none;
	white-space:nowrap;}
.xl8421795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:"mmm\\-yy";
	text-align:center;
	vertical-align:middle;
	background:#D9D9D9;
	mso-pattern:black none;
	white-space:nowrap;}
.xl8521795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl8621795
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:12.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
-->
</style>
</head>

<body>
<!--[if !excel]>&nbsp;&nbsp;<![endif]-->
<!--The following information was generated by Microsoft Excel's Publish as Web
Page wizard.-->
<!--If the same item is republished from Excel, all information between the DIV
tags will be replaced.-->
<!----------------------------->
<!--START OF OUTPUT FROM EXCEL PUBLISH AS WEB PAGE WIZARD -->
<!----------------------------->

<div id="Monthly Adjustments_21795" align=center x:publishsource="Excel">
	<form name="Fuel Sheet Adjustments" method="post" action="monthlyadjustments.asp">
		<table border=0 cellpadding=0 cellspacing=0 width=700 style='border-collapse:
		 collapse;table-layout:fixed;width:728pt'>
		 <col width=24 style='mso-width-source:userset;mso-width-alt:877;width:18pt'>
		 <col class=xl7021795 width=100 style='mso-width-source:userset;mso-width-alt:
		 5449;width:112pt'>
		 <col class=xl7021795 width=16 style='mso-width-source:userset;mso-width-alt:
		 1170;width:24pt'>
		 <col class=xl7021795 width=100 style='mso-width-source:userset;mso-width-alt:
		 5449;width:112pt'>
		 <col class=xl7021795 width=16 style='mso-width-source:userset;mso-width-alt:
		 1170;width:24pt'>
		 <col class=xl7021795 width=100 style='mso-width-source:userset;mso-width-alt:
		 5449;width:112pt'>
		 <col class=xl7021795 width=100 style='mso-width-source:userset;mso-width-alt:
		 4096;width:84pt'>
		 <col class=xl7021795 width=100 span=2 style='mso-width-source:userset;
		 mso-width-alt:5449;width:112pt'>
		 <col width=24 style='mso-width-source:userset;mso-width-alt:877;width:18pt'>
		 <tr height=21 style='height:15.75pt'>
		  <td height=21 class=xl6321795 width=24 style='height:15.75pt;width:18pt'><a
		  name="RANGE!A1:J9">&nbsp;</a></td>
		  <td class=xl7121795 width=100 style='width:112pt'>&nbsp;</td>
		  <td class=xl7121795 width=16 style='width:24pt'>&nbsp;</td>
		  <td class=xl7121795 width=100 style='width:112pt'>&nbsp;</td>
		  <td class=xl7121795 width=16 style='width:24pt'>&nbsp;</td>
		  <td class=xl7121795 width=100 style='width:112pt'>&nbsp;</td>
		  <td class=xl7121795 width=100 style='width:84pt'>&nbsp;</td>
		  <td class=xl7121795 width=100 style='width:112pt'>&nbsp;</td>
		  <td class=xl7121795 width=100 style='width:112pt'>&nbsp;</td>
		  <td class=xl6421795 width=24 style='width:18pt'>&nbsp;</td>
		 </tr>
		 <tr height=21 style='height:15.75pt'>
		  <td height=21 class=xl6521795 style='height:15.75pt'>&nbsp;</td>
		  <td colspan=5 class=xl8621795>SC9 Rate II</td>
		  <td class=xl7221795></td>
		  <td class=xl7221795></td>
		  <td class=xl7221795></td>
		  <td class=xl6621795>&nbsp;</td>
		 </tr>
		 <tr height=21 style='height:15.75pt'>
		  <td height=21 class=xl6521795 style='height:15.75pt'>&nbsp;</td>
		  <td rowspan=2 class=xl7321795 width=149 style='border-bottom:.5pt solid black;
		  width:112pt'>Effective Date</td>
		  <td class=xl7321795 width=32 style='width:24pt'></td>
		  <td colspan=3 class=xl8621795>Retail Access</td>
		  <td class=xl8621795>Residential</td>
		  <td class=xl7221795></td>
		  <td class=xl7221795></td>
		  <td class=xl6621795>&nbsp;</td>
		 </tr>
		 <tr height=21 style='height:15.75pt'>
		  <td height=21 class=xl6521795 style='height:15.75pt'>&nbsp;</td>
		  <td class=xl7621795 width=32 style='width:24pt'>&nbsp;</td>
		  <td class=xl7721795>MAC</td>
		  <td class=xl7721795>&nbsp;</td>
		  <td class=xl7721795>MSC</td>
		  <td class=xl7721795>MSC</td>
		  <td colspan=2 class=xl7721795>NYC GRT Tariffs</td>
		  <td class=xl6621795>&nbsp;</td>
		 </tr>
		 <tr height=21 style='height:15.75pt'>
		  <td height=21 class=xl6521795 style='height:15.75pt'>&nbsp;</td>
		  <td class=xl8121795><input class=box name="ra_ed" value="<%=ra_ed%>" autofocus/></td>
		  <td class=xl7421795></td>
		  <td class=xl8221795><input class=box name="ra_mac" value="<%=ra_mac%>"></td>
		  <td class=xl8521795></td>
		  <td class=xl8221795><input class=box name="ra_msc" value="<%=ra_msc%>"></td>
		  <td class=xl7221795><input class=box name="re_msc" value="<%=re_msc%>"></td>
		  <td class=xl7921795>Effective Date</td>
		  <td class=xl8421795><input class=box name="grt_ed" value="<%=grt_ed%>"></td>
		  <td class=xl6621795>&nbsp;</td>
		 </tr>
		 <tr height=21 style='height:15.75pt'>
		  <td height=21 class=xl6521795 style='height:15.75pt'>&nbsp;</td>
		  <td class=xl7021795></td>
		  <td class=xl7021795></td>
		  <td class=xl7021795></td>
		  <td class=xl7021795></td>
		  <td class=xl7021795></td>
		  <td class=xl7221795></td>
		  <td class=xl7521795>Supply</td>
		  <td class=xl8221795><input class=box name="grt_s" value="<%=grt_s%>"></td>
		  <td class=xl6621795>&nbsp;</td>
		 </tr>
		 <tr height=21 style='height:15.75pt'>
		  <td height=21 class=xl6521795 style='height:15.75pt'>&nbsp;</td>
		  <td class=xl7221795></td>
		  <td class=xl7221795></td>
		  <td class=xl7221795></td>
		  <td class=xl7221795></td>
		  <td class=xl7221795></td>
		  <td class=xl7221795></td>
		  <td class=xl7521795>Delivery</td>
		  <td class=xl8221795><input class=box name="grt_d" value="<%=grt_d%>"></td>
		  <td class=xl6621795>&nbsp;</td>
		 </tr>
		 <tr height=21 style='height:15.75pt'>
		  <td height=21 class=xl6521795 style='height:15.75pt'>&nbsp;</td>
		  <td class=xl7221795></td>
		  <td class=xl7221795></td>
		  <td class=xl7221795></td>
		  <td class=xl7221795></td>
		  <td class=xl7221795></td>
		  <td class=xl7221795></td>
		  <td class=xl7521795>Residential</td>
		  <td class=xl8221795><input class=box name="grt_r" value="<%=grt_r%>"></td>
		  <td class=xl6621795>&nbsp;</td>
		 </tr>
		 <tr height=22 style='height:16.5pt'>
		  <td height=22 class=xl6521795 style='height:16.5pt'>&nbsp;</td>
		  <td class=xl7021795></td>
		  <td class=xl7021795></td>
		  <td class=xl7021795></td>
		  <td class=xl7021795></td>
		  <td class=xl7021795></td>
		  <td class=xl7021795></td>
		  <td class=xl7521795>White Plains</td>
		  <td class=xl8321795><input class=box name="grt_wp" value="<%=grt_wp%>"></td>
		  <td class=xl6621795>&nbsp;</td>
		 </tr>
		 <tr height=22 style='height:16.5pt'>
		  <td height=22 class=xl6521795 style='height:16.5pt'>&nbsp;</td>
		  <td class=xl7021795></td>
		  <td class=xl7021795></td>
		  <td class=xl7021795></td>
		  <td class=xl7021795></td>
		  <td class=xl7021795><input type="submit" name="action" value="Save" class="standard" /></td>
		  <td class=xl7021795></td>
		  <td class=xl7021795></td>
		  <td class=xl7021795></td>
		  <td class=xl6621795>&nbsp;</td>
		 </tr>
		 <tr height=22 style='height:16.5pt'>
		  <td height=22 class=xl6721795 style='height:16.5pt'>&nbsp;</td>
		  <td class=xl6921795>&nbsp;</td>
		  <td class=xl6921795>&nbsp;</td>
		  <td class=xl6921795>&nbsp;</td>
		  <td class=xl6921795>&nbsp;</td>
		  <td class=xl6921795>&nbsp;</td>
		  <td class=xl6921795>&nbsp;</td>
		  <td class=xl6921795>&nbsp;</td>
		  <td class=xl6921795>&nbsp;</td>
		  <td class=xl6821795>&nbsp;</td>
		 </tr>
		 <![if supportMisalignedColumns]>
		 <tr height=0 style='display:none'>
		  <td width=24 style='width:18pt'></td>
		  <td width=149 style='width:112pt'></td>
		  <td width=32 style='width:24pt'></td>
		  <td width=149 style='width:112pt'></td>
		  <td width=32 style='width:24pt'></td>
		  <td width=149 style='width:112pt'></td>
		  <td width=112 style='width:84pt'></td>
		  <td width=149 style='width:112pt'></td>
		  <td width=149 style='width:112pt'></td>
		  <td width=24 style='width:18pt'></td>
		 </tr>
		 <![endif]>
		</table>
	</form>

<%

	rasql = "select * from ratebuilderadjustments where ra_ed > '1/1/2000' order by ra_ed desc"
	grtsql = "select * from ratebuilderadjustments where grt_ed > '1/1/2000' order by grt_ed desc"
	ra.open rasql, cnn1
	grt.open grtsql, cnn1 %>
	<table>
		<tr width="100%">
			<td width = "50%" valign="top">
				<table>
					<%if not ra.eof then %>
						<tr>
							<td class=xl7721795>Effective Date</td>
							<td class=xl7721795>MAC</td>
							<td class=xl7721795>MSC</td>
							<td class=xl7721795>Residential<br>MSC</td>
						</tr>
						<% do until ra.eof %>
							<tr>
								<td style="font-family:verdana;font-size:12"><%= ra("ra_ed") %></td>
								<td align="right" style="font-family:verdana;font-size:12"><%= ra("ra_mac") %></td>
								<td align="right" style="font-family:verdana;font-size:12"><%= ra("ra_msc") %></td>
								<td align="right" style="font-family:verdana;font-size:12"><%= ra("re_msc") %></td>
							</tr>
							<% ra.movenext
						loop
						ra.close
					end if %>
				</table>
			</td>
			<td width="50%" valign="top">
				<table>
					<%if not grt.eof then %>
						<tr>
							<td class=xl7721795>Effective Date</td>
							<td class=xl7721795>Supply</td>
							<td class=xl7721795>Delivery</td>
							<td class=xl7721795>Residential</td>
							<td class=xl7721795>White Plains</td>
						</tr>
						<% do until grt.eof %>
							<tr>
								<td style="font-family:verdana;font-size:12"><%= grt("grt_ed") %></td>
								<td align="right" style="font-family:verdana;font-size:12"><%= grt("grt_s") %></td>
								<td align="right" style="font-family:verdana;font-size:12"><%= grt("grt_d") %></td>
								<td align="right" style="font-family:verdana;font-size:12"><%= grt("grt_r") %></td>
								<td align="right" style="font-family:verdana;font-size:12"><%= grt("grt_wp") %></td>
							</tr>
							<% grt.movenext
						loop
						grt.close
					end if %>
				</table>
			</td>
		</tr>
	</table>
			
</div>
<!----------------------------->
<!--END OF OUTPUT FROM EXCEL PUBLISH AS WEB PAGE WIZARD-->
<!----------------------------->
</body>

</html>
