<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/nonsecure.inc"-->
<%
'if not(allowGroups("Genergy Users")) then
'%><%
'end if

	dim building
	building=trim(request("building"))

	dim cnn1, rst1, rst2, strsql, insertSql
	set cnn1 = server.createobject("ADODB.connection")
	set rst1 = server.createobject("ADODB.recordset")
	cnn1.open getlocalconnect(building)
	
	dim pid, byear, bperiod, utilid, portfolio
	pid = request("pid")
	byear = request("byear")
	bperiod = request("bperiod")
	utilid = request("utilityid")
	
%>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<style id="BillProcessor_6749_Styles">
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
.xl156749
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
.xl636749
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
	text-align:center;
	vertical-align:bottom;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl646749
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:16.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:bottom;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl656749
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:20.0pt;
	font-weight:700;
	font-style:normal;
	text-decoration:none;
	font-family:Calibri, sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:center;
	vertical-align:bottom;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
-->
</style>
</head>

<body>
<div id="BillProcessor_6749" align=center x:publishsource="Excel">
<%
if building<>"" then
	bldg = Replace(building, "+", " ")
	bldg = Replace(bldg, "%20", " ")
	rst1.open "SELECT location, b.bldgnum, b.portfolioid,billurl,logo, logoh, logow, p.portfolio FROM buildings b, portfolio p, billtemplates bt WHERE b.portfolioid=p.id AND bt.id=p.templateid AND bldgnum='"&bldg&"'	", cnn1
		if not rst1.eof then 
			pid = rst1("portfolioid")
			portfolio=rst1("portfolio")		
		end if
		rst1.close 
end if
%>
<table border=0 cellpadding=0 cellspacing=0 width=813 style='border-collapse:
 collapse;table-layout:fixed;width:460pt'>
 <col width=100 style='mso-width-source:userset;mso-width-alt:3657;width:75pt'>
 <col width=260 style='mso-width-source:userset;mso-width-alt:5851;width:220pt'>
 <col width=157 style='mso-width-source:userset;mso-width-alt:2084;width:143pt'>
 <col width=115 style='mso-width-source:userset;mso-width-alt:4205;width:86pt'>
 <col width=181 style='mso-width-source:userset;mso-width-alt:6619;width:136pt'>
<form name=load action="processbills.asp">
	 <tr height=35 style='height:26.25pt'>
	  <td colspan=4 height=35 class=xl656749 width=813 style='height:26.25pt;
	  width:460pt'><a name="RANGE!A1:E50">LeFrak Properties</a></td>
	  <td class=xl646749><input size="1" class=box name="emonth" value="<%=bperiod%>" /><input  size="1" class=box name="eyear" value="<%=byear%>" /><br><input type="submit" name="action" value="Load" class="standard" /></td>
	 </tr>
</form>	 
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl646749 style='height:15.0pt'></td>
  <td class=xl646749></td>
  <td class=xl646749></td>
  <td class=xl646749></td>
  <td class=xl646749></td>
 </tr>
 <form target="bill_pop" name=bills action="bills.asp" onsubmit="window.open('about:blank','bill_pop','width=300,height=500');">
	 <tr height=28 style='height:21.0pt'>
	  <td height=28 class=xl156749 style='height:21.0pt'></td>
	  <td class=xl646749><INPUT type=submit name="action" value="Create"></td>
	  <td class=xl646749></td>
	  <td class=xl646749><INPUT type=submit name="action" value="Delete"></td>
	  <td class=xl646749></td>
	 </tr>
	 <tr height=28 style='height:21.0pt'>
	  <td height=28 class=xl156749 style='height:21.0pt'></td>
	  <td class=xl646749><INPUT type=submit name="action" value="Post"></td>
	  <td class=xl646749></td>
	  <td class=xl646749><INPUT type=submit name="action" value="UnPost"></td>
	  <td class=xl646749></td>
	 </tr>
	 <tr height=20 style='height:15.0pt'>
	  <td height=20 class=xl156749 style='height:15.0pt'></td>
	  <td class=xl156749></td>
	  <td class=xl156749></td>
	  <td class=xl156749></td>
	  <td class=xl156749></td>
	 </tr>
	<INPUT type=hidden name=pid value=<%=pid%>></INPUT>
	<INPUT type=hidden name=by value=<%=byear%>></INPUT>
	<INPUT type=hidden name=bp value=<%=bperiod%>></INPUT>
	<INPUT type=hidden name=utilid value=<%=utilid%>></INPUT>	 
 </form>
 <%

	Dim fso, strFileName, i
	Set fso = CreateObject("Scripting.FileSystemObject")
	dim ctime,absfile,PDFName
	dim bldg, bldgname, root, pdfdir, link, hasfile
 
	dim sql 
	if pid =163 then
			sql = "select db.bldgnum, db.bldgname, count(meterid) as mc from DailyExportBuildings db left join meters m on m.bldgnum = db.bldgnum where m.online=1 group by db.bldgnum,db.bldgname order by bldgname asc"
	else 
		sql = "select upper(db.bldgnum) as bldgnum, db.bldgname, count(meterid) as mc from buildings db left join meters m on m.bldgnum = db.bldgnum where m.online=1 and db.portfolioid = "&pid&" group by db.bldgnum,db.bldgname order by bldgname asc"
	end if
	rst1.open sql, cnn1
	do until rst1.eof
 
 %>
<%
		bldg = rst1("bldgnum")
		root = "D:\WebSites\isabella\genergyonline.com\pdfmaker\"
		pdfdir =  portfolio & "\" & ucase(bldg) & "\"
		PDFName = ucase(bldg) & byear & bperiod & utilid & "1.pdf"
		ctime = year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)	
		
		if fso.fileexists(root&pdfdir&PDFName) then
			link = pdfdir & pdfname & "&dt=" & ctime
			hasfile = true
		else
			hasfile = false
		end if
%>		
		<form name=consolepdf action="consolepdf.asp" method="post" target="pdf_pop" onsubmit="window.open('about:blank','pdf_pop','width=10,height=10');">
		 <tr height=20 style='height:15.0pt'>
		  <td height=20 class=xl156749 style='height:15.0pt'><%= bldg %></td>
		  <td class=xl156749><%= rst1("bldgname") %></td>
		  <td class=xl156749 align=right> | <b><%= rst1("mc") %></b></td>
		  <td class=xl636749><INPUT type=submit value="Generate"></INPUT></td>
		  <% if hasfile then %>
		  <td class=xl156749><a style="font-family:arial;font-size:12;text-decoration:none;" href="http://pdfmaker.genergyonline.com/pdfMaker/<%= link %>" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'blue'"><b><%=pdfname%></b></a> </td>
		  <% else %>
		  <td class=xl156749><%=pdfname%></td>
		  <% end if %>
		 </tr>
		 	<INPUT type=hidden name=bldg value=<%=bldg%>></INPUT>
			<INPUT type=hidden name=byear value=<%=byear%>></INPUT>
			<INPUT type=hidden name=bperiod value=<%=bperiod%>></INPUT>
			<INPUT type=hidden name=utilid value=<%=utilid%>></INPUT>
		</form>
<%
	rst1.movenext
	loop
	rst1.close
%>
 
 <![if supportMisalignedColumns]>
 <tr height=0 style='display:none'>
  <td width=100 style='width:75pt'></td>
  <td width=160 style='width:120pt'></td>
  <td width=57 style='width:43pt'></td>
  <td width=115 style='width:86pt'></td>
  <td width=181 style='width:136pt'></td>
 </tr>
 <![endif]>
</table>

</div>


<!----------------------------->
<!--END OF OUTPUT FROM EXCEL PUBLISH AS WEB PAGE WIZARD-->
<!----------------------------->
</body>

</html>
