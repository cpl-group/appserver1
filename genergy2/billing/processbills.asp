<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/nonsecure.inc"-->
<%
'if not(allowGroups("Genergy Users")) then
'%><%
'end if

	dim building, action, zipname, zipfull, ziplink, utilfilter, utilbills
	building=trim(request("building"))
	action = trim(request("action"))
	dim cnn1, rst1, rst2, rst3, strsql, insertSql
	set cnn1 = server.createobject("ADODB.connection")
	set rst1 = server.createobject("ADODB.recordset")
	set rst2 = server.createobject("ADODB.recordset")
	set rst3 = server.createobject("ADODB.recordset")
	cnn1.open getlocalconnect(building)
	
	dim pid, byear, bperiod, utilid, pname, p
	pid = request("pid")
	byear = request("byear")
	bperiod = request("bperiod")
	utilfilter = request("utilfilter")
	dim thisdate 
	thisdate = dateadd ("m",-1,now)

	if byear = "" then byear = year(thisdate) end if
	if bperiod = "" then bperiod = month(thisdate) end if
	function toNumb(val)
		if val="" or isnull(val) then
			val = 0
		end if
		if IsNumeric(CStr(val)) then
			toNumb = cdbl(val)
		end if
	end function	
	Function ConvertTime(intTotalSecs)
		Dim intHours,intMinutes,intSeconds,Time
		intHours = intTotalSecs \ 3600
		intMinutes = (intTotalSecs Mod 3600) \ 60
		intSeconds = intTotalSecs Mod 60
		ConvertTime = LPad(intHours) & " h : " & LPad(intMinutes) & " m : " & LPad(intSeconds) & " s"
	End Function
	Function LPad(v) 
		LPad = Right("0" & v, 2) 
	End Function
	function hrsago(ptime)
		dim d, h, m
		h = datediff("h",ptime, now())
		if h > 24 then 
			d = datediff("d", ptime, now())
			hrsago = d & "d old"
		elseif h < 1 then
			m = datediff("n", ptime, now())
			hrsago = m & "m old"
		else
			hrsago = h & "h old"
		end if
	end function	
	function fw(txt)
		dim ar, w, a
		w=""
		ar = split(txt, " ")
		for each a in ar
			w = w & replace(left(a,1),"(","")
		next
		fw = w
	end function
	function utilicon(uid)
		select case uid
		case"1" 'steam
			utilicon = "steam.png"
		case"2"	'elec
			utilicon = "electric.colored.png"
		case"3"	'cold water
			utilicon = "water.colored.png"
		case"4"	'gas
			utilicon = "gas.png"
		case"10" 'hot water
			utilicon = "water.orange.png"
		end select
		utilicon = "images\" & utilicon 
	end function
	Function CheckRemoteURL(fileURL)
		ON ERROR RESUME NEXT
		Dim xmlhttp

		Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")

		xmlhttp.open "GET", fileURL, False
		xmlhttp.send
		If(Err.Number<>0) then
			Response.Write "Could not connect to remote server"
		else
			Select Case Cint(xmlhttp.status)
				Case 200, 202, 302
					Set xmlhttp = Nothing
					CheckRemoteURL = True
				Case Else
					Set xmlhttp = Nothing
					CheckRemoteURL = False
			End Select
		end if
		ON ERROR GOTO 0
	End Function	
%>

<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 15">
<meta http-equiv="refresh" content="120;url=processbills.asp?pid=<%= pid %>&bperiod=<%= bperiod %>&pyear=<%= byear %>&utilfiler=<%= utilfilter %>"> 

<style id="BillProcessor_31467_Styles">
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
.xl1531467
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
.xl6331467
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
.xl6431467
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
.xl6531467
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
.xl6631467
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
	text-align:left;
	vertical-align:middle;
	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6731467
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
	text-align:left;
	vertical-align:middle;

	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl6831467
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
.xl6931467
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
.xl7031467
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
	text-align:left;
	vertical-align:middle;

	mso-background-source:auto;
	mso-pattern:auto;
	white-space:nowrap;}
.xl7131467
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
.xl7231467
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
.xl7331467
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:11.0pt;
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
<script>import {Spinner} from "spin.js"</script>

<script>
function pdf(bldg, byear, bperiod, utilid, action){
	var opts = {
	  lines: 20, // The number of lines to draw
	  length: 0, // The length of each line
	  width: 6, // The line thickness
	  radius: 9, // The radius of the inner circle
	  scale: 0.5, // Scales overall size of the spinner
	  corners: 0, // Corner roundness (0..1)
	  color: '#ff0080', // CSS color or array of colors
	  fadeColor: 'transparent', // CSS color or array of colors
	  speed: 1, // Rounds per second
	  rotate: 0, // The rotation offset
	  animation: 'spinner-line-fade-more', // The CSS animation name for the lines
	  direction: 1, // 1: clockwise, -1: counterclockwise
	  zIndex: 2e9, // The z-index (defaults to 2000000000)
	  className: 'spinner', // The CSS class to assign to the spinner
	  top: '50%', // Top position relative to parent
	  left: '50%', // Left position relative to parent
	  shadow: '0 0 1px transparent', // Box-shadow for the lines
	  position: 'absolute' // Element positioning
	};
    if (bldg.length == 0) {
        document.getElementById(bldg+".link").innerHTML = "";
        return;
    } else {
       // var xmlhttp = new XMLHttpRequest();
       // xmlhttp.onreadystatechange = function() {
       //     if (this.readyState == 4 && this.status == 200) {
       //        var target = document.getElementById(bldg+".link");
		//	   var spinner = new Spinner(opts).spin(target);
       //     }
       // };
        xmlhttp.open("GET", "consolepdf.asp?bldg="+bldg+"&byear="+byear+"&bperiod="+bperiod+"&utilfilter="+utilid+"&action="+action,true);
        xmlhttp.send();
    }
}

function allpdfs(bldg, byear, bperiod, utilid, action){
	var opts = {
	  lines: 20, // The number of lines to draw
	  length: 0, // The length of each line
	  width: 6, // The line thickness
	  radius: 9, // The radius of the inner circle
	  scale: 0.5, // Scales overall size of the spinner
	  corners: 0, // Corner roundness (0..1)
	  color: '#ff0080', // CSS color or array of colors
	  fadeColor: 'transparent', // CSS color or array of colors
	  speed: 1, // Rounds per second
	  rotate: 0, // The rotation offset
	  animation: 'spinner-line-fade-more', // The CSS animation name for the lines
	  direction: 1, // 1: clockwise, -1: counterclockwise
	  zIndex: 2e9, // The z-index (defaults to 2000000000)
	  className: 'spinner', // The CSS class to assign to the spinner
	  top: '50%', // Top position relative to parent
	  left: '50%', // Left position relative to parent
	  shadow: '0 0 1px transparent', // Box-shadow for the lines
	  position: 'absolute' // Element positioning
	};
    if (bldg.length == 0) {
        document.getElementById(bldg+".link").innerHTML = "";
        return;
    } else {
       // var xmlhttp = new XMLHttpRequest();
       // xmlhttp.onreadystatechange = function() {
       //     if (this.readyState == 4 && this.status == 200) {
       //        var target = document.getElementById(bldg+".link");
		//	   var spinner = new Spinner(opts).spin(target);
       //     }
       // };
        xmlhttp.open("GET", "consolepdf.asp?portfolio="+pid+"&byear="+byear+"&bperiod="+bperiod+"&utilfilter="+utilid+"&action="+action,true);
        xmlhttp.send();
    }
}
</script>
  <link href="spin.css" rel="stylesheet">

</head>

<body>
<pre>

</pre>
<!--[if !excel]>&nbsp;&nbsp;<![endif]-->
<!--The following information was generated by Microsoft Excel's Publish as Web
Page wizard.-->
<!--If the same item is republished from Excel, all information between the DIV
tags will be replaced.-->
<!----------------------------->
<!--START OF OUTPUT FROM EXCEL PUBLISH AS WEB PAGE WIZARD -->
<!----------------------------->

<div id="BillProcessor_31467" align=center x:publishsource="Excel">
<%
	bldg = Replace(building, "+", " ")
	bldg = Replace(bldg, "%20", " ")
if building="" then
	sql = "select top 1 bldgnum, p.portfolio, p.name from buildings b, portfolio p where portfolioid ="& pid &" and p.id = " & pid
	rst1.open sql, cnn1
	if not rst1.eof then
		bldg = rst1("bldgnum")
		pname=rst1("name")		
		p=rst1("portfolio")
	end if
	rst1.close
else	
	sql = "SELECT b.bldgnum, b.portfolioid, p.portfolio, p.name FROM buildings b, portfolio p WHERE b.portfolioid=p.id AND bldgnum='"&bldg&"'	"
	rst1.open sql, cnn1
	if not rst1.eof then 
		pid = rst1("portfolioid")
		pname=rst1("name")		
		p=rst1("portfolio")
	end if
	rst1.close 
end if
if tonumb(utilfilter)=0 then
	utilbills="All Utility Bills"
else
	sql = "SELECT DISTINCT  u.Utility as util FROM BillYrPeriod byp inner join dbo.tblutility u ON byp.Utility = u.utilityid WHERE u.utilityid="&tonumb(utilfilter)
	rst1.open sql,cnn1
	utilbills=rst1("util") & " Bills"
	rst1.close
end if
%>
<table border=0 cellpadding=0 cellspacing=0 width=689 style='border-collapse:
 collapse;table-layout:fixed;width:517pt'>
 <col width=19 style='mso-width-source:userset;mso-width-alt:694;width:14pt'>
 <col width=100 style='mso-width-source:userset;mso-width-alt:3657;width:75pt'>
 <col width=243 style='mso-width-source:userset;mso-width-alt:8886;width:182pt'>
 <col width=41 span=4 style='mso-width-source:userset;mso-width-alt:1499;
 width:31pt'>
 <col width=97 style='mso-width-source:userset;mso-width-alt:3547;width:73pt'>
 <col width=47 style='mso-width-source:userset;mso-width-alt:1718;width:35pt'>
 <col width=19 style='mso-width-source:userset;mso-width-alt:694;width:14pt'>
 <tr height=20 style='height:15.0pt'>
  <td height=20 class=xl1531467 width=19 style='height:15.0pt;width:14pt'><a
  name="RANGE!A1:J56"></a></td>
  <td class=xl1531467 width=100 style='width:75pt'></td>
  <td class=xl1531467 width=243 style='width:182pt'></td>
  <td class=xl1531467 width=41 style='width:31pt'></td>
  <td class=xl1531467 width=41 style='width:31pt'></td>
  <td class=xl1531467 width=41 style='width:31pt'></td>
  <td class=xl1531467 width=41 style='width:31pt'></td>
  <td class=xl1531467 width=97 style='width:73pt'></td>
  <td class=xl1531467 width=47 style='width:35pt'></td>
  <td class=xl1531467 width=19 style='width:14pt'></td>
 </tr>
<form name=load action="processbills.asp"> 
	 <tr height=35 style='height:26.25pt'>
	  <td height=35 class=xl1531467 style='height:26.25pt'></td>
	  <td colspan=6 class=xl6531467><%= p %> | <%= pname %></br>
		<%= utilbills %>
	  </td>
	  <td colspan=2 class=xl6531467><input size="1" class=box name="bperiod" value="<%=bperiod%>" /><input  size="1" class=box name="byear" value="<%=byear%>" /></td>
	  <td class=xl1531467></td>
	 </tr>
	 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
	  <td height=20 class=xl1531467 style='height:15.0pt'></td>
	  <td class=xl6431467></td>
	  <td class=xl6431467></td>
	  <td class=xl6431467></td>
	  <td class=xl6431467></td>
	  <td class=xl6431467></td>
	  <td class=xl6431467></td>
	  <td colspan=2 class=xl6331467><input type="submit" name="action" value="Load" class="standard" /></td>
	  <td class=xl1531467></td>
	  <INPUT type=hidden name=pid value=<%=pid%>></INPUT>
	  <INPUT type=hidden name=bldg value=<%=bldg%>></INPUT>
	 </tr>
</form>	 
 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
  <td height=20 class=xl1531467 style='height:15.0pt'></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl6331467></td>
  <td class=xl1531467></td>
 </tr>
 <form target="bill_pop" name=bills action="bills.asp" onsubmit="window.open('about:blank','bill_pop','width=300,height=500');">
	 <tr height=28 style='height:21.0pt'>
	  <td height=28 class=xl1531467 style='height:21.0pt'></td>
	  <td class=xl1531467></td>
	  <td class=xl6431467><INPUT type=submit name="action" value="Create"></td>
	  <td colspan=3 class=xl6431467><INPUT type=submit name="action" value="Delete"></td>
	  <td class=xl6431467></td>
	  <td class=xl1531467></td>
	  <td class=xl6431467></td>
	  <td class=xl1531467></td>
	 </tr>
	 <tr height=28 style='height:21.0pt'>
	  <td height=28 class=xl1531467 style='height:21.0pt'></td>
	  <td class=xl1531467></td>
	  <td class=xl6431467><INPUT type=submit name="action" value="Post"></td>
	  <td colspan=3 class=xl6431467><INPUT type=submit name="action" value="UnPost"></td>
	  <td class=xl6431467></td>
	  <td class=xl1531467></td>
	  <td class=xl6431467></td>
	  <td class=xl1531467></td>
	 </tr>
			<INPUT type=hidden name=pid value=<%=pid%>></INPUT>
			<INPUT type=hidden name=by value=<%=byear%>></INPUT>
			<INPUT type=hidden name=bp value=<%=bperiod%>></INPUT>
			<INPUT type=hidden name=utilfilter value=<%=utilfilter%>></INPUT>	
</form>

	 <tr height=28 style='height:21.0pt'>
	  <td height=28 class=xl1531467 style='height:21.0pt'></td>
	  <td class=xl1531467></td>
<form name=consolepdf action="consolepdf.asp" method="post" target="pdf_pop" onsubmit="window.open('about:blank','pdf_pop','width=10,height=10');">	  
	  <td class=xl6431467><INPUT type=submit value="Generate All" onclick="allpdfs('<%=pid%>',<%=byear%>, <%=bperiod%>,<%=utilfilter%>, 'gen');"/></td>
	 <INPUT type=hidden name=portfolio value=<%=pid%>></INPUT>
	 <INPUT type=hidden name=byear value=<%=byear%>></INPUT>
	 <INPUT type=hidden name=bperiod value=<%=bperiod%>></INPUT>
	 <INPUT type=hidden name=utilfilter value=<%=utilfilter%>></INPUT>
	 <INPUT type=hidden name=action value="gen"></INPUT>
</form>	  
	  <td colspan=3 class=xl6431467></td>
<form name=consolepdf action="consolepdf.asp" method="post" target="pdf_pop" onsubmit="window.open('about:blank','pdf_pop','width=10,height=10');">	  
	  <td class=xl6431467><INPUT type=submit value="Zip" onclick="allpdfs('<%=pid%>',<%=byear%>, <%=bperiod%>,<%=utilfilter%>, 'zip');"/>
	  </td>
	 <INPUT type=hidden name=portfolio value=<%=pid%>></INPUT>
	 <INPUT type=hidden name=byear value=<%=byear%>></INPUT>
	 <INPUT type=hidden name=bperiod value=<%=bperiod%>></INPUT>
	 <INPUT type=hidden name=utilfilter value=<%=utilfilter%>></INPUT>
	 <INPUT type=hidden name=action value="zip"></INPUT>
</form>	  	  
	  <td class=xl1531467></td>
	  <td class=xl6431467></td>
	  <td class=xl1531467></td>
	 </tr>
 <tr height=28 style='height:21.0pt'>
  <td height=28 class=xl1531467 style='height:21.0pt'></td>
  <td class=xl1531467></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl1531467></td>
</tr>
 
<tr height=28 style='height:21.0pt'>
  
  <td width="25%">Filters:</td><td>&nbsp;</td>
  <td width ="75%"><table><tr>
 			<td>
			<form name=filter action="processbills.asp" method="post"> 
				<input type="image" src="images\clearfilter.png" alt="Clear" style="width:25px;height:25px;" title="Clear Filters">&nbsp;
				<INPUT type=hidden name=pid value=<%=pid%>></INPUT>
				<INPUT type=hidden name=by value=<%=byear%>></INPUT>
				<INPUT type=hidden name=bp value=<%=bperiod%>></INPUT>
				<INPUT type=hidden name=utilfilter value=0></INPUT>
			</form>
			</td>
	<%
		sql = "SELECT DISTINCT byp.Utility as utilid, u.Utility as util FROM BillYrPeriod byp inner join dbo.tblutility u ON byp.Utility = u.utilityid WHERE BldgNum in ( select bldgnum from buildings where portfolioid="&pid&")"
		rst1.open sql,cnn1
		do until rst1.eof
			utilid = rst1("utilid")
			utilname= rst1("util")
			uicon = utilicon(utilid)
	%>		
			<td>
			<form name=filter action="processbills.asp" method="post"> 
				<input type="image" src="<%= uicon %>" alt="<%= uid %>" style="width:25px;height:25px;" title="<%= utilname %>">&nbsp;
				<INPUT type=hidden name=pid value=<%=pid%>></INPUT>
				<INPUT type=hidden name=by value=<%=byear%>></INPUT>
				<INPUT type=hidden name=bp value=<%=bperiod%>></INPUT>
				<INPUT type=hidden name=utilfilter value=<%=utilid%>></INPUT>
			</form>
			</td>
	<%
		rst1.movenext
		loop
		rst1.close
	%>

		</tr></table></td>
			
 </tr>
 
 <tr height=28 style='height:21.0pt'>
  <td height=28 class=xl1531467 style='height:21.0pt'></td>
  <td class=xl1531467></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl6431467></td>
  <td class=xl1531467></td>
 </tr>

 <tr height=20 style='height:15.0pt'>
  <td height=20 class=xl1531467 style='height:15.0pt'></td>
  <td class=xl1531467></td>
  <td class=xl1531467></td>
  <td class=xl7331467>Utility</td>
  <td class=xl7331467 align=right>C</td>
  <td class=xl7331467 align=right>P</td>
  <td class=xl7331467 align=right>T</td>
  <td class=xl1531467></td>
  <td class=xl1531467></td>
  <td class=xl1531467></td>
 </tr>
  <%
	
	Dim fso, strFileName, i, f
	Set fso = server.CreateObject("Scripting.FileSystemObject")
	dim ctime,absfile,PDFName
	dim bldg, bldgname, root, pdfdir, link, hasfile, tb, cb, pb, uid, uc, ftime, tbldgs, ibldgs, file, tc, pc, cc, util, uicon, last, utilname, newpdfname, have, oldfilefull, newfilefull, blnBillsAvailable
	ibldgs = 0
	tbldgs = 0
	tc = 0
	cc = 0
	pc = 0
	last = ""
	root = "D:\WebSites\isabella\genergyonline.com\pdfmaker\"
	dim sql 

		sql = "select upper(db.bldgnum) as bldgnum, db.bldgname from buildings db left join meters m on m.bldgnum = db.bldgnum where db.offline=0 and m.online=1 and db.portfolioid = "&pid&" group by db.bldgnum,db.bldgname order by bldgnum asc"

	rst1.open sql, cnn1
	do until rst1.eof
					
		
		bldg = rst1("bldgnum")
		bldgname = rst1("bldgname")	
		pdfdir =  p & "\" & ucase(bldg) & "\"
		
		sql = "SELECT DISTINCT byp.Utility as utilid, u.Utility as util FROM BillYrPeriod byp inner join dbo.tblutility u ON byp.Utility = u.utilityid WHERE (BldgNum = '"&bldg&"')"
		rst2.open sql, cnn1
		do until rst2.eof
			utilid = rst2("utilid")
			'response.write utilfilter & ":</br>"
			'response.write utilid
			if tonumb(utilfilter)=0 or tonumb(utilfilter) = utilid then
				if last <> bldg then
					tbldgs= tbldgs + 1
				end if				
				utilname = rst2("Util")
				uid = fw(rst2("util"))
				uicon = utilicon(utilid)
				sql = "SELECT "& _
					"(SELECT count(distinct lup.leaseutilityid) FROM tblleasesutilityprices lup, tblleases l, meters m WHERE l.billingid=lup.billingid and lup.leaseutilityid=m.leaseutilityid and m.nobill=0 and m.meternum not like '%plp%' and l.bldgnum in ('"&bldg&"') and lup.utility in ("&utilid&") and ((online=1 and l.startdate <= eomonth(DATEFROMPARTS("&byear&","&bperiod&",1))) or (online=1 and l.dateexpired > datefromparts("&byear&","&bperiod&", 1) ) ) ) as billsneeded, "&_
					"(SELECT count(*) FROM tblbillbyperiod WHERE reject=0 and leaseutilityid in (SELECT leaseutilityid FROM tblleasesutilityprices lup, tblleases l WHERE leaseexpired=0 and l.billingid=lup.billingid and l.bldgnum='"&bldg&"') and totalamt is not null and billperiod="&bperiod&" and billyear="&byear&" and utility="&utilid&") as billsprocessed, "&_
					"(SELECT count(*) FROM tblbillbyperiod WHERE reject=0 and leaseutilityid in (SELECT leaseutilityid FROM tblleasesutilityprices lup, tblleases l WHERE leaseexpired=0 and l.billingid=lup.billingid and l.bldgnum='"&bldg&"') and totalamt is not null and billperiod="&bperiod&" and billyear="&byear&" and utility="&utilid&" and posted=1) as billsposted, "&_
					"(SELECT count(*) FROM tblbillbyperiod WHERE reject=0 and totalamt is null and bldgnum='"&bldg&"' and billperiod="&bperiod&" and billyear="&byear&" and utility="&utilid&") as billserrored"
					'response.write(sql)&"</br>"
					'response.end
				rst3.open sql, cnn1
				

				if not rst3.eof then
					tb = rst3("billsneeded")
					cb = rst3("billsprocessed")
					pb = rst3("billsposted")
				end if

				rst3.close
				tc = tc + tb

				PDFName = ucase(bldg) & byear & bperiod & utilid & "1.pdf"
				ctime = year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)	
				newpdfname = ucase(bldg) &"_"& byear & "." & Right("0" & bperiod, 2) &"_"& utilname & "_TenantBills.pdf"
				oldfilefull = root&pdfdir&PDFName
				newfilefull = root&pdfdir&newPDFName

				if fso.fileexists(oldfilefull) then

					fso.copyFile oldfilefull, newfilefull, true
					fso.deletefile(oldfilefull)						
					
				else
					hasfile = false
				end if

		%>	
				 <form name=consolepdf action="consolepdf.asp" method="post" target="pdf_pop" onsubmit="window.open('about:blank','pdf_pop','width=10,height=10');">
				 
					 <tr height=20 style='height:15.0pt'>
					  <td height=20 class=xl1531467 style='height:15.0pt'></td>
					  <td rowspan=1 class=xl7031467 ><%= bldg %></td>
					  <td rowspan=1 class=xl7031467 ><%= left(bldgname,30) %></td>
					  <td class=xl7131467><img src="<%= uicon %>" alt="<%= uid %>" style="width:15px;height:15px;"></td>
					  <% if tb = 0 and cb = 0 then %>
						  <td class=xl7131467 align=right>&nbsp;</td>
						  <td class=xl7131467 align=right>&nbsp;</td>
						  <td class=xl7131467 align=right>&nbsp;</td>
						  <td class=xl7231467>&nbsp;</td>				  
					  <% else %>
						  <td class=xl7131467 align=right><%= cb %></td>
						  <td class=xl7131467 align=right><%= pb %></td>
						  <td class=xl7131467 align=right><%= tb %></td>
						  
					  <% end if %>
					<%
					if CheckRemoteURL("http://pdfmaker.genergyonline.com/pdfMaker/"&pdfdir&newpdfname) then
						have = true
						blnBillsAvailable = True
						link = pdfdir & newpdfname & "?dt=" & ctime
						set f = fso.getfile(newfilefull)
						ftime = hrsago(f.datelastmodified)
						
						hasfile = true
						ibldgs = ibldgs + 1		
						cc = cc + cb
						pc = pc + pb					
					%>
							<td class=xl7231467><INPUT type=button value="Delete" onclick="pdf('<%=bldg%>',<%=byear%>, <%=bperiod%>,<%=utilid%>, 'del');"/></td>
							<td class=xl7231467><div id="<%=bldg%>.link"><a style="font-family:arial;font-size:12;text-decoration:none;" href="http://pdfmaker.genergyonline.com/pdfMaker/<%= link %>" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'blue'" download><img src="images\pdf-icon.png" alt="<%= pdfname %> "style="width:15px;height:15px;"> | <%= ftime %></a></div></td>
					  <% else %>
							<td class=xl7231467><INPUT type=submit value="Generate" onclick="pdf('<%=bldg%>',<%=byear%>, <%=bperiod%>,<%=utilid%>, 'gen');"/></td>
							<td class=xl7231467><div id="<%=bldg%>.link">&nbsp;</div></td>
					  <% end if %>
					  <td class=xl1531467></td>
					 </tr>
							<INPUT type=hidden name=bldg value=<%=bldg%>></INPUT>
							<INPUT type=hidden name=byear value=<%=byear%>></INPUT>
							<INPUT type=hidden name=bperiod value=<%=bperiod%>></INPUT>
							<INPUT type=hidden name=utilfilter value=<%=utilid%>></INPUT>
				</form>
	<%
			end if
		rst2.movenext
		loop
		rst2.close
		last = bldg
	rst1.movenext
	loop
	rst1.close
	
%>
 <tr height=20 style='height:15.0pt'>
  <td height=20 class=xl1531467 style='height:15.0pt'></td>
  <td class=xl1531467>&nbsp;</td>
  <td class=xl1531467><%= tbldgs %> buildings</td>
  <td class=xl7331467 align=right>&nbsp;</td>
  <td class=xl7331467 align=right><%= cc %></td>
  <td class=xl7331467 align=right><%= pc %></td>
  <td class=xl7331467 align=right><%= tc %></td>
  <td class=xl1531467></td>
<%
	if tonumb(utilfilter)=0 then utilname = "All_Utilities" end if
	zipname = byear & "." & Right("0" & bperiod, 2) &"_"& utilname & "_TenantBills.zip"
	zipfull = root&p&"/"&zipname
	ziplink = p&"/"&zipname & "?dt=" & ctime
	'response.write utilname
%>
<% if CheckRemoteURL("http://pdfmaker.genergyonline.com/pdfMaker/"&p&"/"&zipname) then 
	set f = fso.getfile(zipfull)
	ftime = hrsago(f.datelastmodified)
%>
	<td class=xl1531467>
		<div id="bills.link"><a style="font-family:arial;font-size:12;text-decoration:none;" href="http://pdfmaker.genergyonline.com/pdfMaker/<%= ziplink %>" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'blue'" download><img src="images\zip.png" alt="<%= zipname %> "style="width:15px;height:15px;">&nbsp;<%= ibldgs %> invoices | <%= ftime %></a></div>
	</td>
<% else %>
	<td class=xl1531467><%= ibldgs %> invoices</td>
<% end if %>  
  
  <td class=xl1531467></td>
 </tr>

 <![if supportMisalignedColumns]>
 <tr height=0 style='display:none'>
  <td width=19 style='width:14pt'></td>
  <td width=100 style='width:75pt'></td>
  <td width=243 style='width:182pt'></td>
  <td width=41 style='width:31pt'></td>
  <td width=41 style='width:31pt'></td>
  <td width=41 style='width:31pt'></td>
  <td width=41 style='width:31pt'></td>
  <td width=97 style='width:73pt'></td>
  <td width=47 style='width:35pt'></td>
  <td width=19 style='width:14pt'></td>
 </tr>
 <![endif]>
</table>

</div>

<!----------------------------->
<!--END OF OUTPUT FROM EXCEL PUBLISH AS WEB PAGE WIZARD-->
<!----------------------------->
</body>

</html>
<% set f=nothing
set fso=nothing %>