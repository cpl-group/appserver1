<%option explicit

dim m, d, b, s, e, i, luid, lmp, tenantmeter
m=Request.QueryString("m")
d=Request.QueryString("d")
b=Request.QueryString("b")
s=Request.QueryString("s")
e=Request.QueryString("e")
i=Request.QueryString("i")
luid=Request.QueryString("luid")
tenantmeter=Request.QueryString("tenantmeter")

if instr(luid,"_") then
    tenantname = right(luid, len(luid)-instr(luid,"_"))
    luid = left(luid, instr(luid,"_")-1)
end if

lmp=request.querystring("lmp")
if isempty(i) then 
	i=100
end if

dim cnn1, rst1, strsql
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_genergy1")

if d="" then 
	if luid = "" then 
		'strsql = "select max(date) as date from pulse_"&b&" where meterid = " & m
		strsql = "select top 1 [date] as date from pulse_"&b&" where meterid="&m&" order by [date] desc"
		rst1.Open strsql, cnn1
		if not rst1.eof then 
			d=rst1("date")
			d=left(d&" ",instr(d&" ", " ")-1)
		end if
		rst1.close
	else
		strsql = "select left(convert(char(20),max(date),101),11) as date from pulse_"&b&" where meterid in (select meterid from meters where leaseutilityid = "& luid &")" 
		rst1.Open strsql, cnn1
		if not rst1.eof then 
			d=rst1("date")		
		end if
		rst1.close
		'response.write strsql
		'response.end
	end if
end if
if s="" then
	s=0
end if
if e="" then 
	e=2400
end if 

if not(isdate(d)) then d=date()


%>
<html>
<head>
<title>LMP Chart</title>
<script>
function propagateVarsUP()
{   var myparent = parent.document.location+""
	if(myparent.indexOf('lmpload2.asp?')==-1)
	{	//alert('propagate');
		parent.document.forms[0].luid.value="<%=luid%>"
	    parent.document.forms[0].b.value="<%=b%>"
	    parent.document.forms[0].m.value="<%=m%>"
	    parent.document.forms[0].d.value="<%=d%>"
	    parent.document.forms[0].nd.value="<%=DateAdd("d",+1,d)%>"
	    parent.document.forms[0].pd.value="<%=DateAdd("d",-1,d)%>"
	    parent.document.forms[0].s.value="<%=s%>"
	    parent.document.forms[0].e.value="<%=e%>"
	    parent.document.forms[0].luid.value="<%=luid%>"
	    parent.document.forms[0].tenantmeter.value="<%=tenantmeter%>"
	    parent.document.forms[0].lmp.value="<%=lmp%>"
		parent.closeLoadBox('loadFrame1');
		parent.shownav('lmpnav')
    	//alert(parent.document.forms[0].d.value)
	}
}
</script>
</head><style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" onload="propagateVarsUP();">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%" align="center">
  <tr>
    <td>
      <div align="center"><img src="<%="makechartlmp2.asp?m="& m & "&d=" & d & "&s=" & s & "&e="&e&"&b="&b&"&i="&i&"&luid="&luid&"&lmp="&lmp&"&tenantmeter="& server.urlencode(tenantmeter)%>"></div>
    </td>
  </tr>
</table>
<%
'response.redirect "makechartlmp2.asp?m="& m & "&d=" & d & "&s=" & s & "&e="&e&"&b="&b&"&i="&i&"&luid="&luid&"&lmp="&lmp&"&tenantmeter="& server.urlencode(tenantmeter)
%>
</body>
</html>
