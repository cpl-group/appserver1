<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, tid, lid, meterid
pid = secureRequest("pid")
bldg = secureRequest("bldg")
tid = secureRequest("tid")
lid = secureRequest("lid")
meterid = secureRequest("meterid")
'dim DBmainmodIP
'DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

dim Meternum, StartDate, DateOff, DateRead, TimeRead, Factor, KWHx, KWx, Location, Riser, online, note, Cumulative, variance, lmp, floor, datasource, meterrefnum, powerfeed
if trim(meterid)<>"" then
	rst1.Open "SELECT * FROM meters LEFT JOIN (SELECT meterid as meterrefid, meternum as meterrefnum FROM meters) ref ON ref.meterrefid=refmeterid LEFT JOIN (SELECT meterid as meterrefid, meternum as powmeternum FROM meters) pow ON pow.meterrefid=powerfeed WHERE meterid='"&meterid&"'", cnn1
	if not rst1.EOF then
		Meternum = rst1("Meternum")
		StartDate = rst1("DateStart")
		DateOff = rst1("DateOffLine")
		DateRead = rst1("DateLastRead")
		TimeRead = rst1("TimeLastRead")
		Factor = rst1("multiplier")
		KWHx = rst1("manualmultiplier")
		KWx = rst1("Demandmultiplier")
		Location = rst1("Location")
		Riser = rst1("Riser")
		online = rst1("online")
		note = rst1("metercomments")
		Cumulative = rst1("Cumulative")
		variance = rst1("variance")
		lmp = rst1("lmp")
		floor = rst1("floor")
		datasource = rst1("datasource")
		meterrefnum = rst1("meterrefnum")
		powerfeed = rst1("powmeternum")
		if trim(meterrefnum)="" then meterrefnum = "N/A"
		if trim(powerfeed)="" then powerfeed = "N/A"
	end if
	rst1.close
end if
%>
<html>
<head>
<title>Building View</title>
<script>
function meterEdit(meterid)
{	document.location.href = 'meteredit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&meterid='+meterid
}
function reloadmeterfrm(){
  if (parent.name == "contentfrm") {
    parent.meterfrm.location.reload();
  }
}
</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body bgcolor="#ffffff" onload="reloadmeterfrm();">
<form name="form2" method="post" action="metersave.asp">

<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#3399cc">
	<td colspan="5">
	<span class="standardheader">
			Update Successful
	</span></td>
</tr>
<tr bgcolor="#eeeeee">
  <td colspan="5" bgcolor="#dddddd" style="border-bottom:1px solid #ffffff;">
  <table border=0 cellpadding="3" cellspacing="1">
  <tr>
    <td align="right"><span class="standard">Meter</span></td> 
    <td><span class="standard"><%=Meternum%> &nbsp;&nbsp;On Line: <%if online="True" then Response.Write "Yes" else Response.Write "No"%></span></td>
  </tr>
  </table>
  </td>
</tr>
<tr valign="top">
  <td bgcolor="#eeeeee" width="30%" style="border-bottom:1px solid #cccccc;"> 
  <table border="0" cellpadding="3" cellspacing="1">
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Start Date</span></td>
    <td><span class="standard"><%=StartDate%></span></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Date Off</span></td>
    <td><span class="standard"><%=DateOff%></span></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Date Read</span></td>
    <td><span class="standard"><%=DateRead%></span></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Time Read</span></td>
    <td><span class="standard"><%=TimeRead%></span></td>
  </tr>
  </table>
  </td>
  <td bgcolor="#eeeeee" width="30" style="border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;">&nbsp;</td>
  <td width="30%" bgcolor="#eeeeee" style="border-left:1px solid #ffffff;border-right:1px solid #cccccc;border-bottom:1px solid #cccccc;">
  <table border="0" cellpadding="3" cellspacing="1">
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Factor</span></td>
    <td><span class="standard"><%=Factor%></span></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">KWH.x</span></td>
    <td><span class="standard"><%=KWHx%></span></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">KW.x</span></td>
    <td><span class="standard"><%=KWx%></span></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Variance</span></td>
    <td><span class="standard"><%=variance%></span></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Meter Reference</span></td>
    <td><span class="standard"><%=meterrefnum%></span></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Power Feed</span></td>
    <td><span class="standard"><%=powerfeed%></span></td>
  </tr>
  </table>
  </td>
  <td bgcolor="#eeeeee" width="30" style="border-bottom:1px solid #cccccc;\>&nbsp;</td>
  <td bgcolor="#eeeeee" style="border-left:1px solid #ffffff;">
  <table border="0" cellpadding="3" cellspacing="1">
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Location</span></td> 
    <td><span class="standard"><%=Location%></span></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Floor</span></td>
    <td><span class="standard"><%=floor%></span></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Riser</span></td> 
    <td><span class="standard"><%=Riser%></span></td>
  </tr>
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">datasource</span></td>
    <td><span class="standard"><%=datasource%></span></td>
  </tr>
  </table>
  </td>
</tr>
<tr bgcolor="#eeeeee">
  <td colspan="5" bgcolor="#eeeeee" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
  <table border=0 cellpadding="3" cellspacing="0">
  <tr bgcolor="#eeeeee">
    <td align="right"><span class="standard">Note</span></td>
    <td><span class="standard"><%=note%></span></td>
    <td><span class="standard">Cumulative: <%if Cumulative="True" then Response.Write "Yes" else Response.Write "No"%></span> &nbsp;&nbsp;&nbsp;<span class="standard">lmp: <%if lmp="True" then Response.Write "Yes" else Response.Write "No"%></span></td>
  </tr>
  </table>
  </td>
</tr>
</table>
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="tid" value="<%=tid%>">
<input type="hidden" name="lid" value="<%=lid%>">
<input type="hidden" name="meterid" value="<%=meterid%>">
</form>
<%if allowGroups("clientOperations") then%>
<script>
//window.parent.document.location='<%="tenantedit.asp?pid="&pid&"&bldg="&bldg&"&tid="&tid%>'
</script>
<%end if%>
</body>
</html>






