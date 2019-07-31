<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Set cnn1 	= Server.CreateObject("ADODB.Connection")
Set rs 		= Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"dbCore")
servername = request("servername")
openwindow = request("openwindow")
icon = request("icon")
if icon = "" then icon = false 
%>
<html>
<head>
<title>Simple Server Monitor</title>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<style type="text/css">
.tblunderline { border-bottom:1px solid #cccccc; }
.tbltopline { border-top:1px solid #cccccc;border-bottom:1px solid #cccccc; }
</style>
<script>
function popUp(page, windowsizew, windowsizeh,scrollstat,id){
	var w = windowsizew;
	var h = windowsizeh;
	var winl = (screen.width - w) / 2;
    var wint = (screen.height - h) / 2;

    winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',status=yes,scrollbars='+scrollstat+',resizable=yes'
     // open new window and use the variables to position it
	//popupwin=window.open(page,'login','WIDTH=400, HEIGHT=300, scrollbars=no,left='+x+',top='+y)
	popupwin=window.open(page,id,winprops)
	popupwin.focus(id)
}
<!--
//enter refresh time in "minutes:seconds" Minutes should range from 0 to inifinity. Seconds should range from 0 to 59
var limit=":59"

if (document.images){
var parselimit=limit.split(":")
parselimit=parselimit[0]*60+parselimit[1]*1
}
function beginrefresh(){
if (!document.images)
return
if (parselimit==1)
window.location.reload()
else{ 
parselimit-=1
curmin=Math.floor(parselimit/60)
cursec=parselimit%60
if (curmin!=0)
curtime=curmin+" minutes and "+cursec+" seconds left until page refresh! Actual Monitors refresh ewery 5 minutes."
else
curtime=cursec+" seconds left until page refresh! Actual Monitors refresh ewery 5 minutes."
window.status=curtime
setTimeout("beginrefresh()",1000)
}
}
window.onload=beginrefresh
//-->
</script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
<head>
<body bgcolor=<%if icon then %>"#FFFFFF"<%else%>"#eeeeee"<%end if%> leftmargin="0" topmargin="0" class="innerbody" onunload="window.status='gEnergyOne I:2 Intranet'">
  
<table width="100%" border=0 cellpadding="3" cellspacing="0" bgcolor="#ffffff">
  <tr align="left" valign="bottom" bgcolor="#dddddd"> 
    <td valign="top" nowrap bgcolor="#336699" class="tblunderline"><span class="standardheader"><font size="2">Simple 
      Server Monitor<% if servername <> "" then %>, history for <%=servername%><%end if%></font></span></td>
    <td align="right" nowrap bgcolor="#336699" class="tblunderline"><table width="163" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="20" bgcolor=<%if icon then%>"#FFFFFF"<%else%>"#FFB000"<%end if%>> 
            <%if icon then%>
            <img src="images/systemcritical.jpg" width="25" height="25"> 
            <%end if%></td>
          <td align="center"><span class="standardheader">&nbsp;Critical Levels 
            Reached</span></td>
        </tr>
      </table></td>
  </tr>
  <tr align="left" valign="bottom" bgcolor="#dddddd"> 
    <td nowrap bgcolor="#FFFFCC" class="tblunderline"> Server Status <%if openwindow = "" then %>| <a href="#" onclick="popUp('overview.asp?openwindow=yes',1024,768,'no','SSM')">Open Window</a><%end if%><%if not icon then %> | <a href="overview.asp?icon=true&openwindow=<%=openwindow%>">Icon View</a><%else%> | <a href="overview.asp?icon=false&openwindow=<%=openwindow%>">Detail View</a><%end if%></td>
    <td align="right" nowrap bgcolor="#FFFFCC" class="tblunderline">&nbsp;</td>
  </tr>
</table>
<table width="100%" border=0 cellpadding="3" cellspacing="0" bgcolor="#f0f0e0">
<%
if trim(servername) ="" then 
	strsql = "select ns.date, ns.servername, isnull(cpulevel,0) as currentCPULevel, isnull(a.avgcpu,0) as avgcpu,isnull(a.maxcpu,0) as maxcpu, isnull(a.mincpu,0) as mincpu, isnull(totalram,0) as totalram, isnull(ns.freeram,0) as currentFreeRam, isnull(b.avgfreeram,0) as AvgFreeRam, isnull(b.maxFreeRam,0) as maxFreeRam,isnull(b.minFreeRam,0) as minFreeRam, description, isnull(c.totspace,0) as totspace, isnull(c.availspace,0) as availspace from netservers ns inner join (select avg(cpulevel) avgcpu, max(cpulevel) as maxcpu, min(cpulevel) as mincpu, servername from netservers group by servername) a on a.servername = ns.servername inner join (select avg(freeram) avgfreeRam, max(freeram) as maxFreeRam, min(freeram) as minFreeRam, servername from netservers group by servername) b on b.servername = ns.servername inner join (select servername,date, sum(totalspace) as totspace, sum(freespace) as availspace from ns_drivetracking group by servername, date) c on c.servername = ns.servername and c.date = ns.date inner join (select servername, max(date) as maxdate from netservers group by servername) d on d.servername = ns.servername and ns.date = d.maxdate order by  ns.servername"
	serverlink = true
else 
	strsql = "SELECT top 50000 ns.date, ns.servername, isnull(ns.cpulevel,0) as cpulevel, isnull(ns.freeram,0) as freeram, isnull(ns.totalram,0) as totalram,  isnull(c.totspace,0) as totspace, isnull(c.availspace,0) as availspace FROM  netservers ns inner join (select servername,date, totalspace as totspace,  freespace as availspace from ns_drivetracking) c on c.servername = ns.servername and c.date = ns.date WHERE  (ns.servername = '"&trim(servername)&"') ORDER BY ns.[date] desc"
	serverlink = false
end if 

rs.Open strsql, cnn1
	
if not icon then 

if trim(servername) = "" then 
%>
  <tr> 
    <td align="left" valign="middle" nowrap class="tblunderline"><strong>Server</strong></td>
    <td align="center" valign="middle" nowrap  class="tblunderline"><strong>Snapshot 
      Date</strong></td>
    <td align="left" valign="middle" nowrap bgcolor="#99FFCC"  class="tblunderline">&nbsp;</td>
    <td colspan=2 align="center" valign="middle" nowrap bgcolor="#99FFCC"  class="tblunderline"><strong>CPU 
      Levels </strong></td>
    <td align="left" valign="middle" nowrap bgcolor="#99FFCC"  class="tblunderline">&nbsp;</td>
    <td align="center" valign="center" nowrap bgcolor="#CCFFCC"  class="tblunderline" colspan=2><strong>Storage</strong></td>
    <td align="center" valign="middle" nowrap bgcolor="#99CC99"   class="tblunderline"><strong>Total 
      Memory</strong></td>
    <td align="left" valign="middle" nowrap bgcolor="#99CC99"  class="tblunderline">&nbsp;</td>
    <td colspan = 2 align="center" valign="middle" nowrap bgcolor="#99CC99"  class="tblunderline"><strong>Available 
      Memory </strong></td>
    <td align="left" valign="middle" nowrap bgcolor="#99CC99"  class="tblunderline">&nbsp;</td>
  </tr>
  <tr> 
    <td align="left" valign="middle" nowrap class="tblunderline">&nbsp;</td>
    <td align="left" valign="middle" nowrap  class="tblunderline">&nbsp;</td>
    <td align="left" valign="middle" nowrap bgcolor="#99FFCC"  class="tblunderline">Current 
    </td>
    <td align="left" valign="middle" nowrap bgcolor="#99FFCC"  class="tblunderline">Average 
    </td>
    <td align="left" valign="middle" nowrap bgcolor="#99FFCC"  class="tblunderline">Max 
    </td>
    <td align="left" valign="middle" nowrap bgcolor="#99FFCC"  class="tblunderline">Min 
    </td>
    <td align="center" valign="middle" nowrap bgcolor="#CCFFCC" class="tblunderline">Total&nbsp;</td>
    <td align="center" valign="middle" nowrap bgcolor="#CCFFCC" class="tblunderline">Available&nbsp;</td>
    <td align="left" valign="middle" nowrap bgcolor="#99CC99" class="tblunderline">&nbsp;</td>
    <td align="center" valign="middle" nowrap bgcolor="#99CC99"  class="tblunderline">Current 
    </td>
    <td align="center" valign="middle" nowrap bgcolor="#99CC99"  class="tblunderline">Average 
    </td>
    <td align="center" valign="middle" nowrap bgcolor="#99CC99"  class="tblunderline">Max 
    </td>
    <td align="center" valign="middle" nowrap bgcolor="#99CC99"  class="tblunderline">Min 
    </td>
  </tr>
<%

else

%>
  <tr> 
    <td align="center" valign="middle" nowrap  class="tblunderline"><strong>Snapshot 
      Date</strong></td>
    <td align="left" valign="middle" nowrap bgcolor="#99FFCC"  class="tblunderline">&nbsp;</td>
    <td colspan=2 align="center" valign="middle" nowrap bgcolor="#99FFCC"  class="tblunderline"><strong>CPU 
      Level</strong></td>
    <td align="left" valign="middle" nowrap bgcolor="#99FFCC"  class="tblunderline">&nbsp;</td>
    <td align="center" valign="center" nowrap bgcolor="#CCFFCC"  class="tblunderline" colspan=2><strong>Storage</strong></td>
    <td align="center" valign="middle" nowrap bgcolor="#99CC99"   class="tblunderline"><strong>Total 
      Memory</strong></td>
    <td align="left" valign="middle" nowrap bgcolor="#99CC99"  class="tblunderline">&nbsp;</td>
    <td colspan = 2 align="center" valign="middle" nowrap bgcolor="#99CC99"  class="tblunderline"><strong>Available 
      Memory </strong></td>
    <td align="left" valign="middle" nowrap bgcolor="#99CC99"  class="tblunderline">&nbsp;</td>
  </tr>
 <%

 end if 
 
  else
  %>
  <br><br>
  <div align="center">
  <%
  end if
  
	if not rs.eof then
	Dim CPU(3),RAM(3)
	iconcount = 0 
   Do until rs.EOF 
		if trim(servername) = "" then    
			sname  = ucase(rs("servername"))
			tDate		= rs("date")
			cpu(0) 	= cdbl(rs("currentCPULevel"))
			cpu(1)  = cdbl(rs("avgcpu"))
			cpu(2)	= cdbl(rs("maxcpu"))
			cpu(3) 	= cdbl(rs("mincpu"))
			totalram 	= cdbl(rs("totalram"))
			
			if totalram <> 0 then 
				RAM(0) 	= cdbl(rs("currentFreeRam"))/totalram
				RAM(1) 	= cdbl(rs("AvgFreeRam"))/totalram
				RAM(2) 	= cdbl(rs("maxFreeRam"))/totalram
				RAM(3) 	= cdbl(rs("minFreeRam"))/totalram
			else
				RAM(0) 	= 0
				RAM(1) 	= 0
				RAM(2) 	= 0
				RAM(3) 	= 0
			end if 
			
			totalSpace = rs("totspace")
			availSpace = rs("availspace")
			
			if totalspace <> 0 then 
				spacediff 		= availspace / totalspace
			else
				spacediff 		= 0
			end if 
			
			if  ram(0) < .20 or cpu(0) > .80 or spacediff < .25 then 
				critical = true
			else
				critical = false
			end if 
		else
			tDate			= rs("date")
			cpu(0) 			= cdbl(rs("cpulevel"))
			totalram 		= cdbl(rs("totalram"))
			if totalram <> 0 then 
				RAM(0) 			= cdbl(rs("FreeRam"))/totalram
			else
				RAM(0)			= 0
			end if
			
			totalSpace 		= rs("totspace")
			availSpace 		= rs("availspace")
			
			if totalspace <> 0 then 
				spacediff 		= availspace / totalspace
			else
				spacediff 		= 0
			end if 
			
			if  ram(0) < .20 or cpu(0) > .80 or spacediff < .25 then 
				critical = true
			else
				critical = false
			end if 
		end if 
		
  if icon then 'Icon View or Not
	  if iconcount = 10 then 
	  	%><br><br>
		<%	  
	  end if
			if critical then 
				imagetxt = "systemcritical.jpg"
			else
				imagetxt = "system.jpg"
			end if
			%><div style="width:80;height:10;display:inline"><img src="images/<%=imagetxt%>" alt="<%=servername%>" title="<%=servername%>"><br><%=servername%></div>&nbsp;&nbsp;
			<%
			iconcount = iconcount + 1				
  else
	if servername = "" then 
	  if critical then 	
		%>
	  <tr style="background-color:#FFB000;"> 
		<%else%>
	  <tr> 
		<%end if%>
		<td align="left" valign="middle" nowrap class="tblunderline"> <a href="#" onclick="popUp('overview.asp?servername=<%=sname%>',800,400,'yes','historical')"><%=sname%></a> 
		</td>
		<td align="left" valign="middle" nowrap  class="tblunderline"><%=tdate%></td>
		<%
		x = 0 
		Do until x > 3 
		if cpu(x) > .70 and x <> 2 then fillcolor = "#FF0000" else fillcolor = "#0033FF" end if 
		%>
		<td align="left" valign="middle" nowrap <% if not critical then %>bgcolor="#99FFCC"<%end if%> class="tblunderline"><%=formatpercent(cdbl(cpu(x)))%> 
		  <table width="<%=formatpercent(cdbl(cpu(x)))%>" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" bgcolor="<%=fillcolor%>">
			<tr> 
			  <td>&nbsp;</td>
			</tr>
		  </table></td>
		<%
		x=x+1
		loop
		if spacediff < .25 then fillcolor = "#FF0000" else fillcolor = "#0033FF" end if 
	%>
		<td align="right" valign="middle" nowrap <% if not critical then %>bgcolor="#CCFFCC"<%end if%>   class="tblunderline"><%=formatnumber(totalSpace/1024,2)%> GB&nbsp;&nbsp;&nbsp;</td>
		<td align="left" nowrap <% if not critical then %>bgcolor="#CCFFCC"<% end if %>  class="tblunderline"><%=formatnumber(availspace/1024,2)%> GB 
		  <table width="<%=formatpercent(spacediff)%>" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" bgcolor="<%=fillcolor%>">
			<tr> 
			  <td>&nbsp;</td>
			</tr>
		  </table></td>
		<td align="right" valign="middle" nowrap <% if not critical then %>bgcolor="#99CC99"<%end if%>   class="tblunderline"><%=formatnumber(totalram/1024,0)%> MB&nbsp;&nbsp;&nbsp;</td>
		<%
		x = 0 
		Do until x > 3 
		select case x 
		case 0
			if ram(x) < .20 then fillcolor = "#FF0000" else fillcolor = "#660000" end if 
		case 1 
			if ram(x) < .20 then fillcolor = "#FF0000" else fillcolor = "#660000" end if 
		case 2,3
			if (ram(2)-ram(3) < .20) and ram(2) < .30 then fillcolor = "#FF0000" else fillcolor = "#660000" end if 
		case else
			fillcolor = "#660000"
		end select 
	
		%>
		<td align="left" nowrap <% if not critical then %>bgcolor="#99CC99"<% end if %>  class="tblunderline"><%=formatpercent(ram(x))%> 
		  <table width="<%=formatpercent(ram(x))%>" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" bgcolor="<%=fillcolor%>">
			<tr> 
			  <td>&nbsp;</td>
			</tr>
		  </table></td>
		<%
		x=x+1  
		loop
	%>
	  </tr>
	  <%
	else
	  if critical then 	
		%>
	  <tr style="background-color:#FFB000;"> 
		<%else%>
	  <tr> 
		<%end if%>
		<td align="left" valign="middle" nowrap  class="tblunderline"><%=tdate%></td>
		<%
		x = 0 
		Do until x > 3 
		if cpu(x) > .70 and x <> 2 then fillcolor = "#FF0000" else fillcolor = "#0033FF" end if 
		%>
		<td align="left" valign="middle" nowrap <% if not critical then %>bgcolor="#99FFCC"<%end if%> class="tblunderline"><%=formatpercent(cdbl(cpu(x)))%> 
		  <table width="<%=formatpercent(cdbl(cpu(x)))%>" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" bgcolor="<%=fillcolor%>">
			<tr> 
			  <td>&nbsp;</td>
			</tr>
		  </table></td>
		<%
		x=x+1
		loop
		if spacediff < .25 then fillcolor = "#FF0000" else fillcolor = "#0033FF" end if 
	%>
		<td align="right" valign="middle" nowrap <% if not critical then %>bgcolor="#CCFFCC"<%end if%>   class="tblunderline"><%=formatnumber(totalSpace/1024,2)%> GB&nbsp;&nbsp;&nbsp;</td>
		<td align="left" nowrap <% if not critical then %>bgcolor="#CCFFCC"<% end if %>  class="tblunderline"><%=formatnumber(availspace/1024,2)%> GB 
		  <table width="<%=formatpercent(spacediff)%>" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" bgcolor="<%=fillcolor%>">
			<tr> 
			  <td>&nbsp;</td>
			</tr>
		  </table></td>
		<td align="right" valign="middle" nowrap <% if not critical then %>bgcolor="#99CC99"<%end if%>   class="tblunderline"><%=formatnumber(totalram/1024,0)%> MB&nbsp;&nbsp;&nbsp;</td>
		<%
		x = 0 
		Do until x > 3 
		select case x 
		case 0
			if ram(x) < .20 then fillcolor = "#FF0000" else fillcolor = "#660000" end if 
		case 1 
			if ram(x) < .20 then fillcolor = "#FF0000" else fillcolor = "#660000" end if 
		case 2,3
			if (ram(2)-ram(3) < .20) and ram(2) < .30 then fillcolor = "#FF0000" else fillcolor = "#660000" end if 
		case else
			fillcolor = "#660000"
		end select 
	
		%>
		<td align="left" nowrap <% if not critical then %>bgcolor="#99CC99"<% end if %>  class="tblunderline"><%=formatpercent(ram(x))%> 
		  <table width="<%=formatpercent(ram(x))%>" border="1" cellpadding="0" cellspacing="0" bordercolor="#000000" bgcolor="<%=fillcolor%>">
			<tr> 
			  <td>&nbsp;</td>
			</tr>
		  </table></td>
		<%
		x=x+1  
		loop
	%>
	  </tr>
	  <%
	end if 'Server Display
  	end if 'Icon View or not  
  	rs.movenext
    loop
else
	%>
  <tr> 
    <td colspan="5">NO DATA FOUND</td>
  </tr>
  <%
end if
rs.close
%>
<%if icon then%></div><%end if%>
</table>
</body>
</html>
