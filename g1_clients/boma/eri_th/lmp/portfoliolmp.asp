<%option explicit
dim pid, cdate, d, cnn, rst, bldglist
pid = request("pid")
d = date()
if trim(d)="" then d = date()
cdate = date()
%>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
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

<style type=3D"text/css"><!--A {text-decoration: none}--></style>


<script>
function zoomentry(){
	var portfolioid = document.forms[0].portfolioid.value
	var tenantmeter = document.forms[0].tenantmeter.value
	var b = document.forms[0].b.value
	var m = document.forms[0].m.value
	var d = document.forms[0].d.value
	var s = document.forms[0].s.value
	var e = document.forms[0].e.value
	var temp = "zoomentry2.asp?b=" + b + "&m=" + m + "&d=" + d + "&s="+s+"&e="+e+"&luid="+l+"&lmp="+lmp+"&portfolioid="+portfolioid+"&tenantmeter="+tenantmeter
	window.open(temp,"","statusbar=0,menubar=0,scrollbars=yes,HEIGHT=125,WIDTH=300")
}
function lmpmoveprev(){
	var pid = document.forms[0].pid.value
	var b = document.forms[0].b.value
	var d = document.forms[0].pd.value
	var nd = document.forms[0].d.value
	var s = document.forms[0].s.value
	var e = document.forms[0].e.value
	var pd = new Date(d)
	pd.setTime(pd.getTime() - 1 * 24 * 60 * 60 * 1000)
	pd = (pd.getMonth()+1) + "/" + pd.getDate() + "/" + pd.getYear()

	document.forms[0].d.value = d
	document.forms[0].pd.value = pd
	document.forms[0].nd.value = nd
	
    var temp
	temp="showpidchart.asp?d="+d+"&pid="+pid
    document.frames.lmp.location=temp;
	openLoadBox('loadFrame1');
}

function lmpmovenext(){
	var pid = document.forms[0].pid.value
	var b = document.forms[0].b.value
	var d = document.forms[0].nd.value
	var pd = document.forms[0].d.value
	var s = document.forms[0].s.value
	var e = document.forms[0].e.value
	var nd = new Date(d)
	nd.setTime(nd.getTime() + 1 * 24 * 60 * 60 * 1000)
	nd = (nd.getMonth()+1) + "/" + nd.getDate() + "/" + nd.getYear()

	document.forms[0].d.value = d
	document.forms[0].pd.value = pd
	document.forms[0].nd.value = nd

    var temp
	temp="showpidchart.asp?d="+d+"&pid="+pid
	document.frames.lmp.location=temp;
	openLoadBox('loadFrame1');
}

function lmpnow(){

	var pid = document.forms[0].pid.value
	var b = document.forms[0].b.value
	var d = document.forms[0].td.value
	var s = document.forms[0].s.value
	var e = document.forms[0].e.value
	var nd = new Date(d)
	nd.setTime(nd.getTime() + 1 * 24 * 60 * 60 * 1000)
	nd = (nd.getMonth()+1) + "/" + nd.getDate() + "/" + nd.getYear()
	var pd = new Date(d)
	pd.setTime(pd.getTime() - 1 * 24 * 60 * 60 * 1000)
	pd = (pd.getMonth()+1) + "/" + pd.getDate() + "/" + pd.getYear()

	document.forms[0].d.value = d
	document.forms[0].pd.value = pd
	document.forms[0].nd.value = nd

    var temp
	temp="showpidchart.asp?d="+d+"&pid="+pid
	document.frames.lmp.location=temp;
	openLoadBox('loadFrame1');
}
var timer = null;
var mousex;
var mousey;
function hoverHelp(link)
{   timer = setTimeout("ShowHelp('"+link+"')",2000);
}
function ShowHelp(link)
{   document.all[link].style.left=mousey+10;
    document.all[link].style.top=mousex+10;
    document.all[link].style.visibility="visible";
    timer = setTimeout("HideHelp('"+link+"')",4000);
}
function HideHelp(link)
{   clearTimeout(timer);
    document.all[link].style.visibility="hidden";
}

function closeLoadBox(name)
{   document.all[name].style.visibility="hidden";
}
function openLoadBox(name)
{   var x=Math.floor(document.body.clientWidth/2-50)
    document.all[name].style.left=x
    document.all[name].style.visibility="visible";
}

function track(e)
{   mousey = event.clientX
    mousex = event.clientY
  return true

}

</script>
<%
Set cnn = Server.CreateObject("ADODB.Connection")
Set rst = Server.CreateObject("ADODB.recordset")
Set bldglist = Server.CreateObject("ADODB.recordset")

cnn.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

dim sql

sql = "SELECT bldgnum FROM buildings WHERE portfolioid = '" & pid & "' AND bldgnum IN (SELECT bldgnum FROM master.dbo.rm)"

bldglist.open sql, cnn

sql = "Select left(convert(char(20),convert(datetime,'"&d&"',101),101),11) as date"

rst.open sql, cnn
d = rst("date")
rst.close

if not bldglist.EOF then 
	while not bldglist.EOF 
	
	
	sql = "select left(convert(char(20),max(date),101),11) as date from pulse_" & bldglist("bldgnum") 
	rst.open sql,cnn
	
	if not rst.EOF then
		if d > rst("date") then
			d = rst("date")
		end if
	end if
	
	bldglist.movenext
	rst.close
	wend

end if
bldglist.close

%>
<body bgcolor="#FFFFFF" text="#000000" link="#FFFFFF" vlink="#FFFFFF" alink="#0099FF" onmousemove="track(this)">
<form name="form1" method="post" action="">
             <input type="hidden" name="b" value="">
             <input type="hidden" name="d" value="<%=d%>">
              <input type="hidden" name="pd" value="<%=DateAdd("d",-1,d)%>">
              <input type="hidden" name="nd" value="<%=DateAdd("d",1,d)%>">
              <input type="hidden" name="td" value="<%=Date()%>">
              <input type="hidden" name="s" value="">
              <input type="hidden" name="e" value="">
              <input type="hidden" name="pid" value="<%=pid%>">

<table width="710" border="1" cellspacing="0" cellpadding="0" align="center">
<tr><td width="687" bgcolor="#000000"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000">
        <tr><td width="100%" height="2" bgcolor="#000000">
                
              <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr>
                  <td><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="3"><b>Portfolio 
                    Load Profile</b></font></td>
                  <td align="right">&nbsp;</td>
                </tr>
                <tr>
                  <td><b><font size="2" face="Arial, Helvetica, sans-serif" color="#FFFFFF"> 
                    <!--<a href="javascript:zoomentry()" style="text-decoration:none;" onMouseOver="this.style.color='lightblue';hoverHelp('zoomhelp')" onMouseOut="this.style.color='white';HideHelp('zoomhelp')">Interval Zoom</a> | -->
                    <a href="javascript:lmpmoveprev()" style="text-decoration:none;" onMouseOver="this.style.color='lightblue';hoverHelp('prevhelp')" onMouseOut="this.style.color='white';HideHelp('prevhelp')">Previous 
                    Day</a> | <a href="javascript:lmpmovenext()" style="text-decoration:none;" onMouseOver="this.style.color='lightblue';hoverHelp('nexthelp')" onMouseOut="this.style.color='white';HideHelp('nexthelp')">Next 
                    Day</a> | <a href="javascript:lmpnow()" style="text-decoration:none;" onMouseOver="this.style.color='lightblue';hoverHelp('tdayhelp')" onMouseOut="this.style.color= 'white';HideHelp('tdayhelp')">Go 
                    To Today</a> </font></b> </td>
                  <td align="right"><b><font size="2" face="Arial, Helvetica, sans-serif" color="#FFFFFF"><a href="/g1_clients/manual/lmp.htm" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Help</a>&nbsp;</font></b></td>
                </tr>
              </table>
            </td>
        </tr></table>
    </td>
</tr></table>
<table width="710" border="1" cellspacing="0" cellpadding="0" height="277" align="center">
<tr><td width="687" height="330">
        <div align="center"><iframe name="lmp" id="lmp" style="" width="100%" height="100%" src="showpidchart.asp?d=<%=d%>&pid=<%=pid%>" scrolling="auto" marginwidth="0" marginheight="0" ></iframe></div>
      </td></tr>
</table>



</form>


<div id="zoomhelp" style="visibility:hidden; position:absolute;left:17;top:35;background-color:lightyellow;font-family:arial;font-size:10px;border-width:1px;border-style:solid">
The Zoom function
</div>
<div id="prevhelp" style="visibility:hidden; position:absolute;left:112;top:35;background-color:lightyellow;font-family:arial;font-size:10px;border-width:1px;border-style:solid">
Go to previous day
</div>
<div id="nexthelp" style="visibility:hidden; position:absolute;left:180;top:35;background-color:lightyellow;font-family:arial;font-size:10px;border-width:1px;border-style:solid">
Go to next day
</div>
<div id="tdayhelp" style="visibility:hidden; position:absolute;left:220;top:35;background-color:lightyellow;font-family:arial;font-size:10px;border-width:1px;border-style:solid">
Go to today
</div>

<div id="loadFrame1" style="visibility:visible; position:absolute;left:320;top:150;background-color:lightyellow;border-width:1px;border-style:solid">
<table border="0" cellpadding="5" cellspacing="0"><tr><td style="font-family:arial;font-size:16px;font-weight:bold">Loading Data...</td></tr></table>
</div>
<font face="Arial, Helvetica, sans-serif" size="2"><b><i>NOTE: Default profile 
shown is the most recent data set</i></b></font> 
</body>
<script>
//openLoadBox('loadFrame1');
</script>

</html>
