<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldg, date1, pid, utype
utype = request("utype")
bldg = request("bldg")
if trim(bldg)="" then bldg = request("b")
pid = request("pid")
date1 = year(date())

dim rst1, cnn1
set rst1 = server.createobject("ADODB.recordset")
set cnn1 = server.createobject("ADODB.connection")
if trim(bldg)<>"" then cnn1.open getLocalConnect(bldg) else cnn1.open getMainConnect(pid)
if trim(pid)="" then
  rst1.open "SELECT portfolioid FROM buildings WHERE bldgnum='"&bldg&"'", cnn1
  if not rst1.eof then pid = cint(rst1("portfolioid"))
  rst1.close
end if

session("pid") = pid
if session("Expenses")="" then
	session("Expenses")=1
	session("Expense_Adjustments")=1
	session("Submeter")=1
	session("ERI")=1
	session("Unreported_Revenue_Adjustments")=1
	session("Mac_Revenue")=0
	session("PLP_Revenue")=0
	session("Net")=1
end if

dim IFrame1, IFrame2
if trim(utype)="" then utype = "2"
IFrame1 = "revChartLoad.asp?bldg="& bldg &"&pid="& pid &"&utype="&utype&"&date1="& date1
IFrame2 = "options.asp?bldg="& bldg &"&pid="& pid &"&utype="&utype&"&date1="& date1
%>
<HTML>

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

<script>
var storedPreferenceHottab //this var is intended to remember the hot tab when preferences has been clicked--for the posiibility of pressing 'next', 'prev' or 'today' while in preferences

function moveprev()
{	var date1 = document.forms['form1'].date1.value;
	var date2 = document.forms['form1'].date2.value;
	var utype = document.forms['form1'].utype.value;
	if (date1 >= 2001){
		date1--;
		document.forms['form1'].date1.value = date1
		if(date2!="")
		{	date2--;
			document.forms['form1'].date2.value = date2;
		}
		loadchart();
		settabs(storedPreferenceHottab);//makes sure that the proper utility tab is on instead of the preferences tab
	} else{
		alert("Data not available prior to 2000");	
	}
}

function movenext()
{	var date1 = document.forms['form1'].date1.value;
	var date2 = document.forms['form1'].date2.value;
	var utype = document.forms['form1'].utype.value;
	date1++;
	document.forms['form1'].date1.value = date1
	if(date2!="")
	{	date2++;
		document.forms['form1'].date2.value = date2;
	}
	loadchart();
	settabs(storedPreferenceHottab);//makes sure that the proper utility tab is on instead of the preferences tab
}

function movenow()
{	
	var date1 = document.forms['form1'].date1.value;
	var date2 = document.forms['form1'].date2.value;
	var utype = document.forms['form1'].utype.value;
	date1=<%=date1%>;
	document.forms['form1'].date1.value = date1
	if(date2!="")
	{	date2=date1-1;
		document.forms['form1'].date2.value = date2;
	}
	loadchart();
	settabs(storedPreferenceHottab);//makes sure that the proper utility tab is on instead of the preferences tab
}

function settabs(hottab)
{	for(i=2;i<=3;i++)
	{	document.all['tab'+i].style.backgroundColor='#CCCCCC'
	}
	hottab.style.backgroundColor='#0099FF'
}

function loadchart()
{	var date1 = document.forms['form1'].date1.value;
	var date2 = document.forms['form1'].date2.value;
	var utype = document.forms['form1'].utype.value;
	document.frames['chart'].document.location.href = 'revChartLoad.asp?bldg=<%=bldg%>&pid=<%=pid%>&utype='+utype+'&date1='+date1+'&date2='+date2;
	openLoadBox('loadFrame1');
	var frame2Href = document.frames['options'].document.location.href
	frame2Href = frame2Href.toLowerCase()
	if((frame2Href.indexOf("monthlydetails.asp")!=-1)||(frame2Href.indexOf("breakdownexpense.asp")!=-1)||(frame2Href.indexOf("breakdownsub_eri.asp")!=-1)||(frame2Href.indexOf("breakdownadjustment.asp")!=-1)||(frame2Href.indexOf("BreakdownPLP_mac.asp")!=-1))
	{	document.frames['options'].document.location.href = 'monthlyDetails.asp?bldg=<%=bldg%>&pid=<%=pid%>&utype='+utype+'&date1='+date1+'&date2='+date2;
		openLoadBox('loadFrame2');
	}
}

function loadoptions()
{	var date1 = document.forms['form1'].date1.value;
	var date2 = document.forms['form1'].date2.value;
	var utype = document.forms['form1'].utype.value;
	document.frames['options'].document.location.href = 'options.asp?bldg=<%=bldg%>&pid=<%=pid%>&utype='+utype+'&date1='+date1+'&date2='+date2;
}

function closeLoadBox(name)
{   document.all[name].style.visibility="hidden";
}
function openLoadBox(name)
{   var x=Math.floor(document.body.clientWidth/2-50)
    document.all[name].style.left=x
    document.all[name].style.visibility="visible";
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
function track(e)
{   mousey = event.clientX
    mousex = event.clientY
	return true
}


</script>

<style type=3D"text/css"><!--A {text-decoration: none}--></style>
<body bgcolor="#FFFFFF" text="#000000" link="#FFFFFF" vlink="#FFFFFF" alink="#0099FF" onmousemove="track(this)">
<h6></h6>
<form name="form1" method="post" action="">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="date1" value="<%=date1%>">
<input type="hidden" name="date2" value="">


<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center"  style="font-family: Arial, Helvetica, sans-serif; font-size: 13px; font-weight: bold; text-decoration:none;">
<tr bgcolor="black">
	<td id="tab7" width="2%" style="background-color:black"><a href="javascript:moveprev()" onMouseOver="this.style.color='lightblue';hoverHelp('prevhelp')" onMouseOut="this.style.color='white';HideHelp('prevhelp')">Previous&nbsp;Year</a>&nbsp;|&nbsp;</td>
	<td id="tab8" width="2%" style="background-color:black"><a href="javascript:movenext()" onMouseOver="this.style.color='lightblue';hoverHelp('nexthelp')" onMouseOut="this.style.color='white';HideHelp('nexthelp')">Next&nbsp;Year</a>&nbsp;|&nbsp;</td>
	<td id="tab9" width="2%" style="background-color:black"><a href="javascript:movenow()" onMouseOver="this.style.color='lightblue';hoverHelp('tdayhelp')" onMouseOut="this.style.color= 'white';HideHelp('tdayhelp')">This&nbsp;Year</a>&nbsp;|&nbsp;</td>
    <td  width="92%" align="right"><a href="/g1_clients/manual/lmp.htm" style="text-decoration:none;color:white" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Help</a>&nbsp;</td>
</tr></table>
<div align="center">&nbsp; <img src="../../images/lock.gif" width="16" height="17"> 
</div>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center"  style="font-family: Arial, Helvetica, sans-serif; font-size: 13px; font-weight: bold; text-decoration:none;">
<tr>
<!-- 	<td id="tab1" width="2%" style="background-color:#CCCCCC" onclick="settabs(this);storedPreferenceHottab=this">&nbsp;<a href="javascript:document.forms['form1'].utype.value='all';loadchart()" onMouseOver="this.style.color='black';" onMouseOut="this.style.color='white';">General&nbsp;Expense</a>&nbsp;</td> -->
	<td id="tab2" width="15%" style="background-color:#0099FF">&nbsp;<a href="javascript:loadchart()" onClick="settabs(document.all['tab2']);storedPreferenceHottab=document.all['tab2']" onMouseOver="this.style.color='black';" onMouseOut="this.style.color='white';">Revenue&nbsp;Profile</a>&nbsp;</td>
<!-- 	<td id="tab3" width="2%" style="background-color:#CCCCCC" onclick="settabs(this);storedPreferenceHottab=this">&nbsp;<a href="javascript:document.forms['form1'].utype.value='water';loadchart();" onMouseOver="this.style.color='black';" onMouseOut="this.style.color='white';">Water</a>&nbsp;</td> -->
<!-- 	<td id="tab4" width="2%" style="background-color:#CCCCCC" onclick="settabs(this);storedPreferenceHottab=this">&nbsp;<a href="javascript:document.forms['form1'].utype.value='steam';loadchart()" onMouseOver="this.style.color='black';" onMouseOut="this.style.color= 'white';">Steam</a>&nbsp;</td> -->
<!-- 	<td id="tab5" width="2%" style="background-color:#CCCCCC" onclick="settabs(this);storedPreferenceHottab=this">&nbsp;<a href="javascript:document.forms['form1'].utype.value='gas';loadchart()" onMouseOver="this.style.color='black';" onMouseOut="this.style.color= 'white';">Gas</a>&nbsp;</td> -->
	<td id="tab3" width="19%" style="background-color:#CCCCCC">&nbsp;<a href="preferences.asp" onClick="settabs(document.all['tab3']);" target="chart" onMouseOver="this.style.color='black';" onMouseOut="this.style.color= 'white';">View&nbsp;Preferences</a>&nbsp;&nbsp;&nbsp;</td>
	<td align="right">
    <select name="utype" onChange="loadchart();loadoptions()">
    <option value="0"<%if 0=cint(utype) then response.write " SELECTED"%>>All Utilities</option>
      <%
        rst1.open "SELECT * FROM tblutility ORDER BY utilitydisplay", getConnect(pid,bldg,"billing")
        do until rst1.eof
          %><option value="<%=rst1("utilityid")%>"<%if cint(rst1("utilityid"))=cint(utype) then response.write " SELECTED"%>><%=rst1("utilitydisplay")%></option><%
          rst1.movenext
        loop
      %>
    </select>
  </td>
	<td width="66%" align="right">&nbsp;</td>
</tr></table>

<table width="100%" border="0" cellspacing="0" cellpadding="0" height="277" align="center">
<tr><td width="100%" height="330"><iframe name="chart" id="lmp" style="width: 100%; height: 100%; border: 2px solid #0099FF;" " src="<%=IFrame1%>" scrolling="auto" marginwidth="0" marginheight="0"></iframe></td></tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center" height="290">
<tr><td><iframe name="options" width="100%" height="100%" src="<%=IFrame2%>" scrolling="auto" marginwidth="0" marginheight="0" style="border: 2px solid #0099FF;"></iframe></td></tr>
</table>







<div id="prevhelp" style="visibility:hidden; position:absolute;left:112;top:35;background-color:lightyellow;font-family:arial;font-size:10px;border-width:1px;border-style:solid">
Go to previous period
</div>
<div id="nexthelp" style="visibility:hidden; position:absolute;left:180;top:35;background-color:lightyellow;font-family:arial;font-size:10px;border-width:1px;border-style:solid">
Go to next period
</div>
<div id="tdayhelp" style="visibility:hidden; position:absolute;left:220;top:35;background-color:lightyellow;font-family:arial;font-size:10px;border-width:1px;border-style:solid">
Go to current period
</div>

<div id="loadFrame1" style="visibility:hidden; position:absolute;left:320;top:150;background-color:lightyellow;border-width:1px;border-style:solid">
<table border="0" cellpadding="5" cellspacing="0"><tr><td style="font-family:arial;font-size:16px;font-weight:bold">Loading Data...</td></tr></table>
</div>
<div id="loadFrame2" style="visibility:hidden; position:absolute;left:320;top:450;background-color:lightyellow;border-width:1px;border-style:solid">
<table border="0" cellpadding="5" cellspacing="0"><tr><td style="font-family:arial;font-size:16px;font-weight:bold">Loading Data...</td></tr></table>
</div>
</form>
</body>
<script>
storedPreferenceHottab = document.all['tab2']
openLoadBox('loadFrame1');
openLoadBox('loadFrame2');
</script>
</html>
