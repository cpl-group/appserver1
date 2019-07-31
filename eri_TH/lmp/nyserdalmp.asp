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

<style type=3D"text/css"><!--A {text-decoration: none}--></style>

<%
d=Request.QueryString("d")
b=Request.QueryString("bldg")
s=Request.QueryString("s")
e=Request.QueryString("e")
enflex = request("enflex")
portfolioid=Request.QueryString("portfolioid")
nozoom=Request.QueryString("nozoom")
luid = Request.QueryString("luid")

dim cnnM, rstM
Set cnnM = Server.CreateObject("ADODB.Connection")
Set rstM = Server.CreateObject("ADODB.recordset")
cnnM.Open application("cnnstr_genergy1")
rstM.open "SELECT meterid, lmnum from meters where bldgnum='"& b &"' and pp=1 order by meternum", cnnM

if not rstM.EOF then 
m=rstM("meterid")
lmp=trim(rstM("lmnum"))
else 
	m = 0
	lmp=1
end if
rstM.close
if session("roleid") = 1 then 'is tenant and needs to pull tenant info
    rstM.Open "SELECT LeaseUtilityId FROM tblLeases INNER JOIN tblLeasesUtilityPrices ON tblLeases.BillingId = tblLeasesUtilityPrices.BillingId WHERE tblLeases.TenantNum='"&session("userid")&"'", cnnM
    luid = rstM("LeaseUtilityId")
end if

if portfolioid<>"" then
    IFrame1 = "PortfolioAgg.asp?portfolioid="&portfolioid & "&d=" & d & "&s=" & s & "&e=" & e
else
    IFrame1 = "lmpload2.asp?m="&m&"&d="&d&"&b="&b&"&s="&s&"&e="&e&"&luid="&luid&"&lmp="&lmp&"&nozoom="&request.querystring("nozoom")
end if
IFrame2 = "options2.asp?m="&m&"&b="&b&"&luid="&luid
if trim(d)="" then d = date()
%>
<script>
function zoomentry(){
	var portfolioid = document.forms[0].portfolioid.value
	var tenantmeter = document.forms[0].tenantmeter.value
	var b = document.forms[0].b.value
	var m = document.forms[0].m.value
	var d = document.forms[0].d.value
	var s = document.forms[0].s.value
	var e = document.forms[0].e.value
	var l = document.forms[0].luid.value
	var lmp = document.forms[0].lmp.value
	var temp = "zoomentry2.asp?b=" + b + "&m=" + m + "&d=" + d + "&s="+s+"&e="+e+"&luid="+l+"&lmp="+lmp+"&portfolioid="+portfolioid+"&tenantmeter="+tenantmeter
	window.open(temp,"","statusbar=0,menubar=0,scrollbars=yes,HEIGHT=125,WIDTH=300")
}
function lmpmoveprev(){
	var portfolioid = document.forms[0].portfolioid.value
	var tenantmeter = document.forms[0].tenantmeter.value
	var b = document.forms[0].b.value
	var m = document.forms[0].m.value
	var d = document.forms[0].pd.value
	var nd = document.forms[0].d.value
	var s = document.forms[0].s.value
	var e = document.forms[0].e.value
	var l = document.forms[0].luid.value
	var lmp = document.forms[0].lmp.value
	var pd = new Date(d)
	pd.setTime(pd.getTime() - 1 * 24 * 60 * 60 * 1000)
	pd = (pd.getMonth()+1) + "/" + pd.getDate() + "/" + pd.getYear()
		
	document.forms[0].d.value = d
	document.forms[0].pd.value = pd
	document.forms[0].nd.value = nd

	loadchart()
}

function lmpmovenext(){
	var portfolioid = document.forms[0].portfolioid.value
	var tenantmeter = document.forms[0].tenantmeter.value
	var b = document.forms[0].b.value
	var m = document.forms[0].m.value
	var d = document.forms[0].nd.value
	var pd = document.forms[0].d.value
	var s = document.forms[0].s.value
	var e = document.forms[0].e.value
	var l = document.forms[0].luid.value
	var lmp = document.forms[0].lmp.value
	var nd = new Date(d)
	nd.setTime(nd.getTime() + 1 * 24 * 60 * 60 * 1000)
	nd = (nd.getMonth()+1) + "/" + nd.getDate() + "/" + nd.getYear()

	document.forms[0].d.value = d
	document.forms[0].pd.value = pd
	document.forms[0].nd.value = nd

	loadchart()
}

function lmpnow(){
	var portfolioid = document.forms[0].portfolioid.value
	var tenantmeter = document.forms[0].tenantmeter.value
	var b = document.forms[0].b.value
	var m = document.forms[0].m.value
	var d = document.forms[0].td.value
	var s = document.forms[0].s.value
	var e = document.forms[0].e.value
	var l = document.forms[0].luid.value
	var lmp = document.forms[0].lmp.value
	var nd = new Date(d)
	nd.setTime(nd.getTime() + 1 * 24 * 60 * 60 * 1000)
	nd = (nd.getMonth()+1) + "/" + nd.getDate() + "/" + nd.getYear()
	var pd = new Date(d)
	pd.setTime(pd.getTime() - 1 * 24 * 60 * 60 * 1000)
	pd = (pd.getMonth()+1) + "/" + pd.getDate() + "/" + pd.getYear()

	document.forms[0].d.value = d
	document.forms[0].pd.value = pd
	document.forms[0].nd.value = nd

	loadchart()
}

function calnow()
{	var d = "<%=month(date())%>/<%=day(date())%>/<%=year(date)%>";
	document.forms[0].d.value = d
	document.forms[0].pd.value = dateAddDays(d,-1);
	document.forms[0].nd.value = dateAddDays(d,1);
	loadcalendar();
}
function dateAddDays(d, days)
{	d = new Date(d);
	d.setTime(d.getTime() + days * 24 * 60 * 60 * 1000);
	d = (d.getMonth()+1) + "/" + d.getDate() + "/" + d.getYear();
	return(d);
}
function calprev()
{	var d = new Date(document.forms[0].d.value)
	var month = d.getMonth()-1;
	var year = d.getYear();
	if(month<0){month=11;year--;}
	d = (month+1) + "/1/" + year;
	document.forms[0].d.value = d
	document.forms[0].pd.value = dateAddDays(d,-1);
	document.forms[0].nd.value = dateAddDays(d,1);
	loadcalendar();
}
function calnext()
{	var d = new Date(document.forms[0].d.value)
	var month = d.getMonth()+1;
	var year = d.getYear();
	if(month>11){month=0;year++;}
	d = (month+1) + "/1/" + year;
	document.forms[0].d.value = d
	document.forms[0].pd.value = dateAddDays(d,-1);
	document.forms[0].nd.value = dateAddDays(d,1);
	loadcalendar();
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

function shownav(navname)
{   if(navname=='lmpnav')
    {   document.all['lmpnav'].style.position='relative';
        document.all['lmpnav'].style.visibility='visible';
        document.all['lmptabs'].style.visibility='visible';
        document.all['peakCnav'].style.position='absolute';
        document.all['peakCnav'].style.visibility='hidden';
        document.all['calnav'].style.position='absolute';
        document.all['calnav'].style.visibility='hidden';
    }else if(navname=='peakCnav')
    {   document.all['peakCnav'].style.position='relative';
        document.all['peakCnav'].style.visibility='visible';
        document.all['lmpnav'].style.position='absolute';
        document.all['lmpnav'].style.visibility='hidden';
        document.all['calnav'].style.position='absolute';
        document.all['calnav'].style.visibility='hidden';
        document.all['lmptabs'].style.visibility='hidden';
    }else if(navname=='calnav')
    {   document.all['calnav'].style.position='relative';
        document.all['calnav'].style.visibility='visible';
        document.all['lmpnav'].style.position='absolute';
        document.all['lmpnav'].style.visibility='hidden';
        document.all['peakCnav'].style.position='absolute';
        document.all['peakCnav'].style.visibility='hidden';
        document.all['lmptabs'].style.visibility='visible';
	}
}

function PeakCmoveprev()
{   var year = dataset.document.forms['PDpieposition'].byear.value;
    var period = dataset.document.forms['PDpieposition'].bperiod.value
    var luid = dataset.document.forms['PDpieposition'].luid.value
    var m = dataset.document.forms['PDpieposition'].m.value
    period--;
    if(period<1)
    {   
		period=12;
        year--;
    }
    //alert('opt_PeakDemand.asp?b=<%=b%>&luid=<%=luid%>&explode=&byear='+ year +'&bperiod='+ period +'&coor=');
    document.dataset.location.href='opt_PeakDemand.asp?b=<%=b%>&explode=&byear='+ year +'&bperiod='+ period +'&luid='+ luid +'&m='+ m +'&coor=';
}

function PeakCmovenext()
{   var year = dataset.document.forms['PDpieposition'].byear.value;
    var period = dataset.document.forms['PDpieposition'].bperiod.value
    var luid = dataset.document.forms['PDpieposition'].luid.value
    var m = dataset.document.forms['PDpieposition'].m.value
    period++;
    if(period>12)
    {   period=1;
        year++;
    }
    //alert('opt_PeakDemand.asp?b=<%=b%>&luid=<%=luid%>&explode=&byear='+ year +'&bperiod='+ period +'&coor=');
    document.dataset.location.href='opt_PeakDemand.asp?b=<%=b%>&explode=&byear='+ year +'&bperiod='+ period +'&luid='+ luid +'&m='+ m +'&coor=';
}

function PeakCnow()
{   var year = <%=datepart("YYYY", date())%>;
    var period = <%=datepart("m", date())%>;
    var luid = dataset.document.forms['PDpieposition'].luid.value
    var m = dataset.document.forms['PDpieposition'].m.value
    document.dataset.location.href='opt_PeakDemand.asp?b=<%=b%>&explode=&byear='+ year +'&bperiod='+ period +'&luid='+ luid +'&m='+ m +'&coor=';
}

var storedPreferenceHottab //this var is intended to remember the hot tab when preferences has been clicked--for the posiibility of pressing 'next', 'prev' or 'today' while in preferences

function settabs(hottab)
{	for(i=1;i<=3;i++)
	{	document.all['tab'+i].style.backgroundColor='#CCCCCC'
	}
	if(hottab!=0) hottab.style.backgroundColor='#0099FF'
}

function loadchart()
{	var m = document.forms['form1'].m.value;
	var d = document.forms['form1'].d.value;
	var b = document.forms['form1'].b.value;
	var l = document.forms['form1'].luid.value;
	var i = document.forms['form1'].zoom.value;
	var lmp = document.forms['form1'].lmp.value;
	var tenantmeter = document.forms['form1'].tenantmeter.value;
  var lmpchartload = document.forms['form1'].lmpchartload.value
  if(lmpchartload=="1"){
  	temp="lmpload2.asp?m="+m+"&d="+d+"&b="+b+"&s=&e=&luid="+l+"&lmp="+lmp+"&tenantmeter="+tenantmeter+"&i="+i;
	  openLoadBox('loadFrame1')
	  document.frames.lmp.location=temp;
	  if(i==1)	{settabs(document.all['tab2']);}
	  else		{settabs(document.all['tab1']);}
    var dataset = document.frames.dataset.location.toString();
    if(dataset.indexOf("loadRateComp")>=0) document.frames.dataset.location = "<%=IFrame2%>"
  }else{
    document.frames.lmp.location="loadRateComp.asp?bldgid="+b+"&lmpdate="+d+"&qrytype=actual&graphtype=6"
    document.frames.dataset.location="loadRateComp.asp?bldgid="+b+"&lmpdate="+d+"&qrytype=dam&graphtype=6"
  }
}

function loadcalendar()
{	var m = document.forms['form1'].m.value;
	var d = document.forms['form1'].d.value;
	var b = document.forms['form1'].b.value;
	var l = document.forms['form1'].luid.value;
	var i = document.forms['form1'].zoom.value;
	var lmp = document.forms['form1'].lmp.value;
	var tenantmeter = document.forms['form1'].tenantmeter.value;
	var popups = '';
  if(document.frames['lmp'].document.forms['popups']!=undefined) popups = document.frames['lmp'].document.forms['popups'].popups.checked;
	temp="graphcalendar.asp?m="+m+"&date="+d+"&b="+b+"&s=&e=&luid="+l+"&lmp="+lmp+"&tenantmeter="+tenantmeter+"&i="+i+"&popups="+popups;
	openLoadBox('loadFrame1')
	document.frames.lmp.location=temp;
  var dataset = document.frames.dataset.location.toString();
  if(dataset.indexOf("loadRateComp")>=0) document.frames.dataset.location = "<%=IFrame2%>"
}

function loadRateComp()
{ var b = document.forms['form1'].b.value;
  var d = document.forms['form1'].d.value;
  turnRateComp('on');
//  settabs(0);
  document.frames.lmp.location="loadRateComp.asp?bldgid="+b+"&lmpdate="+d+"&qrytype=actual&graphtype=6"
  document.frames.dataset.location="loadRateComp.asp?bldgid="+b+"&lmpdate="+d+"&qrytype=dam&graphtype=6"
}

function turnRateComp(onOff)
{ if(onOff=='on')
  { document.forms['form1'].lmpchartload.value = 0;
    document.all['lmptabs'].style.display = "none";
    document.all['RateCompTab'].style.display = "inline";
  }else
  { document.forms['form1'].lmpchartload.value = 1;
    document.all['lmptabs'].style.display = "inline";
    document.all['RateCompTab'].style.display = "none";
  }
}
</script>
<body bgcolor="#FFFFFF" text="#000000" link="#FFFFFF" vlink="#FFFFFF" alink="#0099FF" onmousemove="track(this)">
<h6></h6>
<form name="form1" method="post" action="">
              <input type="hidden" name="ppmeter" value="<%=m%>">
              <input type="hidden" name="b" value="<%=b%>">
              <input type="hidden" name="m" value="<%=m%>">
              <input type="hidden" name="s" value="<%=s%>">
              <input type="hidden" name="e" value="<%=e%>">
              <input type="hidden" name="nozoom" value="<%=nozoom%>">
              <input type="hidden" name="d" value="<%=d%>">
              <input type="hidden" name="pd" value="<%'DateAdd("d",-1,d)%>">
              <input type="hidden" name="nd" value="<%'DateAdd("d",1,d)%>">
              <input type="hidden" name="td" value="<%=Date()%>">
              <input type="hidden" name="luid" value="<%=luid%>">
              <input type="hidden" name="lmp" value="<%=lmp%>">
              <input type="hidden" name="tenantmeter" value="<%=tenantmeter%>">
              <input type="hidden" name="portfolioid" value="<%=portfolioid%>">
              <input type="hidden" name="zoom" value="100">
              <input type="hidden" name="lmpchartload" value="1">

<table width="714" border="0" cellspacing="0" cellpadding="0" align="center">
<tr><td width="687" bgcolor="#000000"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000">
        <tr><td width="100%" height="2" bgcolor="#000000">
                <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr><td><b><font size="2" face="Arial, Helvetica, sans-serif" color="#FFFFFF">                      
                    <div id="lmpnav" style="visibility:visible;position:relative;left:0;top:0;"> 
                      <!-- <a href="javascript:zoomentry()" style="text-decoration:none;" onMouseOver="this.style.color='lightblue';hoverHelp('zoomhelp')" onMouseOut="this.style.color='white';HideHelp('zoomhelp')">Interval Zoom</a> |  -->
                      <a href="javascript:lmpmoveprev()" style="text-decoration:none;" onMouseOver="this.style.color='lightblue';hoverHelp('prevhelp')" onMouseOut="this.style.color='white';HideHelp('prevhelp')">Previous 
                      Day</a> | <a href="javascript:lmpmovenext()" style="text-decoration:none;" onMouseOver="this.style.color='lightblue';hoverHelp('nexthelp')" onMouseOut="this.style.color='white';HideHelp('nexthelp')">Next 
                      Day</a> | <a href="javascript:lmpnow()" style="text-decoration:none;" onMouseOver="this.style.color='lightblue';hoverHelp('tdayhelp')" onMouseOut="this.style.color= 'white';HideHelp('tdayhelp')">Go 
                      To Today</a> </div>
                        <div id="peakCnav" style="visibility:hidden;position:absolute;left:0;top:0;">
                            <a href="javascript:PeakCmoveprev()" style="text-decoration:none;" onMouseOver="this.style.color='lightblue';hoverHelp('prevhelp')" onMouseOut="this.style.color='white';HideHelp('prevhelp')">Previous Period</a> | 
                            <a href="javascript:PeakCmovenext()" style="text-decoration:none;" onMouseOver="this.style.color='lightblue';hoverHelp('nexthelp')" onMouseOut="this.style.color='white';HideHelp('nexthelp')">Next Period</a> | 
                            <a href="javascript:PeakCnow()" style="text-decoration:none;" onMouseOver="this.style.color='lightblue';hoverHelp('tdayhelp')" onMouseOut="this.style.color= 'white';HideHelp('tdayhelp')">Go To Current Period</a>
                        </div>
                        <div id="calnav" style="visibility:hidden;position:absolute;left:0;top:0;">
                            <a href="javascript:calprev()" style="text-decoration:none;" onMouseOver="this.style.color='lightblue';hoverHelp('prevhelp')" onMouseOut="this.style.color='white';HideHelp('prevhelp')">Previous Month</a> | 
                            <a href="javascript:calnext()" style="text-decoration:none;" onMouseOver="this.style.color='lightblue';hoverHelp('nexthelp')" onMouseOut="this.style.color='white';HideHelp('nexthelp')">Next Month</a> | 
                            <a href="javascript:calnow()" style="text-decoration:none;" onMouseOver="this.style.color='lightblue';hoverHelp('tdayhelp')" onMouseOut="this.style.color= 'white';HideHelp('tdayhelp')">Go To Current Month</a>
                        </div>
                        </font></b>
                    </td>
                    <td align="right"><b><font size="2" face="Arial, Helvetica, sans-serif" color="#FFFFFF"><a href="/g1_clients/manual/lmp.htm" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Help</a>&nbsp;</font></b></td>
                </tr></table>
            </td>
        </tr></table>
    </td>
</tr></table>
&nbsp;
  <div id="lmptabs" style="visibility:visible;position:relative;left:0;top:0;display:inline"> 
    <table width="714" border="0" cellspacing="0" cellpadding="0" align="center"  style="font-family: Arial, Helvetica, sans-serif; font-size: 13px; font-weight: bold; text-decoration:none;">
      <tr> 
        <td id="tab1" width="5%" style="background-color:#CCCCCC">&nbsp;<a href="javascript:loadchart()" onclick="settabs(document.all['tab1']);document.forms['form1'].zoom.value='100';" onMouseOver="this.style.color='black';" onMouseOut="this.style.color='white';" style="color:white">Hourly&nbsp;Intervals</a>&nbsp;</td>
        <td id="tab2" width="5%" style="background-color:#0099FF">&nbsp;<a href="javascript:loadchart()" onclick="settabs(document.all['tab2']);document.forms['form1'].zoom.value='1';" onMouseOver="this.style.color='black';" onMouseOut="this.style.color= 'white';" style="color:white">15&nbsp;Minute&nbsp;Intervals</a>&nbsp;</td>
        <td id="tab3" width="5%" style="background-color:#CCCCCC">&nbsp;<a href="javascript:loadcalendar()" onclick="settabs(document.all['tab3']);document.forms['form1'].lmpchartload.value = 1" onMouseOver="this.style.color='black';" onMouseOut="this.style.color= 'white';" style="color:white">Calendar</a>&nbsp;</td>
        <td  width="85%" align="right">&nbsp;</td>
      </tr>
    </table>
  </div>
  <div id="RateCompTab" style="visibility:visible;position:relative;left:0;top:0;display:none"> 
    <table width="714" border="0" cellspacing="0" cellpadding="0" align="center"  style="font-family: Arial, Helvetica, sans-serif; font-size: 13px; font-weight: bold; text-decoration:none;" bgcolor="black">
      <tr><td width="5%">&nbsp;<a href="javascript:loadchart()" onclick="turnRateComp('off')" onMouseOver="this.style.color='lightblue';" onMouseOut="this.style.color='white';" style="color:white;text-decoration:none;">Back&nbsp;To&nbsp;Options</a>&nbsp;</td>
        <td  width="85%" align="right">&nbsp;</td>
      </tr>
    </table>
  </div>
<table width="714" border="0" cellspacing="0" cellpadding="0" height="277" align="center">
<tr><td width="687" height="350"><iframe src="<%=IFrame1%>" name="lmp" id="lmp" width="100%" height="100%" marginwidth="0" marginheight="0" style="border: 2px solid #0099FF;"></iframe></td></tr>
<%if trim(enflex)<>"" then%>
<tr><td width="687" height="310"><iframe src="<%=enflex%>" name="dataset" id="dataset" width="100%" height="100%" marginwidth="0" marginheight="0" style="border: 2px solid #0099FF;"></iframe></td></tr>
<%else%>
<tr><td width="687" height="310"><iframe src="<%=IFrame2%>" name="dataset" id="dataset" width="100%" height="100%" marginwidth="0" marginheight="0" style="border: 2px solid #0099FF;"></iframe></td></tr>
<%end if%>
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

<div id="loadFrame1" style="visibility:hidden; position:absolute;left:320;top:150;background-color:lightyellow;border-width:1px;border-style:solid">
<table border="0" cellpadding="5" cellspacing="0"><tr><td style="font-family:arial;font-size:16px;font-weight:bold">Loading Data...</td></tr></table>
</div>
<div id="loadFrame2" style="visibility:hidden; position:absolute;left:320;top:450;background-color:lightyellow;border-width:1px;border-style:solid">
<table border="0" cellpadding="5" cellspacing="0"><tr><td style="font-family:arial;font-size:16px;font-weight:bold">Loading Data...</td></tr></table></div></body>
<script>
openLoadBox('loadFrame1');
//openLoadBox('loadFrame2');
</script>
</html>
