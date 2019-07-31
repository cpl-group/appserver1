<%option explicit%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<HTML>

<head>
<title>Load Management</title>
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
dim startdate, bldg, pid, meterid, IFrame1, IFrame2, utilityid, billingid, enflex, hideOptions
startdate=Request.QueryString("startdate")
bldg=Request.QueryString("bldg")
billingid=Request.QueryString("billingid")
meterid=Request.QueryString("meterid")
utilityid=Request.QueryString("utility")
pid=Request.QueryString("pid")
enflex = request("enflex")
hideOptions = request("hideoptions")

if isEmpty(hideOptions) then 
	hideOptions = false
else 
	hideOptions = true
end if 

if trim(utilityid)="" then utilityid=2
dim cnnM, rstM
Set cnnM = Server.CreateObject("ADODB.Connection")
Set rstM = Server.CreateObject("ADODB.recordset")
if trim(bldg) <> "" then cnnM.open getLocalConnect(bldg) else cnnM.open getMainConnect(pid) 
dim test1

'test1=session("roleid") 
'response.write "here:" & "<BR>"
'response.write test1' = "dariovno"
'response.end
'or isnull(session("roleid"))
if session("roleid")="" then
session("roleid") =4
end if
if  session("roleid") = 1 and billingid = "" then 'is tenant and needs to pull tenant info
    rstM.Open "SELECT billingid FROM tblLeases WHERE TenantNum='"&session("userid")&"'", cnnM
    billingid = rstM("billingid")
end if


IFrame1 = "lmpload.asp?meterid="&meterid&"&startdate="&startdate&"&bldg="&bldg&"&billingid="&billingid&"&interval=0&utility="&utilityid&"&pid="&pid

if pid<>"" then
    IFrame2 = "portfolioBreakdown.asp?pid="&pid&"&startdate="&startdate&"&utility="&utilityid
else
    IFrame2 = "options.asp?meterid="&meterid&"&bldg="&bldg&"&billingid="&billingid&"&utility="&utilityid
end if

%>
<script>
function lmpmoveprev(){
	var startdate = document.forms['form1'].startdate.value;
	startdate = dateAddDays(startdate,-1)
	document.forms[0].startdate.value = startdate
	<%if pid<>"" then%>
		document.frames.dataset.location = "portfolioBreakdown.asp?pid=<%=pid%>&startdate=" + startdate + "&utility=<%=utilityid%>"
	<%end if%>
	loadchart()

}

function lmpmovenext(){
	var startdate = document.forms['form1'].startdate.value;
	startdate = dateAddDays(startdate,1);
	document.forms[0].startdate.value = startdate
	<%if pid<>"" then%>
		document.frames.dataset.location = "portfolioBreakdown.asp?pid=<%=pid%>&startdate=" + startdate + "&utility=<%=utilityid%>"
	<%end if%>
	loadchart()

}

function lmpnow(){
	var startdate = document.forms['form1'].startdate.value;
  startdate = document.forms[0].today.value
	document.forms[0].startdate.value = startdate
	<%if pid<>"" then%>
		document.frames.dataset.location = "portfolioBreakdown.asp?pid=<%=pid%>&startdate=" + startdate + "&utility=<%=utilityid%>"
	<%end if%>
	loadchart()
}

function calnow()
{	var startdate = "<%=month(date())%>/1/<%=year(date)%>";
	document.forms[0].startdate.value = startdate
	loadcalendar();
}

function dateAddDays(d, days)
{	d = new Date(d);
	d.setTime(d.getTime() + days * 24 * 60 * 60 * 1000);
	d = (d.getMonth()+1) + "/" + d.getDate() + "/" + d.getYear();
	return(d);
}

function calprev()
{	var startdate = new Date(document.forms[0].startdate.value)
	var month = startdate.getMonth()-1;
	var year = startdate.getYear();
	if(month<0){month=11;year--;}
	startdate = (month+1) + "/1/" + year;
	document.forms[0].startdate.value = startdate;
	loadcalendar();
}
function calnext()
{	var startdate = new Date(document.forms[0].startdate.value)
	var month = startdate.getMonth()+1;
	var year = startdate.getYear();
	if(month>11){month=0;year++;}
	startdate = (month+1) + "/1/" + year;
	document.forms[0].startdate.value = startdate;
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
    var meterid = dataset.document.forms['PDpieposition'].meterid.value
  	var utility = document.forms['form1'].utility.value;
    period--;
    if(period<1)
    {   period=12;
        year--;
    }
    document.dataset.location.href='opt_PeakDemand.asp?bldg=<%=bldg%>&utility='+utility+'&explode=&byear='+ year +'&bperiod='+ period +'&luid='+ luid +'&meterid='+ meterid +'&coor=';
}

function PeakCmovenext()
{   var year = dataset.document.forms['PDpieposition'].byear.value;
    var period = dataset.document.forms['PDpieposition'].bperiod.value
    var luid = dataset.document.forms['PDpieposition'].luid.value
    var meterid = dataset.document.forms['PDpieposition'].meterid.value
  	var utility = document.forms['form1'].utility.value;
    period++;
    if(period>12)
    {   period=1;
        year++;
    }
    document.dataset.location.href='opt_PeakDemand.asp?bldg=<%=bldg%>&utility='+utility+'&explode=&byear='+ year +'&bperiod='+ period +'&luid='+ luid +'&meterid='+ meterid +'&coor=';
}

function PeakCnow()
{   var year = <%=datepart("YYYY", date())%>;
    var period = <%=datepart("m", date())%>;
    var luid = dataset.document.forms['PDpieposition'].luid.value
    var meterid = dataset.document.forms['PDpieposition'].meterid.value
  	var utility = document.forms['form1'].utility.value;
    document.dataset.location.href='opt_PeakDemand.asp?bldg=<%=bldg%>&utility='+utility+'&explode=&byear='+ year +'&bperiod='+ period +'&luid='+ luid +'&meterid='+ meterid +'&coor=';
}

var storedPreferenceHottab //this var is intended to remember the hot tab when preferences has been clicked--for the posiibility of pressing 'next', 'prev' or 'today' while in preferences

function settabs(hottab)
{	for(i=1;i<=3;i++)
	{	document.all['tab'+i].style.backgroundColor='#CCCCCC'
	}
	hottab.style.backgroundColor='#0099FF'
}

function loadOpt_T()
{ document.forms['form1'].meterid.value='';
  document.forms['form1'].billingid.value='';
//  document.forms['form1'].billingid.value='';
  var meterid = ''//document.forms['form1'].meterid.value;
	var startdate = document.forms['form1'].startdate.value;
	var bldg = document.forms['form1'].bldg.value;
	var billingid = document.forms['form1'].billingid.value;
	var interval = document.forms['form1'].interval.value;
	var utility = document.forms['form1'].utility.value;
  <%if trim(enflex)="" and  trim(bldg) <> ""  then%>
  var lowerframe = new String(document.frames.dataset.location);
  if(lowerframe.indexOf("opt_tenantPF.asp")!=-1)	document.frames.dataset.location='opt_tenantPF.asp?bldg='+bldg+'&meterid='+meterid+'&utility='+utility//+'&billingid='+billingid
  else document.frames.dataset.location='options.asp?bldg='+bldg+'&meterid='+meterid+'&utility='+utility//+'&billingid='+billingid
  <%end if%>
}

function loadchart()
{	var meterid = document.forms['form1'].meterid.value;
	var startdate = document.forms['form1'].startdate.value;
	var bldg = document.forms['form1'].bldg.value;
	var billingid = document.forms['form1'].billingid.value;
	var interval = document.forms['form1'].interval.value;
	var utility = document.forms['form1'].utility.value;
	var lmpchartload = document.forms['form1'].lmpchartload.value;
  if(lmpchartload=="1"){
  	temp="lmpload.asp?meterid="+meterid+"&startdate="+startdate+"&bldg="+bldg+"&billingid="+billingid+"&interval="+interval+"&utility="+utility+"&pid=<%=pid%>";
  	openLoadBox('loadFrame1')
  	document.frames.lmp.location=temp;
  	if(interval==0)	{settabs(document.all['tab2']);}
  	else		{settabs(document.all['tab1']);}
    <%if trim(enflex)="" and  trim(bldg) <> ""  then%>
		try{
	    var dataset = document.frames.dataset.location.toString();
  	  if(dataset.indexOf("loadRateComp")>=0) document.frames.dataset.location = "<%=IFrame2%>"
		}catch(exception){}
    <%end if%>
  }else{
    document.frames.lmp.location="loadRateComp.asp?bldgid="+bldg+"&lmpdate="+startdate+"&qrytype=actual&graphtype=6"
    try{document.frames.dataset.location="loadRateComp.asp?bldgid="+bldg+"&lmpdate="+startdate+"&qrytype=dam&graphtype=6"}catch(exception){}
  }
}

function loadcalendar()
{	var meterid = document.forms['form1'].meterid.value;
	var startdate = document.forms['form1'].startdate.value;
	var bldg = document.forms['form1'].bldg.value;
	var billingid = document.forms['form1'].billingid.value;
	var interval = document.forms['form1'].interval.value;
	var utility = document.forms['form1'].utility.value;
	var popups = '';
  if(document.frames['lmp'].document.forms['popups']!=undefined)	
	{popups = document.frames['lmp'].document.forms['popups'].popups.checked
	}
	temp="graphcalendar.asp?meterid="+meterid+"&date="+startdate+"&bldg="+bldg+"&billingid="+billingid+"&interval="+interval+"&utility="+utility+"&popups="+popups;
	openLoadBox('loadFrame1')
	document.frames.lmp.location=temp;
  <%if trim(enflex)="" and  trim(bldg) <> ""  then%>
	try{
	  var dataset = document.frames.dataset.location.toString();
  	if(dataset.indexOf("loadRateComp")>=0) document.frames.dataset.location = "<%=IFrame2%>"
	}catch(exception){}
  <%end if%>
}

function loadRateComp()
{ var bldg = document.forms['form1'].bldg.value;
  var startdate = document.forms['form1'].startdate.value;
  turnRateComp('on');
//  settabs(0);
  document.frames.lmp.location="loadRateComp.asp?bldgid="+bldg+"&lmpdate="+startdate+"&qrytype=actual&graphtype=6"
  document.frames.dataset.location="loadRateComp.asp?bldgid="+bldg+"&lmpdate="+startdate+"&qrytype=dam&graphtype=6"
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
              <input type="hidden" name="bldg" value="">
              <input type="hidden" name="meterid" value="<%=meterid%>">
              <input type="hidden" name="startdate" value="">
              <input type="hidden" name="today" value="<%=Date()%>">
              <input type="hidden" name="billingid" value="<%=billingid%>">
              <input type="hidden" name="pid" value="">
              <input type="hidden" name="interval" value="0">
              <input type="hidden" name="lmpchartload" value="1">

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr><td width="100%" bgcolor="#000000"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000">
        <tr><td width="100%" height="2" bgcolor="#000000">
                <table width="100%" cellpadding="0" cellspacing="0" border="0">
                <tr><td><b><font size="2" face="Arial, Helvetica, sans-serif" color="#FFFFFF">
                        
                    
		            <div id="lmpnav" style="visibility:visible;position:relative;left:0;top:0;"> 
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
  
  <div id="lmptabs" style="visibility:visible;position:relative;left:0;top:0;"> 
    <table width="100%" border="0" cellspacing="0" cellpadding="0" align="center"  style="font-family: Arial, Helvetica, sans-serif; font-size: 13px; font-weight: bold; text-decoration:none;">
      <tr> 
        <td id="tab1" width="5%" style="background-color:#CCCCCC">&nbsp;<a href="javascript:loadchart()" onClick="settabs(document.all['tab1']);document.forms['form1'].interval.value='1';" onMouseOver="this.style.color='black';" onMouseOut="this.style.color='white';" style="color:white">Hourly&nbsp;Intervals</a>&nbsp;</td>
        <td id="tab2" width="5%" style="background-color:#0099FF">&nbsp;<a href="javascript:loadchart()" onClick="settabs(document.all['tab2']);document.forms['form1'].interval.value='0';" onMouseOver="this.style.color='black';" onMouseOut="this.style.color= 'white';" style="color:white">15&nbsp;Minute&nbsp;Intervals</a>&nbsp;</td>
        <td id="tab3" width="5%" style="background-color:#CCCCCC"> 
          <%if trim(pid) = "" then%>
          &nbsp;<a href="javascript:loadcalendar()" onClick="settabs(document.all['tab3']);" onMouseOver="this.style.color='black';" onMouseOut="this.style.color= 'white';" style="color:white">Calendar</a>&nbsp; 
          <%end if%>
        </td>
        <td  width="5%" align="right"> 
          <% if not hideoptions then %>
          <select name="utility" onChange="loadOpt_T();loadchart()">
            <%
      rstM.open "SELECT * FROM tblutility ORDER BY utilitydisplay", getConnect(pid,bldg,"billing")
      do until rstM.eof
        %>
            <option value="<%=rstM("utilityid")%>"<%if cint(utilityid)=cint(rstM("utilityid")) then response.write " SELECTED"%>><%=rstM("utilitydisplay")%></option>
            <%
        rstM.movenext
      loop
      rstM.close
      %>
          </select> 
          <%else%>
          <input type="hidden" name="utility" value="<%=utilityid%>"> 
          <%end if%>
        </td>
        <td  width="80%" align="right">&nbsp;</td>
      </tr>
    </table>
  </div>
  
  <div id="RateCompTab" style="visibility:visible;position:relative;left:0;top:0;display:none"> 
    <table width="714" border="0" cellspacing="0" cellpadding="0" align="center"  style="font-family: Arial, Helvetica, sans-serif; font-size: 13px; font-weight: bold; text-decoration:none;" bgcolor="black">
      <tr> 
        <td width="5%">&nbsp;<a href="javascript:loadchart()" onClick="turnRateComp('off')" onMouseOver="this.style.color='lightblue';" onMouseOut="this.style.color='white';" style="color:white;text-decoration:none;">Back&nbsp;To&nbsp;Options</a>&nbsp;</td>
        <td  width="85%" align="right">&nbsp;</td>
      </tr>
    </table>
  </div>

<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center">
<tr><td width="100%" height="350"><iframe src="<%=IFrame1%>" name="lmp" id="lmp" width="100%" height="100%" marginwidth="0" marginheight="0" style="border: 2px solid #0099FF;"></iframe></td></tr>
<tr>
  <td>
  <iframe src="options.asp" name="optionFrame" id="optFrame" width="100%" height="100%" style="border:2px solid #0099ff"></iframe>
  </td>
</tr>

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
<table border="0" cellpadding="5" cellspacing="0"><tr><td style="font-family:arial;font-size:16px;font-weight:bold">Loading Data...</td></tr></table>
</div>
</body>
<script>
openLoadBox('loadFrame1');
<%if trim(enflex)="" and  trim(bldg) <> ""  and not hideOptions then%>
openLoadBox('loadFrame2');
<%end if%>
</script>

</html>
