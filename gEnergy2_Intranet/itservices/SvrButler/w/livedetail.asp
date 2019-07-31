<script>
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

<title>Live Snapshot: <%=request("server")%>_____________</title><body bgcolor="#eeeeee">
<table width="100%"  bgcolor="#eeeeee">
  <tr>
    <td width="50%" valign="top"> 
      <%
on error resume next
Server.ScriptTimeout = 500000000
	strComputer = request("server")
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk where DriveType=3")

	Set object = objWMIService.ExecQuery("Select * from win32_processor")

	For each element in object
		LoadPercent 	= element.LoadPercentage/100
		cpuStatus 	= element.CpuStatus
	
	Next	
	Set object = objWMIService.ExecQuery("Select * From Win32_LogicalMemoryConfiguration")

	For each element in object
		TotalMem = ((element.TotalPhysicalMemory + 1023)/1024)
	Next	

	Set object = objWMIService.ExecQuery("Select * From Win32_PerfRawData_PerfOS_Memory")

	For each element in object
		freemem = element.AvailableMBytes	
	Next	
%>
      <table cellpadding="0" cellspacing="0">
  <tr> 
    <td><font size="2" face="Century Gothic">Server Name</font></td>
    <td colspan=2><font size="2" face="Century Gothic">: <%=ucase(strComputer)%></font></td>
  </tr>
  <tr> 
    <td><font size="2" face="Century Gothic">% CPU</font></td>
    <td colspan=2><font size="2" face="Century Gothic">: <%=formatpercent(loadpercent)%></font></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
          <td style="border-bottom:1px solid #000000" align="center"><font size="2" face="Century Gothic"><strong>Total</strong></font></td>
          <td style="border-bottom:1px solid #000000" align="center"><font size="2" face="Century Gothic"><strong>Free</strong></font></td>
  </tr>
  <tr> 
    <td><font size="2" face="Century Gothic">Memory:</font></td>
    <td align="right" style="border-right:1px solid #000000"><font size="2" face="Century Gothic"><%=formatnumber(totalmem/1024,2)%> GB &nbsp;</font></td>
    <td align="right" style="border-right:1px solid #000000"><font size="2" face="Century Gothic"><%=formatnumber(freemem/1024,2)%> GB &nbsp;</font></td>
  </tr>
  <%	
  	TSS = 0
	TFS = 0
		For each objDisk in colDisks
			Drive = objDisk.DeviceID
			TotalSize = objDisk.Size
			Freespace = objDisk.FreeSpace
			TSS = tss + totalsize
			TFS = tfs + freespace
%>
  <tr> 
          <td><font size="2" face="Century Gothic">Drive <%=Drive%></font></td>
    <td align="right" style="border-right:1px solid #000000"><font size="2" face="Century Gothic">&nbsp;<%=formatnumber((( TotalSize/1024)/1024)/1024,2)%> GB &nbsp;</font></td>
    <td align="right" style="border-right:1px solid #000000"><font size="2" face="Century Gothic">&nbsp;<%=formatnumber((( freespace/1024)/1024)/1024,2)%> GB &nbsp;</font></td>
  </tr>
  <%
		Next
%>
  <tr> 
          <td nowrap><font size="2" face="Century Gothic">Storage Total:</font></td>
    <td align="right" style="border-right:1px solid #000000"><font size="2" face="Century Gothic">&nbsp;<%=formatnumber((( tss/1024)/1024)/1024,2)%> GB &nbsp;</font></td>
    <td align="right" style="border-right:1px solid #000000"><font size="2" face="Century Gothic">&nbsp;<%=formatnumber((( tfs/1024)/1024)/1024,2)%> GB &nbsp;</font></td>
  </tr>
</table>
      <%
	LoadPercent = 0
	cpuStatus = 0
	TotalMem = 0
	FreeMem = 0
%>
    </td>
<td valign="top"><table cellpadding="0" cellspacing="0">
<tr>
          <td><font size="2" face="Century Gothic">Hardware Details</font></td>
        </tr>
</table></td>
</tr>
</table>
