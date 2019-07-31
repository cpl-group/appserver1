<%option explicit%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<HTML>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
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
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<%
dim pid, startdate, utility
pid  = request("pid")
if isempty(trim(pid)) then
	response.write "page was called without a valid pid."
	response.End
end if

utility = trim(request("utility"))
if isempty(utility) then
	utility = "2"
end if


startdate=Request("startdate")
if not(isdate(startdate)) then startdate=date()

dim cnnMainModule, cmd, rst
set rst = server.createobject("adodb.recordset")

set cnnMainModule = server.createobject("adodb.connection")
cnnMainModule.open getConnect(pid,0,"billing")

set cmd = server.createobject("ADODB.Command")
cmd.CommandText = "sp_PLMPDATA_BREAKDOWN"
cmd.CommandType = adCmdStoredProc
cmd.ActiveConnection = cnnMainModule

dim prm
Set prm = cmd.CreateParameter("from", adVarChar, adParamInput,12)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("to", adVarChar, adParamInput,30)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("utility", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("pid", adInteger, adParamInput)
cmd.Parameters.Append prm

cmd.Parameters("from") = startdate
cmd.Parameters("to") = startdate & " 23:59:00"
cmd.Parameters("utility") = utility
cmd.Parameters("pid") = pid

'response.write "sp_PLMPDATA_BREAKDOWN " & cmd.Parameters("from") & "," & cmd.Parameters("to") & "," & cmd.Parameters("utility") & "," &cmd.Parameters("pid")
'response.end
set rst = cmd.execute

'sp_PLMPDATA_BREAKDOWN  @from datetime, @to datetime,@utility int ,@pid int as
if not rst.eof then
	dim counter
	counter = 0
	%>
<script>
function openwin(url, wdth, hght){
	var w = wdth;
	var h = hght;
	var winl = (screen.width - w) / 2;
    var wint = (screen.height - h) / 2;

    winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',status=yes,scrollbars=yes,resizable=yes'
	popupwin=window.open(url,'lmpchrt',winprops)
	popupwin.focus('login')
}
function setlink(){
	if (document.all.headlink.innerHTML == 'View Contibution To Coincidental Peak Demand'){
		parent.lmp.document.location.href='peakDemandPie.asp?pid=<%=pid%>&fdate=<%=startdate%>&tdate=<%=startdate%>&utility=<%=utility%>'; 
	}else{
		parent.lmp.document.location.href='lmpload.asp?interval=0&utility=<%=utility%>&pid=<%=pid%>&startdate=<%=startdate%>'; 
	}
	
	document.all.headlink.innerHTML =  (document.all.headlink.innerHTML == "View Contibution To Coincidental Peak Demand" ? "View Portfolio Load" : "View Contibution To Coincidental Peak Demand");
	
}
</script>

	<table cellspacing="0" cellpadding="2" width = "100%">
		<tr bgcolor="#0099FF">
			
    <td colspan=4> <span class="standardheader"><font color='white'>Portfolio 
      Usage and Demand Breakdown | <a href="#" onclick="setlink();"><span id='headlink'>View 
      Contibution To Coincidental Peak Demand</span></a></font></span></td>
		<tr>	
		<tr bgcolor="#CCCCCC">
			<td width="70%"><span class="standardheader">Property</span></td>
			<td align="center"><span class="standardheader">Total Usage</span></td>
			<td align="center"><span class="standardheader">Peak Demand</span></td>
		</tr>
	</table>
	<div style="overflow:auto; height:85%">
	<table cellspacing="0" cellpadding="0" width = "100%"> 
		<%
		do while not rst.eof
			dim bgcolor
			if counter mod 2 = 1 then bgcolor = "white" else bgcolor = "#eeeeee"
			dim rstbldg, cnnLocal
			set rstbldg = server.createobject("adodb.recordset")
			set cnnLocal = server.createobject("adodb.connection")
			cnnLocal.open getLocalConnect(rst("bldgnum"))
			rstbldg.open "select strt from buildings where bldgnum = '" & rst("bldgnum") & "'", cnnLocal
			%>
			<tr bgcolor="<%=bgcolor%>">
	      <td width="70%"><a href="#" onclick="openwin('lmp.asp?bldg=<%=rst("bldgnum")%>', 800, 550)"><%=rstbldg("strt")%></a></td>
				<td align="right"><%=rst("usage")%></td>
				<td align="right"><%=rst("demand")%></td>
				<td width="10"></td>
			</tr>
			<%
			rst.movenext
			counter = counter + 1
		loop
	%></table><%
end if

%>
</div>	