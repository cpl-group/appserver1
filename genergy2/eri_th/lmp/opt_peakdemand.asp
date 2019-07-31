<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldg, meterid, luid, byear, bperiod, startdate, billingid, utility
bldg = request.querystring("bldg")
luid = request.querystring("luid")
billingid = request.querystring("billingid")
utility = request.querystring("utility")
byear = request("byear")
startdate = request("startdate")
bperiod = request("bperiod")
meterid = request.querystring("meterid")
roleid = getkeyvalue("roleid")

byear = 2012
bperiod = 5

Dim cnn, cmd, rs,sqlstr, roleid,rst
Dim FLD 'As Field
Dim prm 'As Parameter
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")
set rst = server.createobject("ADODB.Recordset")

cnn.Open getLocalConnect(bldg)
cmd.CommandType = adCmdStoredProc
cnn.CursorLocation = adUseClient

if trim(byear)="" or trim(bperiod)="" then
    byear=0
    bperiod=0
end if



Set cmd.ActiveConnection = cnn

if trim(luid)<>"" then
    cmd.CommandText = "sp_peak_metercontribution"
    ' set parameter type and append for tenant contribution pie
    Set prm = cmd.CreateParameter("bldg", adChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("lid", adChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("byear", adinteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("bperiod", adinteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("max", adChar, adParamOutput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("by", adinteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("bp", adinteger, adParamOutput)
    cmd.Parameters.Append prm

    'return set to recordset rs
    cnn.test bldg, luid, byear, bperiod, rs
    cmd.Parameters("bldg") = bldg
    cmd.Parameters("utility") = utility
    cmd.Parameters("byear") = byear
    cmd.Parameters("bperiod") = bperiod
    'response.write "sp_peak_metercontribution '"&bldg&"', '"&utility&"', "&byear&", "&bperiod&",0,0,0"
    'response.end
    set rs = cmd.execute
else
    cmd.CommandText = "sp_peak_contribution"
    if getXMLUserName()="nyserda2" then cmd.CommandText = "sp_peak_contribution_MECH"
    ' set parameter type and append for building contribution pie
    Set prm = cmd.CreateParameter("bldg", adChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("utility", adInteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("byear", adinteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("bperiod", adinteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("max", adChar, adParamOutput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("dmax", adDBTimeStamp, adParamOutput, 11)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("tmax", adChar, adParamOutput, 8)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("by", adinteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("bp", adinteger, adParamOutput)
    cmd.Parameters.Append prm
    'response.write "exec sp_peak_contribution '"&bldg&"', '"&utility&"', "&byear&", "&bperiod&",0,0,0,0,0"
    'response.end
    cmd.Parameters("bldg") = bldg
    cmd.Parameters("utility") = utility
    cmd.Parameters("byear") = byear
    cmd.Parameters("bperiod") = bperiod
	'response.write cmd.CommandText
	'response.end
    set rs = cmd.execute
    
  	dim dmaxdate
	  dmaxdate = left(cmd.Parameters("dmax")&" ",instr(cmd.Parameters("dmax")&" ", " ")-1)
end if
bperiod =  cmd.Parameters("bp")
byear =  cmd.Parameters("by")
%>
<html>
<head>
<title></title>
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
<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000" onload="parent.lmp.document.location.href='peakDemandPieload.asp?bldg=<%=bldg%>&luid=<%=luid%>&byear=<%=byear%>&bperiod=<%=bperiod%>&utility=<%=utility%>'; parent.closeLoadBox('loadFrame2')">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr bgcolor="#000000">
    <td width="46%" height="2"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Contributions to Peak Demand</font></td>
    <td align="right"><font face="Arial, Helvetica, sans-serif" size="2"><a href="options.asp?meterid=<%=meterid%>&bldg=<%=bldg%>&luid=<%=luid%>&utility=<%=utility%>" style="text-decoration:none;color:white" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'" onclick="parent.lmp.document.location.href='lmpload.asp?meterid=<%=meterid%>&bldg=<%=bldg%>&billingid=<%=billingid%>&utility=<%=utility%>'"><b>Return To Options</b></a></font></td>
  </tr>
</table>
<table width="100%" border="0" cellspacing="0" cellpadding="0"><tr>
    <td height="22"><font face="Arial, Helvetica, sans-serif" size="1">This 
    
    
    
    <%
    if trim(luid)<>"" then
        response.write "tenant's peak demand for period "& bperiod &" of "& byear
    elseif cmd.Parameters("dmax")<>"1/1/1900" then
        response.write "building's peak demand for "& cmd.Parameters("dmax")
    else
        response.write "building's peak demand for period "& bperiod &" of "& byear
    end if
    %>
    
    is <%=cmd.Parameters("max")%> kw.</font></td>
</tr></table>

<table border="1" width="100%" height="1" cellspacing="0" cellpadding="0">
  <tr bgcolor="#0099FF"> 
    <td width="70%" align="center"><b><font size="1" face="Arial">
    <%if trim(luid)<>"" then
        response.write "Meter"
    else
        response.write "Tenant"
    end if%>
    </font></b></td>
    <td width="15%" align="center"><b><font face="Arial" size="1">Demand</font></b></td>
    <td width="15%" align="center"><b><font size="1" face="Arial">Percentage</font></b></td>
  </tr>
</table>
<div style="overflow:auto;height:226">
<table cellpadding="0" cellspacing="0" border="0" width="100%">
<%
if rs.state = adstateopen then 
    Do Until rs.EOF
        response.write "<tr style=""cursor:hand"" bgcolor=""#cccccc""" 
        if cDBL(rs("percentage"))>0 and trim(luid)="" then
            response.write "onclick=""parent.lmp.document.location.href='peakdemandpieload.asp?bldg="& bldg &"&byear="& byear &"&bperiod="& bperiod &"&utility="& utility &"&explode=" 'this disables
            if cDBL(rs("percentage"))>0 then response.write rs.AbsolutePosition-1 'this assures that if the entry has 0% resulting in no pie piece, that it can't explode a pie piece
            response.write "';parent.openLoadBox('loadFrame1')"
        end if
        response.write """ onmouseover=""this.style.backgroundColor='lightgreen'"" onmouseout=""this.style.backgroundColor='cccccc'"">"
        response.write "<td width=""70%""><b><font size=""1"" face=""Arial"">"& rs("labelname") &"</font></b></td>"&_
                       "<td width=""15%"" align=""right""><b><font size=""1"" face=""Arial"">"& formatnumber(rs("demand"),1) &"</font></b></td>"&_
                       "<td width=""15%"" align=""right""><b><font size=""1"" face=""Arial"">"& formatnumber(rs("percentage")) &"%</font></b></td>"
        response.write "</tr>"
        rs.MoveNext
    Loop
response.write "</table>"
end if
Set cnn = Nothing
%>
</div>
<form name="PDpieposition">
<input name="byear" type="hidden" value="<%=byear%>">
<input name="bperiod" type="hidden" value="<%=bperiod%>">
<input name="luid" type="hidden" value="<%=luid%>">
<input name="meterid" type="hidden" value="<%=meterid%>">
</form>
</body></html>
