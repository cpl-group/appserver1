<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim b, m, luid, byear, bperiod
b = request.querystring("b")
'b="20"
luid = request.querystring("luid")
byear = request.querystring("byear")
bperiod = request.querystring("bperiod")
m = request.querystring("m")

if trim(byear)="" or trim(bperiod)="" then
    byear=0
    bperiod=0
end if

Dim cnn, cmd, rs
Dim FLD 'As Field
Dim prm 'As Parameter
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")


' open connection
cnn.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
cnn.CursorLocation = adUseClient

' specify stored procedure to run
cmd.CommandType = adCmdStoredProc

'dim h
'h="xxxxxxxxx xxxxx"

'h = left(h&" ", instr(h&" "," ")-1)
'response.write h&"|"
'response.end

' assign internal name to stored procedure
cmd.Name = "test"
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

    Set cmd.ActiveConnection = cnn
    'return set to recordset rs
    cnn.test b, luid, byear, bperiod, rs
else
    cmd.CommandText = "sp_peak_contribution"
    ' set parameter type and append for building contribution pie
    Set prm = cmd.CreateParameter("bldg", adChar, adParamInput, 10)
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
	dim dmaxdate
	dmaxdate = left(cmd.Parameters("dmax")&" ",instr(cmd.Parameters("dmax")&" ", " ")-1)

    Set cmd.ActiveConnection = cnn
    'return set to recordset rs
    cnn.test b, byear, bperiod, rs
end if

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

<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000" onload="parent.lmp.document.location.href='peakDemandPieload2.asp?b=<%=b%>&luid=<%=luid%>&byear=<%=cmd.Parameters("by")%>&bperiod=<%=cmd.Parameters("bp")%>'; parent.closeLoadBox('loadFrame2')">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr bgcolor="#000000">
    <td width="46%" height="2"><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b>Contributions to Peak Demand</font></td>
    <td align="right"><font face="Arial, Helvetica, sans-serif" size="2"><a href="options2.asp?m=<%=m%>&b=<%=b%>&luid=<%=luid%>" style="text-decoration:none;color:white" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'" onclick="parent.lmp.document.location.href='lmpload2.asp?m=<%=m%>&b=<%=b%>&luid=<%=luid%>&lmp=1'"><b>Back To Options</b></a></font></td>
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
'response.write typename(rs)
'response.end

'if trim(luid)="" then ' loop for building
    Do Until rs.EOF
        response.write "<tr style=""cursor:hand"" bgcolor=""#cccccc""" 
        if cDBL(rs("percentage"))>0 and trim(luid)="" then
            response.write "onclick=""parent.lmp.document.location.href='peakdemandpieload2.asp?b="& b &"&byear="& byear &"&bperiod="& bperiod &"&explode=" 'this disables
            if cDBL(rs("percentage"))>0 then response.write rs.AbsolutePosition-1 'this assures that if the entry has 0% resulting in no pie piece, that it can't explode a pie piece
            response.write "';parent.openLoadBox('loadFrame1')"
        end if
        response.write """ onmouseover=""this.style.backgroundColor='lightgreen'"" onmouseout=""this.style.backgroundColor='cccccc'"">"
        response.write "<td width=""70%""><b><font size=""1"" face=""Arial"">"& "Demo Tenant" &"</font></b></td>"&_
                       "<td width=""15%"" align=""right""><b><font size=""1"" face=""Arial"">"& formatnumber(rs("demand"),1) &"</font></b></td>"&_
                       "<td width=""15%"" align=""right""><b><font size=""1"" face=""Arial"">"& formatnumber(rs("percentage")) &"%</font></b></td>"
        response.write "</tr>"
        rs.MoveNext
    Loop
'else 'loop for tenant
'    Do Until rs.EOF
'        response.write "<tr style=""cursor:hand"" bgcolor=""#cccccc"" onclick=""parent.lmp.document.location.href='peakdemandpieload2.asp?b="& b &"&byear="& byear &"&bperiod="& bperiod &"&explode="
'        if cDBL(rs("percentage"))>0 then response.write rs.AbsolutePosition-1 'this assures that if the entry has 0% resulting in no pie piece, that it can't explode a pie piece
'        response.write "';parent.openLoadBox('loadFrame1')"" onmouseover=""this.style.backgroundColor='lightgreen'"" onmouseout=""this.style.backgroundColor='cccccc'"">"
'        response.write "<td width=""70%""><b><font size=""1"" face=""Arial"">"& rs("labelname") &"</font></b></td>"&_
'                       "<td width=""15%"" align=""right""><b><font size=""1"" face=""Arial"">"& rs("demand") &"</font></b></td>"&_
'                       "<td width=""15%"" align=""right""><b><font size=""1"" face=""Arial"">"& formatnumber(rs("percentage")) &"%</font></b></td>"
'        response.write "</tr>"
'        rs.MoveNext
'    Loop
'end if
response.write "</table>"
Set cnn = Nothing
%>
</div>
<form name="PDpieposition">
<input name="byear" type="hidden" value="<%=cmd.Parameters("by")%>">
<input name="bperiod" type="hidden" value="<%=cmd.Parameters("bp")%>">
<input name="luid" type="hidden" value="<%=luid%>">
<input name="m" type="hidden" value="<%=m%>">
</form>
</body></html>
