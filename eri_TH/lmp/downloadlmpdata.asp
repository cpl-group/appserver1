<%option explicit%>
<!--#include file="adovbs.inc"-->
<%
dim b, m, sd,ed, luid
b = request.querystring("bldg")
m = request.querystring("m")
sd = request.querystring("sd")
ed = request.querystring("ed")
luid = request.querystring("luid")

Dim cnn, cmd, rs
Dim FLD 'As Field
Dim prm 'As Parameter
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")


' open connection
cnn.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
cnn.CursorLocation = adUseClient

' set up stored proc
cmd.CommandType = adCmdStoredProc
cmd.Name = "test"
Set cmd.ActiveConnection = cnn

if luid<>"" then
    ' specify stored procedure to run
    cmd.CommandText = "sp_download_luid"
    
    ' set parameter type and append
    Set prm = cmd.CreateParameter("luid", adChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("from", adChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("to", adChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("file", adChar, adParamOutput, 60)
    cmd.Parameters.Append prm

    cnn.test luid,sd,ed 
else
    ' specify stored procedure to run
    cmd.CommandText = "sp_download"
    
    ' set parameter type and append
    Set prm = cmd.CreateParameter("meterid", adChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("from", adChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("to", adChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("file", adChar, adParamOutput, 60)
    cmd.Parameters.Append prm
    
    cnn.test m,sd,ed 
end if


%>
<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#FFFFFF" onload="parent.closeLoadBox('loadFrame2')">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#000000"> 
      <div align="center"><font color="#FFFFFF" face="Arial, Helvetica, sans-serif"><b>Download 
        Data from <%=sd%> to <%=ed%></b></font></div>
    </td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor="#0099FF"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><a href="<%="../sqldownload/" & cmd.Parameters("file")%> " style="text-decoration:none;" onMouseOver="this.style.color = '0099FF'" onMouseOut="this.style.color = 'white'"><font color="#FFFFFF"><b>Click 
        Here to Download Data File</b></font></a></font></div>
    </td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="18"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><b><a href="javascript:document.location='accesslmphistory.asp'" style="text-decoration:none;" onMouseOver="this.style.color = '0099FF'" onMouseOut="this.style.color = 'black'">Return 
        To View / Download Historical Profile</a></b></font></div>
    </td>
  </tr>
  <tr>
    <td height="18">
      <div align="center">
        <hr width="100">
      </div>
    </td>
  </tr>
  <tr> 
    <td height="18"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><b> 
        <a href="javascript:document.location='<%="options2.asp?b=" & b & "&m=" & m %>'" style="text-decoration:none;" onMouseOver="this.style.color = '0099FF'" onMouseOut="this.style.color = 'black'">Return To Options</a></b></font></div>
    </td>
  </tr>
</table>
