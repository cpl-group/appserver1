<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldg, meterid, billingid, sd,ed, luid, utility, interval
bldg = request.querystring("bldg")
meterid = request.querystring("meterid")
billingid = request.querystring("billingid")
sd = request.querystring("sd")
ed = request.querystring("ed")
utility = request("utility")
interval = request("interval")
if isdate(ed) then ed = dateadd("n",-1,dateadd("d",1,ed))

Dim cnn, cmd, rs, prm
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")
cnn.CommandTimeout = 600*5
cnn.Open getLocalConnect(bldg)
cnn.CursorLocation = adUseClient

dim lmpid, lmptype
if trim(meterid)<>"" then
    lmptype="m"
    lmpid = meterid
elseif trim(billingid)<>"" then
    lmptype="L"
    rs.open "SELECT leaseutilityid FROM tblLeases l , tblleasesutilityprices lup WHERE l.billingid=lup.billingid AND lup.utility="&utility&" AND l.billingId="&billingid, cnn
    if not(rs.eof) then luid = cint(rs("leaseutilityid"))
    rs.close
    lmpid=luid
elseif trim(bldg)<>"" then
    lmptype="b"
    lmpid=bldg
end if

' set up stored proc
cmd.CommandType = adCmdStoredProc
Set cmd.ActiveConnection = cnn
' specify stored procedure to run
cmd.CommandText = "sp_download_v2"

' set parameter type and append
Set prm = cmd.CreateParameter("from", adVarChar, adParamInput, 25)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("to", adVarChar, adParamInput, 25)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("code", adChar, adParamInput, 2)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("string", adVarChar, adParamInput, 30)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("utility", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("interval", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("file", adVarChar, adParamOutput, 50)
cmd.Parameters.Append prm

cmd.Parameters("from") = sd
cmd.Parameters("to") = ed
cmd.Parameters("code") = lmptype
cmd.Parameters("string") = lmpid
cmd.Parameters("utility") = utility
cmd.Parameters("interval") = interval
'response.write "exec sp_download_v2 '"&cmd.Parameters("from")&"','"&cmd.Parameters("to")&"','"&cmd.Parameters("code")&"','"&cmd.Parameters("string")&"','"&cmd.Parameters("utility")&"','"&cmd.Parameters("interval")&"'"
'response.end
cmd.execute()

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
	<%if cmd.Parameters("file")<>"-1" then%>
      <div align="center"><font face="Arial, Helvetica, sans-serif"><a href="<%="/eri_th/sqldownload/" & cmd.Parameters("file")%> " style="text-decoration:none;" onMouseOver="this.style.color = '0099FF'" onMouseOut="this.style.color = 'white'"><font color="#FFFFFF"><b>Click Here to Download Data File</b></font></a></font></div>
	<%else%>
      <div align="center"><font face="Arial, Helvetica, sans-serif">The file was unable to create. Please try again or contact Genergy personel at <a href="mailto:support@genergy.com">support@genergy.com</a>.</font></div>
	<%end if%>
    </td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="18"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><b><a href="javascript:document.location='accesslmphistory.asp?<%="bldg="&bldg&"&meterid="&meterid&"&billingid="&billingid&"&utility="&utility%>'" style="text-decoration:none;" onMouseOver="this.style.color = '0099FF'" onMouseOut="this.style.color = 'black'">Return To View / Download Historical Profile</a></b></font></div>
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
        <a href="javascript:document.location='<%="options.asp?bldg="&bldg&"&meterid="&meterid&"&billingid="&billingid&"&utility="&utility%>'" style="text-decoration:none;" onMouseOver="this.style.color = '0099FF'" onMouseOut="this.style.color = 'black'">Return To Options</a></b></font></div>
    </td>
  </tr>
</table>
