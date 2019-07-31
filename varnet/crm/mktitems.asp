<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<html>
<head>
<script>
function updateEntry(id,mkid){
	parent.frames.mktdetail.location="mktdetail.asp?id="+id+"&mkid="+mkid
}
</script>

<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<%
id=request("key")

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set cmm1 = Server.CreateObject("ADODB.command")

cnn1.Open application("cnnstr_main")



sqlstr = "select * from mkt_progressitems where mid=" & Request.querystring("mkid") &" order by [date] desc"

if request("delete")="delete" then
    cmm1.commandText = "DELETE mkt_progressitems where id='"&id&"'"
    cmm1.commandType = adCmdText
    cmm1.ActiveConnection = cnn1
    cmm1.execute
    'response.write "delete"
    'response.end
end if

rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.eof then
%>

<table width="100%" border="0">
  <tr>
    <td>
      <div align="center"> 
        <p><font face="Arial, Helvetica, sans-serif"><i>No MKT Items Found </i></font></p>
        <hr>
        <p><font face="Arial, Helvetica, sans-serif"><i>New MKT Contact </i></font></p>
        </div>
    </td>
  </tr>
</table>
<%
else
%>
<table width="100%" border="0">
  <tr bgcolor="#CCCCCC"> 
    <td width="2%" height="2"> 
    <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Date</font></td>
    <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Action</font></td>
    <td width="18%" height="10"><font face="Arial, Helvetica, sans-serif">Comment</font></td>
    <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Follow-Up Date </font></td>
    <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Follow-Up Action</font></td>
    <td width="18%" height="10"><font face="Arial, Helvetica, sans-serif">Follow-Up Comment</font></td>
  </tr>
  <% While not rst1.EOF %>
  <form name="form1" method="post" action="mktitems.asp?mkid=<%=Request.querystring("mkid")%>">
    <tr valign="top"> 
      <input type="hidden" name="key" value="<%=rst1("id")%>">
      <input type="hidden" name="mkid" value="<%=Request.querystring("mkid")%>">
      <td width=6%> 
        <input type="button" name="edit" value="view/edit" size="7" onClick="updateEntry(key.value,mkid.value)">
        <a href="javascript:document.location.href='mktitems.asp?mkid=<%=Request.querystring("mkid")%>&key=<%=rst1("id")%>&delete=delete'"><img border="0" src="delete.gif"></a>
      </td>
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <%=rst1("date")%></font></td>
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <%=rst1("action")%> </font></td>
      <td width="18%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <%=rst1("comments")%></font></td>
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"><%=rst1("followupdate")%> 
        </font></td>
      <td width="11%" height="19" > <font face="Arial, Helvetica, sans-serif"><%=rst1("followup")%> 
        </font></td>
      <td width="18%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <%=rst1("fcomment")%></font></td>
    </tr>
  </form>
  <%
		rst1.movenext
		Wend
		%>
</table>
<%
end if
rst1.close
%>
</body>
</html>
