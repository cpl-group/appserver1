<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<%
dim poid, user, message, status, cnn1, strsql
poid=secureRequest("poid")
user=Session("login")
if user="" or isnull(user) then user = getXMLUserName()
message=securerequest("message")
status=request("status")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"intranet")
strsql="sp_po_reject " & poid & "," & status & ",'" & user & "','" & message & "'"
'response.write strsql & "<BR>"
'response.write user
'response.end
cnn1.execute(strsql)
%>
<head>
<title>Requisition Forms</title>
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<script>
function reloadOpener(){
	if (opener != null) {
		opener.window.location.reload();
	}
}

</script>
</head>
<body bgcolor="#FFFFFF" onload="javascript:reloadOpener();">
<%
dim clr,msg
if status="Reject" then
	%>
	
	<table border=0 cellpadding="3" cellspacing="0" width="100%">
		<tr>
			<td bgcolor="#666699"><span class="standardheader">Requisition Forms</span></td>
		</tr>
		<tr>
			<td>
				The requisition form has been rejected.<br>
				<p><input type='button' value='Close Window' onclick='javascript:reloadOpener();window.close();'></p>
			</td>
		</tr>
	</table>
	
	<% 
else 

	dim jid, rst2
	jid = request("jid")
	Set rst2 = Server.CreateObject("ADODB.recordset")
	Dim company,job
	rst2.Open "SELECT company, job FROM MASTER_JOB WHERE id='"&jid&"'", cnn1
	
	if not rst2.EOF then
		company = rst2("company")
		job = rst2("job")
	end if
	
	rst2.close		
	%>
	
	The requisition form has been withdrawn.		<%
	
	if request("caller") = "joblog" then		
		response.redirect "posearch.asp?caller=joblog&select=jobnum&findvar=" & right(job,4)%>
		<br><br><br>
		<input type="button" value="Back to Job Log" onclick="javascript:document.location='posearch.asp?caller=joblog&select=jobnum&findvar=<%=right(job,4)%>';">		<%
	end if
end if	%>