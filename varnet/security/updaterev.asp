<%@Language="VBScript"%>
<%
action=Request("action")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=security;"
if action = 1 then 
	strsql = "Insert into revtrack (revdate, guid, pid, revdescriptor) "_
	& "values ("_
	& "'" & Request.Form("revdate") & "', "_
	& "'" & Request.Form("guid") & "', "_
	& "'" & Request.Form("pid") & "', "_
	& "'" & Request.Form("revdesc") & "')"
else
	if action = 0 then
	strsql = "delete from revtrack where id = '" & Request("id") & "'"
	else 
		strsql = "update revtrack set revdate='" & Request.Form("revdate") & "' ,   guid='" & Request.Form("guid") & "',   pid='" & Request.Form("pid") & "',  revdescriptor='" & Request.Form("revdesc") & "' where id = " & request("id")	
	
	end if

end if
cnn1.execute strsql

set cnn1=nothing
%>
<script>
opener.location = <%="'revtrack.asp?pid=" & Request.Form("pid") &"'"%>
window.close()
</script>



