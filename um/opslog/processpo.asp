<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<%
if not allowGroups("Genergy_Corp,Genergy Users") then 
	response.write "<br><div align=center>You don't have permissions for this function.</div>"
	response.end
end if

Dim cnn1, jid
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")

jid = request("jid")
caller = request("caller")

if jid <> "" then 
	Dim company,job
	rst2.Open "SELECT company, job FROM MASTER_JOB WHERE id='"&jid&"'", cnn1
	if not rst2.EOF then
		company = rst2("company")
		job = rst2("job")
	end if
	rst2.close
end if 
if getXmlUserName()<>"" then
	if request("poaction") = "submit" then
		sqlstr = "exec sp_po_submitted " & request("poid") & ", [" & getXmlUserName() & "]"
		response.write "<br><div align=center>PO has been submitted. All parties are being notified via email.</div>"
		if lcase(company) = "ge" then
			sendupdate 
		end if
	else
		sqlstr = "delete po where id=" & request("poid") & " delete po_item where poid=" & request("poid") & ""
		response.write "<br><div align=center>PO has been deleted.</div>"
	end if
	
	cnn1.Execute sqlstr 
end if
if caller = "joblog" then
	response.redirect "posearch.asp?caller=joblog&select=jobnum&findvar=" & right(job,4)%>
	<BR><BR><BR>
	<input type="button" value="Back to Job Log" onclick="javascript:document.location='posearch.asp?caller=joblog&select=jobnum&findvar=<%=right(job,4)%>';">		<%
end if
%>
<%
function sendupdate()
	
	emailarray = "tara_clark@genergy.com;leighton_greenidge@genergy.com;jose.cotto@genergy.com"
	subject = "New PO was submitted by " & getkeyvalue("fullname") & " for Job " & request("jobnum")
	masternote = "PO Details"&vbcrlf &"=========="&vbcrlf &"PO Number: " & request("jobnum")&"."&request("ponum") & vbcrlf & "Submitted by: " & getkeyvalue("fullname") &"("&date()&" "&time()&")"  & vbcrlf & vbcrlf & "PO Description:" & request("podesc")  & vbcrlf & vbcrlf &"PO Total Amount:" & formatcurrency(request("poamt"),2)  & vbcrlf & vbcrlf & "Job Address: " & request("jobaddr")& vbcrlf & "Shipping Address: " & request("shipaddr")
	sendmail emailarray,"GSA",subject, masternote

end function
%>
