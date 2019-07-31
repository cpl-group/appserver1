<%
dim rsGetPulse
set rsGetPulse = server.createobject("ADODB.recordset")
function getMeterSource(meterid)
	rsGetPulse.open "SELECT datasource FROM meters m WHERE meterid="&meterid, application("cnnstr_genergy2")
	if not rsGetPulse.eof then getMeterSource = rsGetPulse("datasource")
	rsGetPulse.close
end function
%>
