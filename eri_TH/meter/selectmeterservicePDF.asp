<%option explicit%>
<%
dim cnn1, rst1
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_genergy1")
rst1.open "SELECT * from buildings order by bldgnum", cnn1
%>


<select onchange="document.all['infoframe'].src='/eri_th/meter/meterservicesPDF.asp?bldg='+this.value">
<option value="">Select Building</option>
<%
do until rst1.eof
	%><option value="<%=server.urlencode(rst1("bldgnum"))%>">(<%=rst1("bldgnum")%>)<%=rst1("strt")%></option><%
	rst1.movenext
loop
%>
</select>
<iframe id="infoframe" width="100%" height="100%"></iframe>
