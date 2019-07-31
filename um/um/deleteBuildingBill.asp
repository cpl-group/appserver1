<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim pid, bldg, ypid, sql
pid = request.querystring("pid")
bldg = request.querystring("bldgnum")
ypid = request.querystring("ypid")

dim cnn1, rst1
set cnn1 = server.createobject("ADODB.Connection")
set rst1 = server.createobject("ADODB.Recordset")
cnn1.open application("cnnstr_genergy1")

if trim(ypid)<>"" then
	rst1.open "SELECT count(posted) FROM tblbillbyperiod WHERE ypid="&ypid&" and posted=1", cnn1
	if cint(rst1(0)) = 0 then
		sql = "DELETE FROM tblbillbyperiod WHERE ypid="&ypid
		cnn1.execute sql
	else
		response.write "Building already posted, delete <b style=""color:red;font-size:20;filter:DropShadow(color=#C0C0C0, offx=5, offy=5)""><blink>DENIED</blink></b>."
		response.end
	end if
end if

response.redirect "tenantbilllist.asp?bldg="&bldg&"&ypid="&ypid
%>
