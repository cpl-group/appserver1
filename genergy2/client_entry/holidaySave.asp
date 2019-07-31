<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open application("cnnstr_genergy2")

dim rid, hid, action, holidaydate, holiday
rid = request("rid")
hid = request("hid")
holiday = request("holiday")
holidaydate = request("holidaydate")
action = request("action")

if trim(action)="Save" then
	strsql = "INSERT INTO rateholiday (holiday, date, regionid) values ('"&holiday&"', '"&holidaydate&"', '"&rid&"')"
else
	strsql = "UPDATE rateholiday set holiday='"&holiday&"', date='"&holidaydate&"' WHERE id="&hid
end if
'response.Write strsql
'response.End
cnn1.Execute strsql
if trim(action)="Save" then 'need to find the bid for the building just added
	rst1.Open "SELECT max(id) as id FROM rateholiday", cnn1
	if not rst1.eof then hid = rst1("id")
end if

Response.redirect "holidayEdit.asp?rid="&rid&"&hid="&hid
%>