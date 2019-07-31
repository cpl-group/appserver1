<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldg, pid, readcode
bldg = request("bldg")
pid = request("pid")
readcode = request("readcode")

dim cnn1, sql
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getLocalConnect(bldg)

sql = "UPDATE buildings SET readgroup='"&readcode&"' WHERE bldgnum='"&bldg&"'"
cnn1.execute sql
response.redirect "buildingtc.asp?pid="&pid
%>