<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
user="ghnet\"&trim(request("name"))
Response.Write value
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"intranet")

dim startweek, endweek
startweek = Request("startweek")
endweek = Request("endweek")

strsql = "Update user_cost Set startweek='" & startweek & "', endweek='" & endweek & "' where (username='"& user &"')"
'Response.Write strsql
'response.end

cnn1.execute strsql

set cnn1=nothing

dim userparam
userparam = trim(request("name"))
%>
<script>
parent.form1.start.value = "<%=startweek%>"
parent.form1.end1.value = "<%=endweek%>"
parent.tstop.window.document.location="timesheet.asp?name=<%=userparam%>"
window.document.location = "timedetail.asp?name=<%=userparam%>"

</script>