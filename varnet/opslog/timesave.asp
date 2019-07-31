<%@Language="VBScript"%>
<%

user="ghnet\"&trim(Session("login"))
Response.Write value
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_main")

strsql = "Update user_cost Set startweek='" & Request("startweek") & "', endweek='" & Request("endweek") & "' where (username='"& user &"')"

'Response.Write strsql
'response.end
cnn1.execute strsql

set cnn1=nothing


%>
