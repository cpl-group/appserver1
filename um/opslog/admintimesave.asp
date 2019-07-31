<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%

user=("payroll")
Response.Write value
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"intranet")

strsql = "Update time_submission Set startweek='" & Request.querystring("startweek") & "', endweek='" & Request.querystring("endweek") & "' where (username='"& user &"')"

'Response.Write strsql
cnn1.execute strsql

set cnn1=nothing


%>
