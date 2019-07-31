<%@Language="VBScript"%>
<%

user=("payroll")
Response.Write value
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"

strsql = "Update time_submission Set startweek='" & Request.querystring("startweek") & "', endweek='" & Request.querystring("endweek") & "' where (username='"& user &"')"

'Response.Write strsql
cnn1.execute strsql

set cnn1=nothing


%>
