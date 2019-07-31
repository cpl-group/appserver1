<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
bldg	=	Request("bldg")
budget = 	Request("budget")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getLocalConnect(bldg)

sql = "insert into BudgetsByBuilding (annualamt, bldgid) values (" & clng(budget) & ",'" &  bldg & "')"

cnn1.execute sql
set cnn1=nothing

Response.redirect "null.htm"
%>