<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

sql = "Update buildings Set readgroup='"&Request.form("readcode")&"' where bldgnum='"& Request.form("bldgnum") & "'"

cnn1.execute sql

set cnn1=nothing

tmpMoveFrame =  "parent.frames.admin.location = " & Chr(34) & _
				  "buildingtc.asp?pid="& request.form("pid") & chr(34) & vbCrLf 
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf  
%>