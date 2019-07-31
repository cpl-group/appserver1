<%@Language="VBScript"%>
<%
	
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"
id1=request.querystring("surveyid")
key= Request.Querystring("key")

		sqlstr = "delete tblsurveyitem where id=" & Request.Querystring("key") & ""
		
cnn1.Execute sqlstr 

tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "surveyitems.asp?surveyid="&id1 & chr(34) & vbCrLf 
				 
				  
				  
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 				  
%>