<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
jobnum=secureRequest("jobnum")
podate=secureRequest("podate")
vid=secureRequest("vid")
jobadd=secureRequest("jobaddr")
shipadd=secureRequest("shipaddr")
req=secureRequest("req")
descr=secureRequest("description")
samt=secureRequest("ship_amt")
caller=secureRequest("caller")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")

strsql = "insert po (jobnum,podate,vid,jobaddr,shipaddr,requistioner,description)values ('" & jobnum& "', '" &podate & "', '" & vid & "', '" & jobadd & "', '" & shipadd & "','" & req & "', '" & descr & "')"
cnn1.execute strsql

strsql = "select max (id) as id from po"
rst.Open strsql, cnn1, 0, 1, 1


if not rst.eof then
	poid = rst("id")
end if

set cnn1=nothing



tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "poview.asp?poid="& poid & "&caller=" & caller & "&jid=" & jobnum & chr(34) & vbCrLf 
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>
