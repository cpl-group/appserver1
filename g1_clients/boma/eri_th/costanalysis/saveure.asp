<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%option explicit
dim entrytype, desc, amt, period, pid, b, date1
entrytype=Request.Form("type")
desc=Request.Form("description")
amt=Request.Form("amt")
period=Request.Form("period")
pid = Request.Form("pid")
b=Request.Form("b")
date1=Request.Form("date1")

if entrytype = 0 and amt >= 0 then 

	amt = amt * -1
	
end if

dim cnn1, sql
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_genergy1")
sql = "Insert into tblRPentries (pid, description, amt, period,bldgnum, year, type) "_
	& "values ("_
	& "'" & pid & "', "_
	& "'" & desc & "', "_
	& "" & amt & ", "_
	& "'" & period & "', "_ 
	& "'" & b & "', "_
	& "'" & date1 & "', "_
	& "'" & entrytype & "')"
'response.write sql
'response.end

cnn1.execute sql
set cnn1=nothing

dim urltemp
urltemp = "unreported.asp?building=" & b & "&date1=" & date1 & "&pid=" & pid
dim tmpMoveFrame
tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				urltemp & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>