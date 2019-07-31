<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
action=Request.Form("modify")
value=Request.Form("value")
value=Formatcurrency(value, 2)
user="ghnet\"&Session("login")
if (Request.Form("job")="") then
	Response.Redirect  "timedetail.asp"
end if
job=Request.Form("job")

id=Request.Form("id")
temp=Request.Form("date")
description=Request.Form("description")
Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst=Server.CreateObject("ADODB.Recordset")
cnn1.Open getConnect(0,0,"intranet")

if(action = "Save") then
    if (Request.Form("hrs")=0) then
		Response.Redirect  "timedetail.asp?day="&temp&"&job="&job
	end if
	
	'See if this is first time entered against this job
	rst.open	"select top 1 id from times where jobno='"&job&"'",cnn1
	if rst.eof then ' job is new
	  cnn1.execute "update master_job set last_invoice='"&(cdate(Request.Form("date"))-1)&"' where id="&job
	end if
	rst.close
	set rst= Nothing
		
	strsql = "Insert into Times (date, jobno, description, hours, overt, expense, value, matricola) "_
	& "values ("_
	& "'" & Request.Form("date") & "', "_
	& "'" & Request.Form("job") & "', "_
	& "'" & description & "', "_
	& "'" & Request.Form("hrs") & "', "_
	& "'" & Request.Form("ot") & "', "_
	& "'" & Request.Form("exp") & "', "& value &" ,"_
	& "'" & user & "')"

    'Response.Write strsql	
    cnn1.execute strsql


else		
  id=Request.Form("id")
  if (Request.Form("hrs")=0) then
	Response.Redirect  "timedetail.asp?day="&temp&"&id="&id
  end if	
  strsql = "Update Times Set date='" & Request.Form("date") & "', jobno='" & Request.Form("job") & "', description='" & description & "', hours='" & Request.Form("hrs") & "', overt='" & Request.Form("ot") & "', expense='" & Request.Form("exp") & "', value=" & value & " where (matricola='"& user &"' and id='"& id &"')"

  'Response.Write strsql
  cnn1.execute strsql

  'set cnn1=nothing

end if
tmpMoveFrame =  "parent.location = " & Chr(34) & _
				  "time.asp" & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>