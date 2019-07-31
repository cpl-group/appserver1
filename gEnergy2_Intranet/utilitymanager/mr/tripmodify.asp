<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
action=lcase(Request("modify"))
action_sub = lcase(Request("modify_sub"))

if trim(action_sub) <> "" then 
	action = action_sub
end if

Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst=Server.CreateObject("ADODB.Recordset")

cnn1.Open application("cnnstr_palmserver")


Select Case action 
	Case "save"
		
		if request.form("applytoall") then 	
			strsql = "select distinct bldgnum from tripcodeindex where id = " & Request.Form("tripid") & " and utilityid = " & Request.Form("utility") & " and bldgnum <> '" & Request.Form("bldgnum") & "' and bldgnum not in (select bldgnum from tripcodeindex where billperiod = '" & Request.Form("billperiod") & "' and billyear = '" & Request.Form("billyear") & "' and id = " & Request.Form("tripid") & ")"
			rst.Open strsql, cnn1, 0, 1, 1		
			if not rst.eof then 
				while not rst.eof
					strsql = "Insert into tripcodeindex (id, bldgnum, utilityid, tripdate, billyear, billperiod) "_
					& "values ("_
					& "'" & Request.Form("tripid") & "', "_
					& "'" & rst("bldgnum") & "', "_
					& "'" & Request.Form("utility")  & "', "_
					& "'" & Request.Form("tripdate") & "', "_
					& "'" & Request.Form("billyear") & "', "_
					& "'" & Request.Form("billperiod") & "')"
				
					cnn1.execute strsql
				rst.movenext
				wend
			end if
		rst.close
		end if
		strsql = "select * from tripcodeindex where id = " & Request.Form("tripid") & " and utilityid = " & Request.Form("utility") & " and bldgnum = '" & Request.Form("bldgnum") & "' and billyear='" & Request.Form("billyear") & "' and billperiod = '"  & Request.Form("billperiod") & "'"

		rst.Open strsql, cnn1, 0, 1, 1
		
		if rst.eof then 
			strsql = "Insert into tripcodeindex (id, bldgnum, utilityid, tripdate, billyear, billperiod) "_
			& "values ("_
			& "'" & Request.Form("tripid") & "', "_
			& "'" & Request.Form("bldgnum") & "', "_
			& "'" & Request.Form("utility")  & "', "_
			& "'" & Request.Form("tripdate") & "', "_
			& "'" & Request.Form("billyear") & "', "_
			& "'" & Request.Form("billperiod") & "')"
			cnn1.execute strsql
		end if
		rst.close	
	
  	Case "update"
	  key=cint(Request.Form("key"))
	
	  strsql = "Update tripcodeindex Set id='" & Request.Form("tripid") & "', bldgnum='" & Request.Form("bldgnum") & "', utilityid='" & Request.Form("utility") & "', tripdate='" & Request.Form("tripdate") & "', billyear='" & Request.Form("billyear") & "', billperiod='" & Request.Form("billperiod") & "' where autoid="& key 
	
	  cnn1.execute strsql
	  
	Case "deletealltrips" 
	  tripcode=cint(Request("tripcode"))
	  tripdate = trim(request("tripdate"))
	  strsql = "delete from tripcodeindex where id="& tripcode & " and tripdate = '" &tripdate& "'"

	  cnn1.execute strsql
	Case "delete trip" 
	  key = trim(request("key"))
	  strsql = "delete from tripcodeindex where autoid="& key

	  cnn1.execute strsql	
	End Select
	  set cnn1=nothing
	  
		 
	tmpMoveFrame =  "parent.frames.tripsheet.location = 'tripcodes.asp';parent.document.all.te.style.visibility='hidden';"

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>