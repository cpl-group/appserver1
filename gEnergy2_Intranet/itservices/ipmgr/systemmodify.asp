<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
action=lcase(Request("modify"))

Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst=Server.CreateObject("ADODB.Recordset")
cnn1.Open getConnect(0,0,"dbCore")

Select Case action 
	Case "save system"
		
		strsql = "select id from systemsindex  where serial = '" & trim(Request("serial")) & "'"

		rst.Open strsql, cnn1, 0, 1, 1
		
		if rst.eof then 
			strsql = "Insert into systemsindex (serial, systemtype, processor, memory, harddrive, nic, video, monitor) "_
			& "values ("_
			& "'" & Request.Form("serial") & "', "_
			& "'" & Request.Form("systemtype") & "', "_
			& "'" & Request.Form("processor")  & "', "_
			& "'" & Request.Form("memory") & "', "_
			& "'" & Request.Form("harddrive") & "', "_
			& "'" & Request.Form("nic") & "', "_
			& "'" & Request.Form("video") & "', "_
			& "'" & Request.Form("monitor") & "')"

			cnn1.execute strsql
		end if
		rst.close	
	
  	Case "update system"
	  key=cint(Request.Form("key"))
	  
	  strsql = "Update systemsindex Set serial='" & Request.Form("serial") & "', systemtype='" & Request.Form("systemtype") & "', processor='" & Request.Form("processor") & "', memory='" & Request.Form("memory") & "', harddrive='" & Request.Form("harddrive") & "', nic='" & Request.Form("nic") & "', video='" & Request.Form("video") & "', monitor='" & Request.Form("monitor") & "', note='" & Request.Form("note") & "' where id="& key 
	  cnn1.execute strsql
	  
	Case "delete system" 
	  key = trim(request("key"))
	  strsql = "delete from systemsindex where id="& key
	  cnn1.execute strsql
	  strsql = "Update ipindex Set systemid=0 where systemid="& key 
	  cnn1.execute strsql	  
  	Case "update ip"
	  key=cint(Request.Form("key"))
	  
	  strsql = "Update ipindex Set userid='" & Request.Form("userid") & "', ip='" & Request.Form("ip") & "', ipname='" & Request.Form("ipname") & "', systemid='" & Request.Form("systemid") & "', adate='" & Request.Form("assigndate") & "' where id="& key 
	  cnn1.execute strsql
  	Case "delete ip" 
	  key = trim(request("key"))
	  strsql = "delete from ipindex where id="& key
	  cnn1.execute strsql
	  
  	Case "save ip"
	  key=cint(Request.Form("key"))
	  
		strsql = "select id from ipindex where ip = '" & Request.Form("ip") & "'"

		rst.Open strsql, cnn1, 0, 1, 1
		
		if rst.eof then 
			strsql = "Insert into ipindex (ip, ipname, systemid, userid, adate) "_
			& "values ("_
			& "'" & Request.Form("ip") & "', "_
			& "'" & Request.Form("ipname") & "', "_
			& "'" & Request.Form("systemid")  & "', "_
			& "'" & Request.Form("userid") & "', "_
			& "'" & Request.Form("assigndate") & "')"

			cnn1.execute strsql
		end if
		rst.close	


	End Select
	  set cnn1=nothing
	  
		 
	tmpMoveFrame =  "parent.frames.subnetlist.location = 'subnets.asp';parent.document.all.se.style.visibility='hidden';"

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>