<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim UMinfo
Select case trim(request("mode"))

	Case "requestupdate"
		dim ticketid, uid, userTo,userFrom,MyBody, mySubject,sql
		ticketid = trim(request("ticketid")) 
		uid = getKeyValue("user")
		userTo = trim(request("mailto"))
		
		dim rsReqUp
		set rsReqUp = server.createObject("adodb.recordset")
		sql = "SELECT empto.email AS toMail, empfrom.email AS fromMail FROM ADusers_GenergyUsers empto, ADusers_GenergyUsers empfrom WHERE (empto.username = '"&userTo&"') AND (empfrom.username = '"&getKeyValue("user")&"')"
		
		rsReqUp.open sql, getConnect(0,0,"dbCore")
		
		
		if not rsReqUp.eof then 
			userFrom = rsReqUp("fromMail")
			userTo = rsReqUp("toMail")
		else
			userFrom = "GSA"
			userTo = "devteam@genergy.com"
		end if
		
		rsReqUp.close
		
		rsReqUp.open "select * from tickets where id = '" & ticketId & "'", getConnect(0,0,"dbCore")
		
		UMinfo = getUMInfo(rsReqUp("ticketfortype"), rsReqUp("ticketfor"), "0")
		mySubject="Update requested for ticket " & ticketid
		MyBody = "Request update on Ticket " & ticketid & ". "&UMinfo&" Ticket due " & rsReqUp("duedate") & vbCrLf & vbCrLf
		MyBody = MyBody & "Initial Problem: " & vbCrLf
		MyBody = MyBody & "   " & rsReqUp("initial_trouble") & vbCrLf & vbCrLf
		rsReqUp.close
		
		rsReqUp.open "select * from ttnotes where ticketid ='" & ticketId &"' order by date", getConnect(0,0,"dbCore")
		if rsReqUp.eof then
			MyBody = MyBody & "No notes for this ticket. " & vbCrLf
		else
			MyBody = MyBody & "Notes: " & vbCrLf
			do while not rsReqUp.eof
				MyBody = MyBody & "   " & rsReqUp("date") & ":  " & rsReqUp("note") & vbCrLf 
				rsReqUp.movenext
			loop
		end if
		rsReqUp.close
		set rsReqUp = nothing
		sendmail userTo,userFrom,mySubject, myBody

		response.redirect "note.asp?mode=save&ticketid=" & ticketid & "&note=" & server.urlencode(getKeyValue("user") & " requested update")
		response.end
	Case "new"
	%>
	
	<html>
	<head>
	<title>New Note</title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<script>
	function editcustomer(cid, company){
		pageURL = "cis_update.asp?cid=" + cid + "&company=" + company
		document.location = pageURL
		//window.resizeTo(600,300)
	}
	function closepage(){
		if (confirm("Cancel changes?")){
		  window.close()
		}
	}
	
	</script>
	<link rel="Stylesheet" href="../../styles.css" type="text/css">   
	</head>
	<body bgcolor="#dddddd">
	<form name="form1" method="post" action="note.asp">
	
	<table border=0 cellpadding="3" cellspacing="0" width="100%">
		<tr valign="top" bgcolor="#6699cc">
			<td><span class="standardheader">New Note</span></td>
		</tr>
		<tr valign="middle" bgcolor="#eeeeee">
			<td>
				<textarea name="note" cols="50" rows="5" id="note"></textarea>
			</td>
		</tr>
		<tr bgcolor="#eeeeee">
			<td style="border-bottom:1px solid #cccccc;">
				Hours for note:
				<input type="text" size="1" width="1" name="time">
			</td>
		</tr>
		<tr>
			<td> 
				<input type="checkbox" name="notemode" value="close">CLOSE TICKET<br>
				<input type="hidden" name="mode" value="save">
				<input type="hidden" name="child" value="<%=request("child")%>">		  
				<input type="hidden" name="ticketid" value="<%=request("ticketid")%>">
				<input name="submit" type="submit" value="Save">
				<input name="button" type="button" onClick="closepage();" value="Cancel">
				<input type = "hidden" name="ticketfor" value="<%=request("ticketfor")%>"> 
				<input type = "hidden" name="ticketfortype" value="<%=request("ticketfortype")%>"> 
			</td>
		</tr>
	</table>
	<br>
	
	</form>
	</body>
	</html>
	<%
	case "save"
		dim note, cnn, rst, strsql, notemode, noteTime
		
		ticketid    = trim(request("ticketid")) 
		note 		= trim(request("note"))
		uid 		= getKeyValue("user") 
		notemode	= request("notemode")
		noteTime = request("time")

		if (not isNumeric(noteTime)) or isNull(noteTime) or noteTime = "" then
			noteTime = 0
		end if
		set cnn = server.createobject("ADODB.connection")
		set rst = server.createobject("ADODB.recordset")
		cnn.open getConnect(0,0,"dbCore")
		note = replace(note, "'", "''")
		strsql = "insert into ttnotes (ticketid,note, uid, [time]) values ('"&ticketid&"', '"&note&"','"&uid&"',"&noteTime&")"
		cnn.Execute strsql
		
		%>
		<script>
			<%
			if trim(notemode) = "close" then 
				response.Write("opener.document.location='ticket.asp?mode=close&tid="&ticketid&"&child=" & request("child") & "&ticketfor="&request("ticketfor")&"&ticketfortype="&request("ticketfortype")&"'")
				
			else
				sendupdate ticketid
				response.Write("opener.document.location.reload()")
			end if
			%>
			window.close()
		</script>
		<%
  Case "transfer"
  %>
    
    <html>
    <head>
    <title>transfer ticket</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <script>
    function closepage()
    {
      if (confirm("Cancel changes?")){
        window.close()
      }
    }

    </script>
<link rel="Stylesheet" href="../../styles.css" type="text/css">   
    </head>
    <body bgcolor="#dddddd">
	<%
	dim cnnRequester, rstRequester, sqlRequester, requester, currentDueDate
	
	ticketid    = trim(request("ticketid")) 
	set cnnRequester = server.createobject("ADODB.connection")
	set rstRequester = server.createobject("ADODB.recordset")
	cnnRequester.open getConnect(0,0,"dbCore")
	sqlRequester = "SELECT requester, duedate FROM tickets WHERE id='" & ticketid & "'"
	rstRequester.open sqlRequester, cnnRequester
	
	if not rstRequester.eof then
		requester = rstRequester("requester")
		currentDueDate = rstRequester("duedate")
	end if
	
	%>
    <form name="form1" method="post" action="note.asp">
    
    <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr valign="top" bgcolor="#6699cc">
      <td><span class="standardheader">Transfer ticket <%=request("ticketid")%></span></td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee">
      <td style="border-bottom:1px solid #cccccc;">
         transfer ticket from <%=request("uid")%> to       
		   <% 
		  	dim userlist, usernamelist, UsersRS,tracktype
			set cnn 	= server.createobject("ADODB.connection")
			
			cnn.open getConnect(0,0,"dbCore")
			set UsersRS = server.createobject("ADODB.recordset")
			UsersRS.open "select * from adusers_genergyusers where email is not null order by company,department,fullname", cnn
			UsersRS.MoveFirst 
			GenerateUserList "uid",UsersRS,"","",getKeyValue("user")
			
			%>
			        </td>
    </tr>
		
	<%
	if allowGroups("Genergy_Corp") or getKeyValue("user") = requester then
		%>
	<tr bgcolor="#eeeeee">
		<td style="border-bottom:1px solid #cccccc;">
			Date (xx/xx/xxxx)<input name="newduedate" type="text" value="<%=currentDueDate%>" size="10" maxlength="10">
			
		</td>
	</tr>
		<%
	else
		%>
		<input name="newduedate" type="hidden" value="<%=currentDueDate%>">
		<%
	end if
	%>
	<tr><td><p>
			<input name="oldduedate" type="hidden" value="<%=currentDueDate%>">
          <input type="hidden" name="note" value="<%="transfer ticket from " & request("uid") & " to "%>">
          <input type="hidden" name="mode" value="savetransfer">
          <input type="hidden" name="ticketid" value="<%=request("ticketid")%>">
          <input name="submit" type="submit" value="Save">
          <input name="button" type="button" onClick="closepage();" value="Cancel">
        </p></td></tr>
    </table>
    <br>
    
    </form>
    </body>
    </html>
  <%
  case "savetransfer"
	dim newuid, newduedate, oldduedate
	newuid		= trim(request("uid"))
	ticketid    = trim(request("ticketid")) 
	note 		= trim(request("note")) & " " & newuid & "."
	response.write("note:" & note)
	
	newduedate	= trim(request("newduedate"))
	oldduedate	= trim(request("oldduedate"))
	if oldduedate <> newduedate then
		note = note & "  Due date changed from " & oldduedate & " to " & newduedate & "."
	end if
			
	uid 		=  getKeyValue("user")
	
	notemode	= request("notemode")
	set cnn = server.createobject("ADODB.connection")
	set rst = server.createobject("ADODB.recordset")
	cnn.open getConnect(0,0,"dbCore")
            
	strsql = "insert into ttnotes (ticketid,note, uid) values ('"&ticketid&"', '"&note & "','" &uid&"')"
	cnn.Execute strsql
	strsql = "Update tickets set userid = '" & newuid & "', duedate = '" & newduedate &"' where id = " & cdbl(ticketid)
	cnn.Execute strsql
	
	sendupdate ticketid
	%>
	<script>
	<%
	response.Write("opener.document.location.reload()")
	%>
	window.close()
	</script>
<%
	Case "view"
	%>
	
	<html>
	<head>
	<title>Note <%=Request("nid")%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<script>
	function closepage(){
		window.close()
	}
	
	function createtimeentry(action){
		if(action=='save'){
			var date=document.form1.duedate.value
			var jobnumber=document.form1.jobnum.value
			var description=document.form1.note.value
			var hrs=document.form1.hrs.value
			var name='<%=getKeyValue("user")%>'
			var url = '/um/opslog/timesheet-beta/timemodify.asp?modify=save&date='+date+'&job='+jobnumber+'&description='+description+'&hrs='+hrs+'&name='+name+'&source=personaltasks'
			document.location = url
		}else{
			document.getElementById("hours").style.display = "block";
			document.getElementById("actionblock").style.display = "none";
		}
	}
	</script>
	<link rel="Stylesheet" href="../../styles.css" type="text/css">   
	</head>
	<body bgcolor="#dddddd">
	<%
	set cnn = server.createobject("ADODB.connection")
	set rst = server.createobject("ADODB.recordset")
	cnn.open getConnect(0,0,"dbCore")
	strsql = "select * from ttnotes where id =" & request("nid")
	rst.open strsql, cnn
	if not rst.eof then 
		%>
		<form name="form1" method="post" action="note.asp">
		<table border=0 cellpadding="3" cellspacing="0" width="100%">
		
			<tr valign="top" bgcolor="#6699cc">
				<td><span class="standardheader">Note created by <%=rst("uid")%> on <%=rst("date")%></span></td>
			</tr>
			
			<tr valign="middle" bgcolor="#eeeeee">
				<td style="border-bottom:1px solid #cccccc;">
					<div style="width:100%; overflow:auto; height:100;border-bottom:1px solid #eeeeee;">
						<%=replace(rst("note"),vbcrlf,"<br>")%>
					</div>
				</td>
			</tr>
			<tr bgcolor="eeeeee">
				<td style="border-bottom:1px solid #cccccc;">
					Note time: <%=rst("time")%>
				</td>
			</tr>
			
		</table>
		<input name="button" type="button" onClick="closepage();" value="Close">
		<br>		
		</form>
		<%
	end if 
	rst.close
	%>
	</body>
	</html>
	
	<%  
  case else
end select
%>
<%
function sendupdate(ticket)
	
	dim cnn2,rs,requestor,assignedto,userarray, emailarray,dateopen,masternote,tixstatus,subject,sql,duedate
    set cnn2 = server.createobject("ADODB.connection")
    set rs 	 = server.createobject("ADODB.recordset")
    cnn2.open getConnect(0,0,"intranet")

	if ticket = "new" then 
		sql = "select top 1 * from tickets order by id desc" 
	else
		sql = "select * from tickets where id =" & ticket
	end if
	
	
	rs.open sql, cnn
	if not rs.eof then 
	UMinfo = getUMInfo(rs("ticketfortype"), rs("ticketfor"),"0")
	requestor  = rs("requester")
	assignedto = rs("userid")
	userarray  = rs("ccuid")
	dateopen   = rs("date")
	duedate    = rs("duedate")
	masternote = "[Original Note: "&rs("initial_trouble") &" submitted by "&requestor&" and currently assigned to "&assignedto&"]"
	ticket 	   = rs("id")
	
	if rs("closed") then 
		tixstatus = "Ticket Closed"
	else
		tixstatus = "Open Ticket"
	end if
	
	end if
	rs.close
	
	userarray  = userarray & ";" & requestor &";"&assignedto
	userarray  = split(userarray,";")
	
	for each x in userarray 
		rs.open "select email from ADusers_GenergyUsers where username = '" & trim(x) & "'", getConnect(0,0,"dbCore") 
		if not rs.eof then 
			if emailarray <> "" then 
			emailarray = emailarray & ";" & rs("email")
			else
			emailarray = rs("email")			
			end if 
		end if
		rs.close		
	next
	sql = "select * from ttnotes where ticketid = " & ticket & " order by date desc"
	rs.open sql, cnn
	if not rs.eof then 
		
		while not rs.eof
			masternote = masternote & vbCrLf & vbCrLf & " On "& rs("date") & " User " & rs("uid") &" added: " & rs("note") & vbCrLf 
		rs.movenext
		wend
	
	end if
	rs.close
	
	subject = tixstatus & ":" & ticket &" (Opened on:"&dateopen&")"
	sendmail emailarray,"GSA",subject, UMinfo& vbcrlf &masternote

end function
%>
<!--#INCLUDE VIRTUAL="/genergy2_intranet/itservices/ttracker/UMInfo.asp"-->



