<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%

Select case trim(request("mode"))

	Case "requestupdate"
		dim ticketid, uid, userTo
		ticketid = trim(request("ticketid")) 
		uid = getKeyValue("user")
		userTo = trim(request("mailto"))
		
		Dim MyBody
		Dim MyCDONTSMail
		Set MyCDONTSMail = CreateObject("CDONTS.NewMail")
		
		dim rsReqUp
		set rsReqUp = server.createObject("adodb.recordset")
		
		rsReqUp.open "SELECT empto.email AS toMail, empfrom.email AS fromMail FROM employees empto, employees empfrom WHERE (empto.username = 'ghnet\"&userTo&"') AND (empfrom.username = 'ghnet\"&getKeyValue("user")&"')", getConnect(0,0,"intranet")		
		MyCDONTSMail.From= rsReqUp("fromMail")
		MyCDONTSMail.To= rsReqUp("toMail")		
		rsReqUp.close
		
		rsReqUp.open "select * from tickets where id = '" & ticketId & "'", getConnect(0,0,"dbCore")
		
		MyCDONTSMail.Subject="Update requested for ticket " & ticketid
		MyBody = "Request update on Ticket " & ticketid & ".  Ticket due " & rsReqUp("duedate") & vbCrLf & vbCrLf
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
				
		MyCDONTSMail.Body= MyBody
		MyCDONTSMail.Send
		set MyCDONTSMail=nothing
		response.redirect "notemanage.asp?mode=save&ticketid=" & ticketid & "&note=" & server.urlencode(getKeyValue("user") & " requested update")
		response.end
  Case "new"
  		dim headerlabel, notefortype, notefor

'2/11/2008 N.AMbo amended variables
  		'ticketid = trim(request("ticketid")) 
  		notefortype =  trim(request("notefortype"))
  		notefor = trim(request("notefor"))
		headerlabel = trim(request("headerlabel"))

	%>
	
	<html>
	<head>
	<title>New Note for <%=headerlabel%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<script>
	function closepage(){
		if (confirm("Cancel changes?")){
		  document.location="./notemanage.asp?mode=new&notefortype=<%=notefortype%>&notefor=<%=notefor%>"
		}
	}
	</script>
	<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">   
	</head>
	<body bgcolor="#dddddd">
	<form name="form1" method="post" action="/genergy2/NOTES/notemanage.asp">
	
	<table border=0 cellpadding="3" cellspacing="0" width="100%">
		<tr valign="top" bgcolor="#6699cc">
			<td><span class="standardheader">New Note for <%=headerlabel%></span></td>
		</tr>
		<tr valign="middle" bgcolor="#eeeeee">
			<td>
				<textarea name="note" cols="100%" rows="5" id="note" ></textarea>
			</td>
		</tr>
		<tr bgcolor="#eeeeee">
			<td style="border-bottom:1px solid #cccccc;">
				Hours for note:
				<input type="text" size="1" width="1" name="time">
			</td>
		</tr>
		
    <tr> 
      <td><br>
				<input type="hidden" name="mode" value="save">
				<input type="hidden" name="child" value="<%=request("child")%>">		  
				<input type="hidden" name="notefortype" value="<%=request("notefortype")%>">
				<input type="hidden" name="notefor" value="<%=request("notefor")%>">
				<input name="submit" type="submit" value="Save">
				<input name="button" type="button" onClick="closepage();" value="Clear">
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
	
	'2/11/2008 N.AMbo amended variables
		'ticketid    = trim(request("ticketid")) 
		notefortype =  trim(request("notefortype"))
  		notefor = trim(request("notefor"))
		note 		= trim(request("note"))
		uid 		= getKeyValue("user") 
		notemode	= request("notemode")
		noteTime = request("time")

		if (not isNumeric(noteTime)) or isNull(noteTime) or noteTime = "" then
			noteTime = 0
			notefortype =  trim(request("notefortype"))
		end if
		set cnn = server.createobject("ADODB.connection")
		set rst = server.createobject("ADODB.recordset")
		cnn.open getConnect(0,0,"dbCore")
		note = replace(note, "'", "''")
		strsql = "insert into master_notes (notefortype, notefor,note, uid, [time]) values ('"&notefortype&"', '"&notefor&"', '"&note&"','"&uid&"',"&noteTime&")"
		cnn.Execute strsql
			
		%>
		<script>
		parent.document.location.reload()
		</script>
		<%
	Case "view"
	'ticketid = request("ticketid")
	notefor = trim(request("notefor"))
	headerlabel = trim(request("headerlabel"))
	%>
	
	<html>
	<head>
	<title>Note <%=Request("nid")%></title>
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
	<script>
	function closepage(){
		  document.location="./notemanage.asp?mode=new&notefortype=<%=notefortype%>&notefor=<%=notefor%>"
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
	<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">   
	</head>
	<body bgcolor="#dddddd">
	<%
	set cnn = server.createobject("ADODB.connection")
	set rst = server.createobject("ADODB.recordset")
	cnn.open getConnect(0,0,"dbCore")
	strsql = "select * from master_notes where id =" & request("nid")
	rst.open strsql, cnn
	if not rst.eof then 
		%>
		<form name="form1" method="post" action="/genergy2/NOTES/notemanage.asp">
		<table border=0 cellpadding="3" cellspacing="0" width="100%">
		
			<tr valign="top" bgcolor="#6699cc">
				<td><span class="standardheader">Note created by <%=rst("uid")%> on <%=rst("date")%></span></td>
			</tr>
			
			<tr valign="middle" bgcolor="#eeeeee">
				<td style="border-bottom:1px solid #cccccc;">
					<div style="width:100%; overflow:auto; height:100;border-bottom:1px solid #eeeeee;">
						<%=rst("note")%>
					</div>
				</td>
			</tr>
			<tr bgcolor="eeeeee">
				<td style="border-bottom:1px solid #cccccc;">
					Note time: <%=rst("time")%>
				</td>
			</tr>
			
		</table>
		<input name="button" type="button" onClick="closepage();" value="New Note">
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




