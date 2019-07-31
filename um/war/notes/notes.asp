<%option explicit%>
<%'2/14/ 2008 modified functionality of notes - now using new table master_notes %>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
	<title>Invoice Notes</title>
	<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
	<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<script>
function openwin(url,h,w) {

window.open(url,"window","scrollbars=yes,width=900,height=600,resizeable=no")
}
</script>
<body bgcolor="#dddddd">
<%
Dim Invoiceid, cnn, rst, sql,note, requester, department, userid, clientticket, ccuid, runticket, notefor, notefortype,action, ticketid,ticketfound,c,amanager, customer
dim custno,invoicedate,amount,Invnum,j
custno= request("custno")
invoicedate =request("invoicedate")
amount=request("amount") 

Invnum =request("Invnum")
j=request("j") 

notefortype = "arinvoice"


invoiceid=request("iid")
amanager = request("manager")
	if trim(amanager) = "" then
 		amanager = "ARManager"
	end if 
action = request("action")
c = request("c")
customer = request("customer")

set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open getConnect(0,0,"dbCore")

	'Check if Ticket is open for Invoice
	do
	
	'2/12/2008 N.Ambo amdemded SQL to reflect new notes table master_notes
	'sql = "select * from tickets where ticketfortype = 'arinvoice' and ticketfor = '" & invoiceid & "'" 
	sql = "select count(*) as counter from master_notes where notefortype = 'arinvoice' and notefor =  '" & invoiceid & "'" 
	
	rst.open sql, cnn
	
	if not rst.eof then 
		ticketfound = true
		'ticketid = rst("id")
		rst.close
	else
		'2/12/2008 N.Ambo removed because we are no longer using master tickets to record notes
		'No ticket found for invoice, initialize ticket in tickets table		
			'note = c&"-Invoice-"&invoiceid&"|Customer: " & customer
			'requester = amanager
			'department = "AR"
			'userid = amanager
			'clientticket = 1
			'ccuid  = ""
			'runticket = 1
			'notefor = invoiceid
			'notefortype = "arinvoice"
			'sql = "insert into tickets (initial_trouble, requester,department,userid, client,ccuid, runticket, ticketfor, ticketfortype) values ('"&trim(note)&"','"&trim(requester)&"','"&trim(department)&"','"& trim(userid)& "','" & trim(clientticket) & "','" & trim(ccuid) &"','" &trim(runticket)&"','" &trim(ticketfor)&"','" &trim(ticketfortype)& "')"
			'cnn.execute sql
			rst.close
	end if 
	loop  until ticketfound

			sql = "select * from master_notes where notefortype = 'arinvoice' and notefor = '" &invoiceid& "'order by date desc"
			rst.open  sql, cnn 
			
			if not rst.eof then 
			%>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			  <tr>
				<td bgcolor="6699cc"><span class="standardheader">Notes for Invoice <%=invoiceid%></span> 
				</td>
			  </tr>
			</table>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
					  <tr> 
							<td width="34%" >Date</td>
							<td width="55%" >Note</td>
							<td width="10%">uid</td>
							<td></td>
					  </tr>
			 </table>
			<div style="width:100%; overflow:auto; height:100;border-bottom:1px solid #cccccc;">
				<table width="100%" border="0" cellspacing="0" cellpadding="0">
			<%
				while not rst.eof
					%>
					<tr bgcolor="#cccccc" valign="top" onMouseOver="this.style.backgroundColor = 'lightgreen'" style="cursor:hand" 
						onMouseOut="this.style.backgroundColor = '#cccccc'" onClick="javascript:document.frames.view.location = './notemanage.asp?mode=view&nid=<%=rst("id")%>&notefortype=<%="arinvoice"%>&notefor=<%=invoiceid%>'" > 
						
						<td width="35%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=rst("date")%></td>
						<td width="55%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=left(rst("note"),36)%>...</td>
						<td width="12%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=rst("uid")%></td>
					</tr>
					<% 
					rst.movenext
				wend
				%>  
			 	</table> 
			</div>
<div align="center"><b><em>Click any row above to vew the complete note</em></b> 
  <br>
  <br>
  <%end if %>
</div>
<iframe id="view" height="200" width="100%" frameborder="0" src="./notemanage.asp?mode=new&notefortype=<%=notefortype%>&notefor=<%=invoiceid%>&custno=<%=custno%>&invoicedate=<%=invoicedate%>&amount=<%=amount%>&InvNum=<%=InvNum%>&j=<%=j%>&manager=<%=amanager%>"></iframe>
	
<input name="Close Window" type="button" value="Close Window" onclick="window.close()">
</body>
</html>
<%

Set cnn = nothing
Set rst = nothing
%>