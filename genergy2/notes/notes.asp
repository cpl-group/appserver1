<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Dim cnn, rst, sql,note, requester, department, userid, clientticket, ccuid, runticket, notefor, notefortype,action, ticketid,ticketfound,c,headerlabel, notecount

'2/11/2008 N.Ambo amended variables to reflect vlaues for ntoes rather than tickets
'ticketid =request("ticketid")
notefortype = request("notefortype")
notefor = request("notefor")
headerlabel = request("hlabel")+" " + notefor
action = request("action")
%>
<html>
<head>
	<title>Notes for <%=headerlabel%></title>
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
set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open getConnect(0,0,"dbCore")

'2/11/2008 N.Ambo amended statement to reflect new naster_notes table
			sql = "select * from master_notes where notefortype= '" &notefortype& "' and notefor= '" &notefor& "' order by date desc"
			rst.open  sql, cnn 
			notecount = 0
			if not rst.eof then 
			%>
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
			  <tr>
				<td bgcolor="6699cc"><span class="standardheader">Notes for <%=headerlabel%></span> 
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
					notecount = notecount + 1
					%>
					<tr bgcolor="#cccccc" valign="top" onMouseOver="this.style.backgroundColor = 'lightgreen'" style="cursor:hand" 
						onMouseOut="this.style.backgroundColor = '#cccccc'" onClick="javascript:document.frames.view.location = '/genergy2/notes/notemanage.asp?mode=view&nid=<%=rst("id")%>&notefortype=<%=notefortype%>&notefor=<%=notefor%>'" > 
						
						<td width="35%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=rst("date")%></td>
						<td width="55%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=left(rst("note"),36)%><%if len(rst("note"))>36 then%>...<%end if%></td>
						<td width="12%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><%=rst("uid")%></td>
					</tr>
					<% 
					rst.movenext
				wend
				rst.close
				%>  
			 	</table> 
			</div>
<div align="center"><b><em>Click any row above to vew the complete note</em></b> 
  <br>
  <br>
  <%end if %>
</div>
<iframe id="view" height="200" width="100%" frameborder="0" src="/genergy2/notes/notemanage.asp?mode=new&notefortype=<%=notefortype%>&notefor=<%=notefor%>&headerlabel=<%=headerlabel%>"></iframe>
	
<input name="Close Window" type="button" value="Close Window" onclick="opener.document.all.notecount.innerHTML='<%=notecount%>';window.close()">
</body>
</html>
<%

Set cnn = nothing
Set rst = nothing
%>