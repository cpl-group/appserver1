<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
		dim cnn, rs, sql
		
		set cnn = server.createobject("ADODB.Connection")
		set rs = server.createobject("ADODB.Recordset")
		' open connection
		cnn.open application("cnnstr_main")
		
%>
<html>
<head>
<title>Genergy Intranet</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="styles.css" type="text/css">
</head>
<script>
function openwin(url,mwidth,mheight){
cwin = window.open(url,"childwin","status=no, menubar=no,HEIGHT="+mheight+", WIDTH="+mwidth)
cwin.focus()
}
function deletetask(taskid){

		if (confirm("Delete Task?")){
		openwin('employeetasks.asp?mode=delete&taskid='+taskid,200,200)
		}
}
function deletedoc(docid){

		if (confirm("Delete Document Link?")){
		openwin('corporatedocs.asp?mode=delete&docid='+docid,200,200)
		}
}
function deletenews(docid, summary){

		if (confirm("Delete News Link : " + summary + "?")){
		openwin('corporatenews.asp?mode=delete&docid='+docid,200,200)
		}
}

</script>
<body bgcolor="#ffffff" topmargin="0">
<table border="0" width="100%" height="100%" cellpadding="3" cellspacing="1" bgcolor="#dddddd">
  <tr bgcolor="#6699cc"> 
    <td height="19" width="50%"><span class="standardheader">Personal Tasks</span></td>
    <td width="50%"><span class="standardheader">Open Trouble Tickets You're Associated 
      With </span></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff"> 
    <td height="50%"> 
      <%	
	sql = "select * from employee_todos where userid ='" & trim(session("user")) & "' and [complete%] <> 1 order by [duedate]"
	rs.open sql, cnn
	
	if not rs.eof then%>
      <table width="100%" border="0">
        <tr bgcolor="black">
          <td width="6%">&nbsp;</td>
          <td width="15%" align="center"><span class="standardheader">Due Date</span></td>
          <td width="52%"><span class="standardheader">Note</span></td>
          <td width="14%" align="center"><span class="standardheader">%</span></td>
        </tr>
      </table>
      <div style="width:100%; overflow:auto; height:80%;border-bottom:1px solid #cccccc;"> 
        <table width="100%" border="0">
          <%
		 Dim overduetickets 
 
		 while not rs.eof %>
          <tr>
            <td width="6%"><img src="../gEnergy2_Intranet/itservices/images/delete.gif" width="26" height="22" onclick="deletetask('<%=rs("id")%>')"></td>
            <td width="15%" align="right"> 
              <%if rs("duedate") < date() then %>
              <font color="#FF0000"><b>*</b></font> 
              <%
			  end if
			  %>
              <a href="javascript:openwin('employeetasks.asp?mode=view&taskid=<%=rs("id")%>',400,230)"><%=FormatDateTime(trim(rs("duedate")),2)%></a> 
            </td>
            <td width="52%"><%=left(rs("note"), 45)%>...</td>
            <td width="14%" align="right"><%if rs("complete%") > 0 then response.write formatpercent(rs("complete%"),0) else response.write "NA"%></td>
          </tr>
          <% 
		 rs.movenext
		 wend
		 %>
        </table>
      </div>
      <%
	    else 
	  		response.write "There are currently no tasks."
		end if 
 		 rs.close
 %>
      <div style="padding:3px;">
      <a href="javascript:openwin('employeetasks.asp?mode=new',400,230)">Add Task</a> 
      <br><font color="#FF0000">* = overdue ticket</font>
      </div>
      </td>
    <td height="50%"> 
      <%	
	sql = "select * from it_services.dbo.tickets where (userid ='" & trim(session("user")) & "' or requester = '" & trim(session("user")) & "' or ccuid = '" & trim(session("user")) & "') and closed = 0 order by [date]"
	rs.open sql, cnn
	if not rs.eof then%>
      <table width="100%" border="0">
        <tr bgcolor="black"> 
          <td width="18%" align="center"><span class="standardheader">Due Date</span></td>
          <td width="82%"><span class="standardheader">Note</span></td>
        </tr>
      </table>
      <div style="width:100%; overflow:auto; height:80%;border-bottom:1px solid #cccccc;"> 
        <table width="100%" border="0">
          <% 
		 while not rs.eof %>
          <tr> 
            <td width="18%" align="right">
              <%if rs("duedate") < date() then %>
              <font color="#FF0000">			  
			  <%
   				  if trim(rs("userid")) = trim(session("user")) then 
				  overduetickets = true
				  end if
			  
			  end if%>
              <%if rs("duedate") < date() then %>
              <b>*</b>
              <%end if%><a href="javascript:openwin('/genergy2_intranet/itservices/ttracker/ticket.asp?mode=update&tid=<%=rs("id")%>&child=1',660,270)"><%=FormatDateTime(trim(rs("duedate")),2)%></a>
              </font>
            </td>
            <td width="82%"><%if trim(rs("userid")) = trim(session("user")) then%><b><%end if%><%=Left(rs("initial_trouble"),57)%>...<%if rs("userid") = session("user") then%></b><%end if%></td>
          </tr>
          <% 
		 rs.movenext
		 wend
		 %>
        </table>
      </div>
      <%
	    else 
	  		response.write "There are currently no tasks."
		end if 
	  		 rs.close
%>
      <div style="padding:3px;">
      <a href="javascript:openwin('/genergy2_intranet/itservices/ttracker/ticket.asp?mode=new&child=1', 650, 250)">Add 
      Task</a> | <a href="/genergy2_intranet/itservices/ttracker/index.asp">Go 
      to Trouble Tickets</a> <br><font color="#FF0000">* = overdue ticket</font> | <b>bold type = assigned to you</b>
      </div>
      </td>
  </tr>
  <tr valign="top" bgcolor="#6699cc"> 
    <td height="18"><span class="standardheader">Corporate Documents</span></td>
    <td><span class="standardheader">Company News</span></td>
  </tr>
  <tr valign="top" bgcolor="#ffffff"> 
    <td height="50%"> 
      <%	
   if checkgroup("IT Services") or checkgroup("IT Consultants") or checkgroup("Department Supervisors") then 
	sql = "select * from corporate_docs order by name"
   else
	sql = "select * from corporate_docs where secure = 0 order by name"
    end if
	rs.open sql, cnn
	
	if not rs.eof then%>
      <table width="100%" border="0">
        <tr bgcolor="black">
       <%if checkgroup("Department Supervisors") then %>
		<td width="7%">;</td>
		<%end if%>
          <td width="15%"><div align="center"><span class="standardheader">Date</span></div></td>
          <td width="79%"><span class="standardheader">File</span></td>
        </tr>
      </table>
      <div style="width:100%; overflow:auto; height:80%;border-bottom:1px solid #cccccc;"> 
        <table width="100%" border="0">
          <%while not rs.eof %>
          <tr>
          <%if checkgroup("Department Supervisors") then %>
            <td width="7%"><img src="../gEnergy2_Intranet/itservices/images/delete.gif" width="26" height="22" onclick="deletedoc('<%=rs("id")%>')"></td>
		  <%end if%>	
            <td width="15%"><div align="center"><%=FormatDateTime(trim(rs("date")),2)%></div>
              </td>
			 <%if rs("weblink") = 0 then %>
            	<td width="79%"><a href="file:<%=rs("link")%>" target="_blank" title="<%=rs("desc")%>"><%=rs("name")%></a></td>
			 <% else %>
            	
            <td width="79%"><a href="<%=rs("link")%>" title="<%=rs("desc")%>" target="_blank"><%=rs("name")%></a> 
            </td>
			 <% end if %>
          </tr>
          <% 
		 rs.movenext
		 wend
		 %>
        </table>
      </div>
      <%
	    else 
	  		response.write "There are currently no corporate documents available."
		end if 
			 rs.close
  %>
      <%if checkgroup("Department Supervisors") then %>
      <a href="javascript:openwin('corporatedocs.asp?mode=new',400,230)">Add Document Link </a></td>
	  <%end if%>
    <td height="50%"> 
      <%	
   if checkgroup("Department Supervisors") then 
		sql = "select * from corporate_news where date >= '" & date() - 120 & "' order by date desc"
   else
		sql = "select * from corporate_news where date >= '" & date() - 120 & "' and secure = 0 order by date desc"
   end if


	rs.open sql, cnn
	
	if not rs.eof then%>
      <table width="100%" border="0">
        <tr bgcolor="black">
       <%if checkgroup("Department Supervisors") then %>
		<td width="7%">;</td>
		<%end if%>
          <td width="21%"><span class="standardheader">Date</span></td>
          <td width="64%"><span class="standardheader">News Summary</span></td>
        </tr>
      </table>
      <div style="width:100%; overflow:auto; height:80%;border-bottom:1px solid #cccccc;"> 
        <table width="100%" border="0">
          <% 
		 while not rs.eof %>
          <tr>
          <%if checkgroup("Department Supervisors") then %>
            <td width="7%"><img src="../gEnergy2_Intranet/itservices/images/delete.gif" width="26" height="22" onclick="deletenews('<%=rs("id")%>','<%=rs("summary")%>')"></td>
		  <%end if%>	
            <td width="21%"><%=FormatDateTime(trim(rs("date")),2)%></td>
            <td width="64%"><a href="javascript:openwin('corporatenews.asp?mode=view&id=<%=rs("id")%>',600,230)"><%=rs("summary")%></a> </td>
          </tr>
          <% 
		 rs.movenext
		 wend
		 set cnn = nothing
		 %>
        </table>
      </div>
      <%
	    else 
	  		response.write "No corporate new available."
		end if 
			 rs.close
  %>
         <%if checkgroup("Department Supervisors") then %>
		<a href="javascript:openwin('corporatenews.asp?mode=new',650,330)">Add New Corporate News</a>		
		<%end if%>
 </td>
  </tr>
</table>
</body>
<%if overduetickets=true then %>
<script>
if (confirm("Please review your open trouble tickets. You currently have some tickets open past due dates. Would you like to go to Trouble Tickets now?")){
	document.location = "/genergy2_intranet/itservices/ttracker/index.asp"
}
</script>
<%end if%>
</html>
