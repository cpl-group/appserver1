<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Select case trim(request("mode"))
  Case "new"
  %>
    
    <html>
    <head>
    <title>New Personal Task</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <script>
    function closepage()
    {
      if (confirm("Cancel changes?")){
        window.close()
      }
    }

    </script>
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">   
    </head>
    <body bgcolor="#dddddd">
    <form name="form1" method="post" action="./employeetasks.asp">
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr valign="top" bgcolor="#6699cc"> 
      <td width="27%" colspan="2"><span class="standardheader">New Personal Task</span></td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;">Job Number </td>
      <td style="border-bottom:1px solid #cccccc;"><input type="text" name="jobnum">
        (optional) </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;">Due Date </td>
      <td style="border-bottom:1px solid #cccccc;"><input type="text" name="duedate">
        (optional) </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;">Completion % </td>
      <td style="border-bottom:1px solid #cccccc;"><select name="complete_per" id="select2">
          <option value="0">0</option>
          <option value=".25">25%</option>
          <option value=".5">50%</option>
          <option value=".75">75%</option>
          <option value="1">100%</option>
        </select>
        (optional) </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td valign="top" style="border-bottom:1px solid #cccccc;"> <p>Task Note</p></td>
      <td style="border-bottom:1px solid #cccccc;"><textarea name="note" cols="50" rows="5" id="textarea"></textarea></td>
    </tr>
    <tr> 
      <td><p> 
          <input type="hidden" name="mode" value="save">
          <input name="submit" type="submit" value="Save">
          <input name="button" type="button" onClick="closepage();" value="Cancel">
        </p></td>
      <td></td>
    </tr>
  </table>
    <br>
    </form>
    </body>
    </html>
  <%
  case "save"
    dim note, cnn, rst, strsql,jobnum, duedate, complete_per
	
            jobnum    = trim(request("jobnum")) 
			note 		= trim(request("note"))
			duedate		= trim(request("duedate"))
            complete_per	= request("complete_per")
            set cnn = server.createobject("ADODB.connection")
            set rst = server.createobject("ADODB.recordset")
            cnn.open getConnect(0,0,"intranet")'getConnect(0,0,"intranet")
			if duedate = "" then 
				duedate = date() + 7
			end if 
            
            strsql = "insert into employee_todos (userid,note,jobnum,duedate, [complete%]) values ('"&getKeyValue("user")&"', '"&note&"', '"&jobnum&"','"&duedate&"','"&complete_per&"')"
            cnn.Execute strsql
			%>
            <script>
			opener.document.location.reload()
			window.close()
            </script>
  <%
  Case "view"
  %>
    
    <html>
    <head>
    <title>Personal Tasks</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <script>
    function closepage()
    {
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
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">   
    </head>
    <body bgcolor="#dddddd">
    		  <%
      set cnn = server.createobject("ADODB.connection")
      set rst = server.createobject("ADODB.recordset")
      cnn.open getConnect(0,0,"intranet") 
	  
	  strsql = "select * from employee_todos where id ='"& request("taskid")&"' and userid ='" & getKeyValue("user") & "'"
	  
	  rst.open strsql, cnn
	  if not rst.eof then 
	  %>
    <form name="form1" method="post" action="./employeetasks.asp">
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr valign="top" bgcolor="#6699cc"> 
      <td width="27%" colspan="2"><span class="standardheader">Personal Task</span><input name="id" type="hidden" value="<%=rst("id")%>"></td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;">Job Number </td>
      <td style="border-bottom:1px solid #cccccc;"><input type="text" name="jobnum" value="<% if cint(rst("jobnum")) = 0 then %>NA<%else%><%=rst("jobnum")%><%end if%>">
        (optional) </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;">Due Date </td>
      <td style="border-bottom:1px solid #cccccc;"><input type="text" name="duedate" value="<%=rst("duedate")%>">
        (optional) </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;">Completion % </td>
      <td style="border-bottom:1px solid #cccccc;"><select name="complete_per" id="select2">
          <option value="0" <%if trim(rst("complete%")) = 0 then%> selected <%end if%>>NA</option>
          <option value=".25" <%if trim(rst("complete%")) = .25 then%> selected <%end if%>>25%</option>
          <option value=".5" <%if trim(rst("complete%")) = .5 then%> selected <%end if%>>50%</option>
          <option value=".75" <%if trim(rst("complete%")) = .75 then%> selected <%end if%>>75%</option>
          <option value="1" <%if trim(rst("complete%")) = 1 then%> selected <%end if%>>100%</option>
        </select>
        (optional) </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td valign="top" style="border-bottom:1px solid #cccccc;"> <p>Task Note</p></td>
      <td style="border-bottom:1px solid #cccccc;"><textarea name="note" cols="50" rows="5" id="textarea"><%=rst("note")%></textarea></td>
    </tr>
    <tr> 
      <td colspan=2>
          <div id="actionblock">
		  <input type="hidden" name="mode" value="update">
          <input name="submit" type="submit" value="Save">
          <input name="button" type="button" onClick="closepage();" value="Cancel">
		  <input name="button" type="button" onClick="createtimeentry('show');" value="Timesheet Entry">
		  </div>
		  <div id="hours" style="display:none;">
		  Enter Hours <input name="hrs" type="text" size="4" maxlength="3">
		  <input name="button" type="button" onClick="createtimeentry('save');" value="Save to My Timesheet">
		  <br>(note:due date, job number, and task note will be taken for timesheet entry)
		 </div>
      </td>
    </tr>
  </table>
    <br>
    
    </form>
	<%
	end if 
	rst.close
	%>
    </body>
    </html>

<%  
  case "update"
  dim tid
			tid = trim(request("id"))
            jobnum    = trim(request("jobnum")) 
			note 		= trim(request("note"))
			duedate		= request("duedate") 
            complete_per	= request("complete_per")
            set cnn = server.createobject("ADODB.connection")
            set rst = server.createobject("ADODB.recordset")
            cnn.open getConnect(0,0,"intranet")
            
            strsql = "update employee_todos set note='"& note &"',jobnum='"&jobnum&"',duedate='"&duedate&"', [complete%]='"&complete_per&"' where id = " & tid
            cnn.Execute strsql
			%>
            <script>
			opener.document.location.reload()
			window.close()
            </script>
  <%
    case "delete"
			tid = trim(request("taskid"))
            set cnn = server.createobject("ADODB.connection")
            set rst = server.createobject("ADODB.recordset")
            cnn.open getConnect(0,0,"intranet")
            
            strsql = "delete from employee_todos where id = " & tid
            cnn.Execute strsql
			%>
            <script>
			opener.document.location.reload()
			window.close()
            </script>
  <%

  case else
end select
%>




