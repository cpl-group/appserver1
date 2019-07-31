<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Select case trim(request("mode"))
  Case "new"
  %>
    
    <html>
    <head>
    <title>New Job Task</title>
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
    <form name="form1" method="post" action="./edittasks.asp">
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr valign="top" bgcolor="#6699cc"> 
      <td width="27%" colspan="2"><span class="standardheader">New Job Task</span></td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;">Job Number </td>
      <td style="border-bottom:1px solid #cccccc;"><%=trim(request("jobnum"))%></td>
       <input type="hidden" name="jobnum" value="<%=trim(request("jobnum"))%>" >
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;">Due Date </td>
      <td style="border-bottom:1px solid #cccccc;"><input type="text" name="duedate">
        (optional) </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;">Status</td>
      <td style="border-bottom:1px solid #cccccc;"><select name="status" id="Status1">
          <option value="Unstarted">Unstarted</option>
          <option value="In Progress">In Progress</option>
          <option value="Complete">Complete</option>         
        </select>
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
      <td valign="top" style="border-bottom:1px solid #cccccc;"> <p>Description</p></td>
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
  'Inserts new task record into master-job_tasks table. Defaults dues to date to 7 days from current day if left blank.
  case "save"
    dim note, cnn, rst, strsql,jobnum, duedate, complete_per, status
	
            jobnum    = trim(request("jobnum")) 
			note 		= trim(request("note"))
			duedate		= trim(request("duedate"))
            complete_per	= request("complete_per")
			status = trim(request("status"))
            set cnn = server.createobject("ADODB.connection")
            set rst = server.createobject("ADODB.recordset")
            cnn.open getConnect(0,0,"intranet")'getConnect(0,0,"intranet")
			if duedate = "" then 
				duedate = date() + 7
			end if 
            
            strsql = "insert into master_job_tasks(userid,description,jobid,duedate, percentcomplete, status) values ('"&getKeyValue("user")&"', '"&note&"', '"&jobnum&"','"&duedate&"','"&complete_per&"','"&status&"')"
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
    <title>Job Tasks</title>
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
	  
	  strsql = "select * from master_job_tasks where taskid ='"& request("taskid")&"' and userid ='" & getKeyValue("user") & "'"
	  
	  rst.open strsql, cnn
	  if not rst.eof then 
	  %>
    <form name="form1" method="post" action="./edittasks.asp">
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr valign="top" bgcolor="#6699cc"> 
      <td width="27%" colspan="2"><span class="standardheader">Job Task</span><input name="id" type="hidden" value="<%=rst("taskid")%>"></td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;">Job Number </td>
      <td style="border-bottom:1px solid #cccccc;"><%=trim(rst("jobid"))%></td>
      <input type="hidden" name="jobnum" value="<%=trim(rst("jobid"))%>" ID="Hidden1">

    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;">Due Date </td>
      <td style="border-bottom:1px solid #cccccc;"><input type="text" name="duedate" value="<%=rst("duedate")%>">
        (optional) </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;">Status</td>
      <td style="border-bottom:1px solid #cccccc;"><select name="status" >
          <option value="Unstarted" <%if trim(rst("status"))="Unstarted" then%> selected <%end if%>>Unstarted</option>
          <option value="In Progress" <%if trim(rst("status"))="In Progress" then%> selected <%end if%>>In Progress</option>
          <option value="Complete" <%if trim(rst("status"))="Complete" then%> selected <%end if%>>Complete</option>         
        </select>
        (optional) </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;">Completion % </td>
      <td style="border-bottom:1px solid #cccccc;"><select name="complete_per" id="select2">
          <option value="0" <%if trim(rst("percentcomplete")) = 0 then%> selected <%end if%>>NA</option>
          <option value=".25" <%if trim(rst("percentcomplete")) = .25 then%> selected <%end if%>>25%</option>
          <option value=".5" <%if trim(rst("percentcomplete")) = .5 then%> selected <%end if%>>50%</option>
          <option value=".75" <%if trim(rst("percentcomplete")) = .75 then%> selected <%end if%>>75%</option>
          <option value="1" <%if trim(rst("percentcomplete")) = 1 then%> selected <%end if%>>100%</option>
        </select>
        (optional) </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td valign="top" style="border-bottom:1px solid #cccccc;"> <p>Description</p></td>
      <td style="border-bottom:1px solid #cccccc;"><textarea name="note" cols="50" rows="5" id="textarea"><%=rst("description")%></textarea></td>
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
            status = request("status")
            set cnn = server.createobject("ADODB.connection")
            set rst = server.createobject("ADODB.recordset")
            cnn.open getConnect(0,0,"intranet")
            
            strsql = "update master_job_tasks set description='"& note &"',jobid='"&jobnum&"',duedate='"&duedate&"', percentcomplete="&complete_per&",  status='"&status&"' where taskid = " & tid
           'response.Write strsql
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
            
            strsql = "delete from master_job_tasks where taskid = " & tid
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




