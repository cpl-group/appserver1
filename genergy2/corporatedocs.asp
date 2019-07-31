<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Select case trim(request("mode"))
  Case "new"
  %>
    
    <html>
    <head>
    <title>Corporate Documents</title>
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
    <form name="form1" method="post" action="./corporatedocs.asp">
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr valign="top" bgcolor="#6699cc"> 
      <td width="27%" colspan="2"><span class="standardheader">New Corporate Document</span></td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;">File Name</td>
      <td style="border-bottom:1px solid #cccccc;"><input type="text" name="filename"> 
      </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;">File Link</td>
      <td style="border-bottom:1px solid #cccccc;"><input type="File" name="filelink"> 
      </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee">
      <td valign="top" style="border-bottom:1px solid #cccccc;">File Description</td>
      <td style="border-bottom:1px solid #cccccc;"><textarea name="desc" rows="3"></textarea></td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td valign="top" style="border-bottom:1px solid #cccccc;">Web Link</td>
      <td style="border-bottom:1px solid #cccccc;"><input name="weblink" type="checkbox" value="1">
        (optional) </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td valign="top" style="border-bottom:1px solid #cccccc;">Secure</td>
      <td style="border-bottom:1px solid #cccccc;"><input name="secure" type="checkbox" value="1">
        (optional) </td>
    </tr>
    <tr> 
      <td colspan = 2> Files may only be linked from G:\ or via an HTTP address</td>
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
    dim note, cnn, rst, strsql,filename, filelink, secure,weblink,desc
	
            filename    = trim(request("filename")) 
			filelink 	= trim(request("filelink"))
			desc		= trim(request("desc"))
			weblink		= trim(request("weblink"))
			secure		= trim(request("secure"))
			set cnn = server.createobject("ADODB.connection")
            set rst = server.createobject("ADODB.recordset")
            cnn.open application("cnnstr_main")
			
			if weblink <> "1" then  
				if filelink = "" or filename = ""  or lcase(left(filelink,2)) <> "g:" then 
					response.write "File Link Failure: either File name or File link was not provided or the file was linked to a location other then g:\ or an HTTP address [<a href='javascript:history.back()'>back</a>]"
					response.end
				else
					filelink = replace(lcase(filelink),"g:","\\10.0.7.2\genergy")
				end if 
			end if
			            
            strsql = "insert into corporate_docs (userid,name,link,secure,weblink,[desc]) values ('"&session("user")&"', '"&filename&"', '"&filelink&"','"&secure&"','"&weblink&"','"&desc&"')"
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

    </script>
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">   
    </head>
    <body bgcolor="#dddddd">
    		  <%
      set cnn = server.createobject("ADODB.connection")
      set rst = server.createobject("ADODB.recordset")
      cnn.open application("cnnstr_main")
	  
	  strsql = "select * from employee_todos where id ='"& request("taskid")&"' and userid ='" & session("user") & "'"
	  
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
      <td style="border-bottom:1px solid #cccccc;"><input type="text" name="jobnum" value="<%=rst("jobnum")%>">
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
      <td><p> 
          <input type="hidden" name="mode" value="update">
          <input name="submit" type="submit" value="Save">
          <input name="button" type="button" onClick="closepage();" value="Cancel">
        </p></td>
      <td>&nbsp;</td>
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
            cnn.open application("cnnstr_main")
            
            strsql = "update employee_todos set note='"& note &"',jobnum='"&jobnum&"',duedate='"&duedate&"', [complete%]='"&complete_per&"' where id = " & tid
            cnn.Execute strsql
			%>
            <script>
			opener.document.location.reload()
			window.close()
            </script>
  <%
    case "delete"
			tid = trim(request("docid"))
            set cnn = server.createobject("ADODB.connection")
            set rst = server.createobject("ADODB.recordset")
            cnn.open application("cnnstr_main")
            
            strsql = "delete from corporate_docs where id = " & tid
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




