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
    <form name="form1" method="post" action="./corporatenews.asp">
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr valign="top" bgcolor="#6699cc"> 
      <td width="27%" colspan="2"><span class="standardheader">New Corporate News</span></td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;"> Summary</td>
      <td style="border-bottom:1px solid #cccccc;"><input name="summary" type="text" size="100%"> 
      </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td valign="top" style="border-bottom:1px solid #cccccc;">Details</td>
      <td style="border-bottom:1px solid #cccccc;"><textarea name="details" rows="3" cols="100%"></textarea></td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;">File Name</td>
      <td style="border-bottom:1px solid #cccccc;"><input type="text" name="filename">
        (optional) </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td style="border-bottom:1px solid #cccccc;">File Link</td>
      <td style="border-bottom:1px solid #cccccc;"><input type="File" name="filelink">
        (optional) </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee"> 
      <td valign="top" style="border-bottom:1px solid #cccccc;">File Description</td>
      <td style="border-bottom:1px solid #cccccc;"><textarea name="desc" rows="3"></textarea>
        (optional) </td>
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
    dim note, cnn, rst, strsql,filename, filelink, secure,weblink,desc, summary,details
	
            filename    = trim(request("filename")) 
			filelink 	= trim(request("filelink"))
			desc		= trim(request("desc"))
			weblink		= trim(request("weblink"))
			secure		= trim(request("secure"))
			summary		= trim(request("summary"))
			details		= trim(request("details"))
			set cnn = server.createobject("ADODB.connection")
            set rst = server.createobject("ADODB.recordset")
            cnn.open application("cnnstr_main")
			
			if len(filename) <> 0 then  
				if weblink <> "1" then 
					if filelink = "" or filename = ""  or lcase(left(filelink,2)) <> "g:" then 
						response.write "File Link Failure: either File name or File link was not provided or the file was linked to a location other then g:\ or an HTTP address [<a href='javascript:history.back()'>back</a>]"
						response.end
					else
						filelink = replace(lcase(filelink),"g:","\\10.0.7.2\genergy")
					end if 
				 end if			
			end if
			if summary = "" or details = "" then 
				response.write "News Post Failure: either Summary or Details was not provided [<a href='javascript:history.back()'>back</a>]"
				response.end
			end if 
			            
            strsql = "insert into corporate_news (userid,filename,filelink,secure,weblink,[filedesc], summary, details) values ('"&session("user")&"', '"&filename&"', '"&filelink&"','"&secure&"','"&weblink&"','"&desc&"','" & summary &"','" & details &"')"
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
    <title>Corporate News</title>
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
	  
	  strsql = "select * from corporate_news where id ='"& request("id")&"'"
	  
	  rst.open strsql, cnn
	  if not rst.eof then 
	  %>
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr valign="top" bgcolor="#6699cc"> 
      <td width="27%" colspan="2"><span class="standardheader">Corporate news posted <%=FormatDateTime(rst("date"))%>
	  	<% if trim(rst("filename")) <> "" then %> 
		 <%if rst("weblink") = 0 then %>
			. This post contains a file attachment [<a href="file:<%=rst("filelink")%>" target="_blank" title="<%=rst("filedesc")%>"><%=rst("filename")%></a>]
		 <% else %>
			. This post contains a web attachment [<a href="<%=rst("filelink")%>" title="<%=rst("filedesc")%>" target="_blank"><%=rst("filename")%></a>]
		 <% end if 
		end if	
	%>
</td>
    </tr>
	<tr>
	<td>
	<%=rst("summary")%> 
	</td>
	</tr>
  </table>
  <div style="width:100%; overflow:auto; height:70%;border-bottom:1px solid #cccccc;"> 
  <table border=0 cellpadding="3" cellspacing="0" height="100%" width="100%" bgcolor="#eeeeee" style="border-bottom:1px solid #cccccc;">
    <tr valign="Top"> 
      <td><%=rst("details")%></td>
    </tr>
  </table> 
  </div>
	[<a href="javascript:window.close()" title="Close Window">exit</a>]   	<%
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
            
            strsql = "delete from corporate_news where id = " & tid
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




