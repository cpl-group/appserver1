<%@Language="VBScript"%>
<%option explicit
	dim b, pid, date1, action, iframeurl
	b=Request.QueryString("building")
	if b="" then Request.QueryString("b")
	pid = Request.QueryString("pid")
	date1=Request.QueryString("date1")
	action=Request.Querystring("action")
	iframeurl = "listure.asp?b=" & b & "&pid="&pid&"&date1="&date1
%>
<html>
<head>
<title>Adjustments</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head><style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>

<script>
function deleteentry(entryid,b,pid,date1){
	var temp="deleteure.asp?eid="+entryid+"&b="+b+"&pid="+pid+"&date1="+date1
	document.location=temp
}
</script>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#0099FF"> <font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif">Adjustments 
      : Expenses &amp; Revenue </font></b></font></td>
  </tr>
</table>
<p><font face="Arial, Helvetica, sans-serif"><b>Current Entries</b></font></p>
<IFRAME name="current" width="100%" height="150" src="<%=iframeurl%>"  scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor="#CCCCCC"><i><font face="Arial, Helvetica, sans-serif"> 
      <% if action="edit" then Response.write "Edit" else Response.write "New Entry" end if %>
      </font></i></td>
  </tr>
  <tr> 
    <td> 
      <% if action="edit" then %> <form name="form1" method="post" action="updateure.asp"><% else %> <form name="form1" method="post" action="saveure.asp"><% end if %>
	  <%if action="edit" then 
			dim rs, cnn1, sql, amt, newurl
			Set rs = Server.CreateObject("ADODB.recordset")
			Set cnn1 = Server.CreateObject("ADODB.Connection")
			cnn1.Open application("cnnstr_genergy1")
			
			sql = "select * from tblRPentries where id=" & Request.QueryString("entryid")
			
				rs.Open sql, cnn1, 0, 1, 1
			if rs.EOF then %>
			NO DATA FOUND
			<%else
			if rs("amt") < 0 then 
					amt = rs("amt") * -1
			else
					amt = rs("amt")
			end if
			%>
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td> <font face="Arial, Helvetica, sans-serif"> 
			  <input type="hidden" name="entryid" value="<%=Request.QueryString("entryid")%>">
			  <input type="hidden" name="b" value="<%=b%>">
			  <input type="hidden" name="pid" value="<%=pid%>">
			  <input type="hidden" name="date1" value="<%=date1%>">
              <input type="radio" name="type" value="0" <%if not rs("type") then Response.write "checked" end if %>>
              Expense 
              <input type="radio" name="type" value="1" <%if rs("type") then Response.write "checked" end if %>>
              Revenue </font></td>
          </tr>
          <tr> 
            <td><font face="Arial, Helvetica, sans-serif">Description (max 150 
              characters)</font></td>
          </tr>
          <tr> 
            <td> <font face="Arial, Helvetica, sans-serif"> 
              <textarea name="description" cols="50" rows="5" wrap="PHYSICAL"><%=rs("description")%></textarea>
              </font></td>
          </tr>
          <tr>
            <td><font face="Arial, Helvetica, sans-serif">Amount $ 
              <input type="text" name="amt" size="15" maxlength="15" value="<%=clng(amt)%>">
              Period 
              <select name="period">
                <option value="0" <%if rs("period") = 0 then Response.write "selected" end if %>>date1</option>
                <option value="1" <%if rs("period") = 1 then Response.write "selected" end if %>>Period 1</option>
                <option value="2" <%if rs("period") = 2 then Response.write "selected" end if %>>Period 2</option>
                <option value="3" <%if rs("period") = 3 then Response.write "selected" end if %>>Period 3</option>
                <option value="4" <%if rs("period") = 4 then Response.write "selected" end if %>>Period 4</option>
                <option value="5" <%if rs("period") = 5 then Response.write "selected" end if %>>Period 5</option>
                <option value="6" <%if rs("period") = 6 then Response.write "selected" end if %>>Period 6</option>
                <option value="7" <%if rs("period") = 7 then Response.write "selected" end if %>>Period 7</option>
                <option value="8" <%if rs("period") = 8 then Response.write "selected" end if %>>Period 8</option>
                <option value="9" <%if rs("period") = 9 then Response.write "selected" end if %>>Period 9</option>
                <option value="10" <%if rs("period") = 10 then Response.write "selected" end if %>>Period 10</option>
                <option value="11" <%if rs("period") = 11 then Response.write "selected" end if %>>Period 11</option>
                <option value="12" <%if rs("period") = 12 then Response.write "selected" end if %>>Period 12</option>
              </select>
              </font></td>
          </tr>
          <tr> 
            <td>
              <input type="submit" name="Submit" value="Save">
              <input type="button" name="Button" value="Delete" onclick="deleteentry(entryid.value,b.value,pid.value,date1.value)">
			  <%newurl="unreported.asp?b=" & b & "&date1=" & date1 & "&pid=" & pid &"&action=new"%>
              <input type="button" name="Button" value="New" onclick="javascript:parent.document.location='<%=newurl%>'">
              <input type="button" name="Submit2" value="Close" onclick="javascript:window.close()">
            </td>
          </tr>
        </table>
		
	  
	  <%
	  rs.close
	  end if
	  else %>
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td> <font face="Arial, Helvetica, sans-serif"> 
			  <input type="hidden" name="b" value="<%=b%>">
			  <input type="hidden" name="pid" value="<%=pid%>">
			  <input type="hidden" name="date1" value="<%=date1%>">
              <input type="radio" name="type" value="0" checked>
              Expense 
              <input type="radio" name="type" value="1">
              Revenue </font></td>
          </tr>
          <tr> 
            <td><font face="Arial, Helvetica, sans-serif">Description (max 150 
              characters)</font></td>
          </tr>
          <tr> 
            <td> <font face="Arial, Helvetica, sans-serif"> 
              <textarea name="description" cols="50" rows="5" wrap="PHYSICAL"></textarea>
              </font></td>
          </tr>
          <tr>
            <td><font face="Arial, Helvetica, sans-serif">Amount $ 
              <input type="text" name="amt" size="15" maxlength="15">
              Period 
              <select name="period">
                <option value="1">Period 1</option>
                <option value="2">Period 2</option>
                <option value="3">Period 3</option>
                <option value="4">Period 4</option>
                <option value="5">Period 5</option>
                <option value="6">Period 6</option>
                <option value="7">Period 7</option>
                <option value="8">Period 8</option>
                <option value="9">Period 9</option>
                <option value="10">Period 10</option>
                <option value="11">Period 11</option>
                <option value="12">Period 12</option>
              </select>
              </font></td>
          </tr>
          <tr> 
            <td>
              <input type="submit" name="Submit" value="Save">
              <input type="reset" name="Reset" value="Clear">
              <input type="button" name="Submit22" value="Close" onClick="javascript:window.close()">
            </td>
          </tr>
        </table>
		<%end if%>
      </form>
    </td>
  </tr>
</table>
</body>
</html>
