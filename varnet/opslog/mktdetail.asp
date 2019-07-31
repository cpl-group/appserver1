<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<%
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")

cnn1.Open "driver={SQL Server};server=10.0.7.110;uid=genergy1;pwd=g1appg1;database=main;"

ID1= Request.Querystring("id")
mktid=Request.Querystring("mkid")
'response.write Request.Querystring("mkid")

if isempty(id1) then
%>
<form name="form2" method="post" action="savemktitem.asp">

<table width="100%" border="0">
  <tr bgcolor="#CCCCCC"> 
    <td width="2%" height="2"> 
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Date</font></td>
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Action</font></td>
      <td width="18%" height="10"><font face="Arial, Helvetica, sans-serif">Comment</font></td>
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Follow-Up Date </font></td>
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Follow-Up Action</font></td>
      <td width="18%" height="10"><font face="Arial, Helvetica, sans-serif">Follow-up Comment</font></td>
  </tr>
  <tr valign="top"> 
    <td width=6%> 
      <input type="submit" name="choice2"  value="SAVE">
      <input type="hidden" name="mktid" value="<%=mktid%>">
    </td>
    <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
      <input type="text" name="date" value="<%=date()%>" size="10">
      </font></td>
    <td width="14%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="action">
                  <%Set rst2 = Server.CreateObject("ADODB.recordset")
			sqlstr = "Select action from mkt_actions order by id"
   			rst2.Open sqlstr, cnn1, 0, 1, 1
			if not rst2.eof then
					do until rst2.eof
		%>
                  <option value="<%=rst2("action") %>"><i><b><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><%=rst2("action") %></font></b></i></option>
                  <%
				 
					rst2.movenext
					loop
					end if
					rst2.close
				%>
                </select>
        <font face="Arial, Helvetica, sans-serif"> </font></font></td>
      <td width="18%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <textarea name="comment" cols="25" rows="3" wrap="PHYSICAL"></textarea>
        </font></td>

      
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif">
        <input type="text" name="fdate" size="20" value="<%=dateadd("d", 7,date())%>">
        </font></td>
      <td width="11%" height="19" > <font face="Arial, Helvetica, sans-serif"> 
        <select name="faction">
          <%Set rst3 = Server.CreateObject("ADODB.recordset")
			sqlstr = "Select action from mkt_actions order by id"
   			rst3.Open sqlstr, cnn1, 0, 1, 1
			if not rst3.eof then
					do until rst3.eof
					%>
          <option value="<%=rst3("action") %>"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><%=rst3("action")%> 
          </font></option>
          <%
					rst3.movenext
					loop
					end if
					rst3.close
				%>
        </select>
        </font></td>
      <td width="14%"> <font face="Arial, Helvetica, sans-serif"> <font face="Arial, Helvetica, sans-serif"> 
        <textarea name="fcomment" cols="25" rows="3" wrap="PHYSICAL"></textarea>
        </font></font></td>
  </tr>
</table></form>
<%
else
Set rst1 = Server.CreateObject("ADODB.recordset")
sqlstr = "select * from mkt_progressitems where id="&id1
'response.write sqlstr
'response.end
rst1.Open sqlstr, cnn1, 0, 1, 1


%>
<form name="form1" method="post" action="mktitemupdate.asp">

<table width="100%" border="0">
  <tr bgcolor="#CCCCCC"> 
    <td width="2%" height="2"> 
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Date</font></td>
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Action</font></td>
      <td width="18%" height="10"><font face="Arial, Helvetica, sans-serif">Comment</font></td>
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Follow-Up Date </font></td>
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Follow-Up Action</font></td>
      <td width="18%" height="10"><font face="Arial, Helvetica, sans-serif">Follow-up Comment</font></td>
  </tr>
  
  <tr valign="top"> 
    
    <td width=6%> 
      <input type="submit" name="choice"  value="UPDATE">
	  <input type="hidden" name="key" value="<%=rst1("id")%>">
	<input type="hidden" name="mkid" value="<%=mktid%>">
    </td>
    <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
      <input type="text" name="date" value="<%=rst1("date")%>" size="15">
      </font></td>
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <select name="action">
          <%Set rst4 = Server.CreateObject("ADODB.recordset")
			sqlstr = "Select action from mkt_actions order by id"
   			rst4.Open sqlstr, cnn1, 0, 1, 1
			if not rst4.eof then
			do until rst4.eof
					If rst1("action")= rtrim(rst4("action")) then	
		%>
          <option value="<%=rst4("action")%>" selected><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><%=rst4("action")%></font></option>
          <%else
				  %>
          <option value="<%=rst4("action")%>"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><%=rst4("action")%></font></option>
          <%
				 	end if
					rst4.movenext
					loop
					end if
					rst4.close
				%>
        </select>
        </font></td>
      <td width="18%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <textarea name="comment" cols="25" rows="3" wrap="PHYSICAL"><%=rst1("comments")%></textarea>
        </font></td>

      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif">
        <input type="text" name="fdate" value="<%=rst1("followupdate")%>" size="15">
        </font></td>
      <td width="11%" height="19" > <font face="Arial, Helvetica, sans-serif"> 
        <select name="faction">
          <%Set rst5 = Server.CreateObject("ADODB.recordset")
			sqlstr = "Select action from mkt_actions order by id"
   			rst5.Open sqlstr, cnn1, 0, 1, 1
			if not rst5.eof then
					do until rst5.eof
					If rst1("followup")= rtrim(rst5("action")) then	
		%>
          <option value="<%=rst5("action") %>"selected><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><%=rst5("action") %></font></option>
          <%else
				  %>
          <option value="<%=rst5("action") %>"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><%=rst5("action")%> 
          </font></option>
          <%
				  end if
					rst5.movenext
					loop
					end if
					rst5.close
				%>
        </select>
        </font></td>
      <td width="18%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <textarea name="fcomment" cols="25" rows="3" wrap="PHYSICAL"><%=rst1("fcomment")%></textarea>
        </font></td>
    
  </tr>
  
</table>
 </form>      
 <%
end if
%>        
</body>
</html>
