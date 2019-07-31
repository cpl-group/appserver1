<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<%
job= Request.Querystring("job")
user=Session("login")
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_main")

'sqlstr = "select * from " & "[job log]" & " where " & "[entry id]" & "=" & job 

sqlstr = "select Distinct customers.companyname, [employees].[First Name] + ' ' + [employees].[Last Name] AS projmanager, [job log].* from employees join [job log] on (employees.id=[job log].manager) join customers on ([job log].customer=customers.customerid )"

'response.write sqlstr
rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.eof then
%>

<table width="100%" border="0">
  <tr>
    <td>
      <div align="center">
        <p><font face="Arial, Helvetica, sans-serif"><i>Job <%=job%> not found 
          - please resubmit query or contact your system administrator </i></font></p>
        <p><font face="Arial, Helvetica, sans-serif"><i>
          <input type="button" name="Button" value="BACK" onclick="Javascript:history.back()">
          </i></font></p>
      </div>
    </td>
  </tr>
</table>
<%
else
%>
<form name="form1" method="post" action="savejob.asp">
<table width="100%" border="0">
  <tr> 
      <td bgcolor="#3399CC" height="30"> 
        <table width="100%" border="0" height="33">
          <tr> 
            <td width="73%" height="29"><b><i><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">New 
              Job</font></i></b><font face="Arial, Helvetica, sans-serif"><b><i><font color="#FFFFFF"> 
              <%=job%> 
              <input type="hidden" name="job" value="<%=job%>">
              </font></i></b></font></td>
            <td height="29" width="27%"> 
              <div align="right"><font face="Arial, Helvetica, sans-serif"><i> 
                <input type="button" name="Button2" value="BACK" onClick="Javascript:history.back()">
                </i></font></div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td height="2"> 
      <div align="left"> 
        <table width="100%" border="0">
          <tr bgcolor="#CCCCCC"> 
              <td width="35%"><font face="Arial, Helvetica, sans-serif">Customer:</font></td>
			  <td width="24%"><font face="Arial, Helvetica, sans-serif">Billing 
                Type:</font></td>
              <td width="41%"><font face="Arial, Helvetica, sans-serif">Secondary 
                Billing Type:</font></td>
          </tr>
          <tr> 
              <td width="35%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="customer">
                  <%Set rst4 = Server.CreateObject("ADODB.recordset")
			sqlstr = "select distinct customerid, companyname from customers order by companyname"
   			rst4.Open sqlstr, cnn1, 0, 1, 1
			if not rst4.eof then
				do until rst4.eof	
		%>
                  <option value="<%=rst4("customerid")%>"><font face="Arial, Helvetica, sans-serif"><%=rst4("companyname")%></font></option>
                  <%
					rst4.movenext
					loop
					end if
					rst4.close
				%>
                </select>
                <input type="hidden" name="cid" value="<%=rst1("customer")%>">
                </font></td>
			 
			  <td width="24%"><font face="Arial, Helvetica, sans-serif"> 
                <select name="cost">
                  <%Set rst3 = Server.CreateObject("ADODB.recordset")
			  str3="select * from jobtypes order by id"
			  rst3.Open str3, cnn1, 0, 1, 1
			  do until rst3.eof 
			  %>
                  <option value="<%=rst3("jobtype")%>"><%=rst3("jobtype")%></option>
                  <%
			      
			  rst3.movenext
			  loop
			  rst3.close
			  %>
                </select>
                $ 
                <input type="text" name="amt" size="7" maxlength="7" value="0" >
                </font> </td>
			   
              <td width="41%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="seccost">
                  <%Set rst3 = Server.CreateObject("ADODB.recordset")
			  str3="select * from jobtypes order by id"
			  rst3.Open str3, cnn1, 0, 1, 1
			  do until rst3.eof 
			  %>
                  <option value="<%=rst3("jobtype")%>"><font face="Arial, Helvetica, sans-serif"><%=rst3("jobtype")%></font></option>
                  <%
			      
			  rst3.movenext
			  loop
			  rst3.close
			  %>
                </select>
                $ 
                <input type="text" name="secamt" size="7" maxlength="7" value="0" >
                </font></td>
          </tr>
		  </table>
		  
		 <table width="100%" border="0">
          <tr bgcolor="#CCCCCC">
		      <td width="31%"><font face="Arial, Helvetica, sans-serif">Type:</font></td> 
              <td width="31%"><font face="Arial, Helvetica, sans-serif">Contact 
                Name:</font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">Customer 
                Phone # </font></td>
              <td width="39%"><font face="Arial, Helvetica, sans-serif">Fax Number</font></td>
			  
          </tr>
          <tr> 
              <td width="31%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="type1">
                  <%Set rst2 = Server.CreateObject("ADODB.recordset")
			sqlstr = "SELECT [Type ID]FROM [Genergy Entry Types]where job=1 ORDER BY [Type ID] "
   			rst2.Open sqlstr, cnn1, 0, 1, 1
			if not rst2.eof then
				do until rst2.eof	
		%>
                  <option value="<%=rst2("Type ID")%>"><font face="Arial, Helvetica, sans-serif"><%=rst2("Type ID")%></font></option>
                  <%
					rst2.movenext
					loop
				
					end if
					rst2.close
					%>
                </select>
                <input type="hidden" name="entrytype" value="<%=rst1("entry type")%>">
                </font></td>
              <td width="31%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="contactname" size="40" maxlength="40">
              </font></td>
              <td width="30%"> <font face="Arial, Helvetica, sans-serif">
                <input type="text" name="customerphone" >
                </font></td>
              <td width="39%"> <font face="Arial, Helvetica, sans-serif">
                <input type="text" name="customerfax" >
                </font></td>
			
          </tr>
          <tr bgcolor="#CCCCCC"> 
              <td width="31%"><font face="Arial, Helvetica, sans-serif">Customer Email</font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">Requested 
                By Name:</font></td>
              <td width="39%"><font face="Arial, Helvetica, sans-serif">Requested 
                By Phone Number:</font></td>
			  <td width="39%"><font face="Arial, Helvetica, sans-serif">Referred 
                By</font></td>
          </tr>
          <tr> 
             <td width="31%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="email" size="40" maxlength="40">
              </font></td>
              <td width="30%"> <font face="Arial, Helvetica, sans-serif">
                <input type="text" name="reqname" size="40" maxlength="40">
                </font></td>
              <td width="39%"> <font face="Arial, Helvetica, sans-serif">
                <input type="text" name="reqphone" >
                </font></td>
			  <td width="39%"> <font face="Arial, Helvetica, sans-serif">
                <input type="text" name="refby" >
                </font></td>
          </tr>
       </table>
	    <table width="100%" border="0">
            <tr bgcolor="#CCCCCC"> 
              <td width="76%"><font face="Arial, Helvetica, sans-serif">Floor 
                / Room</font></td>
            </tr>
            <tr> 
              <td width="76%"><font face="Arial, Helvetica, sans-serif"> 
                <select name="floorroom">
                  <%
			  Set rst7 = Server.CreateObject("ADODB.recordset")
			  str7 = "select * from floors "
   			  rst7.Open str7, cnn1, 0, 1, 1
			  if not rst7.eof then
			      do until rst7.eof	
		       	  if rst7("floor")=rst1("floor/room") then
			  %>
                  <option value="<%=rst1("floor/room")%>" selected ><%=rst1("floor/room")%></option>
                  <%
			      else
			  %>
                  <option value="<%=rst7("floor")%>"><%=rst7("floor")%></option>
                  <%
			  	  end if
				  rst7.movenext
				  loop
			  end if
			  %>
                </select>
                </font></td>
            </tr>
          </table>
        <table width="100%" border="0">
          <tr> 
            <td><font face="Arial, Helvetica, sans-serif">Description / Comments</font></td>
          </tr>
          <tr> 
              <td valign="top"> <font face="Arial, Helvetica, sans-serif"> 
                <textarea name="description" rows="5" cols="75" ></textarea>
                </font></td>
          </tr>
        </table>
        <table width="100%" border="0">
          <tr bgcolor="#CCCCCC"> 
              <td width="11%"><font face="Arial, Helvetica, sans-serif">Entered 
                By</font></td>
              <td width="15%"><font face="Arial, Helvetica, sans-serif">Project 
                Manager</font></td>
              <td width="19%"><font face="Arial, Helvetica, sans-serif">Recording 
                Date (mm/dd/yyyy)</font></td>
				 
              <td width="18%"><font face="Arial, Helvetica, sans-serif">Start 
                Date (mm/dd/yyyy) </font></td>
			  <td width="37%"><font face="Arial, Helvetica, sans-serif">End Date 
                (mm/dd/yyyy) </font></td>
          </tr>
          <tr> 
              <td width="11%"><font face="Arial, Helvetica, sans-serif"> 
                <input type="hidden" name="EnteredBy" value="<%=user%>">
              <%=user%>
              </font></td>
			  
              <td width="15%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="manager">
                  <%
				Set rst8 = Server.CreateObject("ADODB.recordset")

				sqlstr = "select * from Managers order by lastname, firstname"
				rst8.Open sqlstr, cnn1, 0, 1, 1
				do until rst8.eof%>
					<option value="<%=rst8("mid")%>"><%=rst8("lastname")%>, <%=rst8("firstname")%></option><%
					rst8.movenext
				loop
				rst8.close
				%>
                </select>
                <input type="hidden" name="mid" value="<%=rst1("manager")%>">
                </font></td>
              <td width="19%"><font face="Arial, Helvetica, sans-serif"> 
                <input type="hidden" name="recdate" value="<%=date()%>">
				<%=date()%>
                </font></td>
				
              <td width="18%"><font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="stdate" value="<%=date()%>">
                </font></td>
				
              <td width="37%"><font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="enddate" value="<%=date()+7%>">
                </font></td>
          </tr>
          <tr bgcolor="#CCCCCC"> 
              <td width="11%"><font face="Arial, Helvetica, sans-serif">Current 
                Status</font></td>
              <td width="15%"><font face="Arial, Helvetica, sans-serif">% completed 
                </font></td>
              <td width="19%"><font face="Arial, Helvetica, sans-serif">Last Bill 
                Date (mm/dd/yyyy)</font></td>
				
              <td width="18%"></td>
				
              <td width="37%"></td>
          </tr>
          <tr> 
           
			  <td width="11%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="status">
                  <%
				Set rst7 = Server.CreateObject("ADODB.recordset")

				sqlstr = "select status from status where job=1"
				rst7.Open sqlstr, cnn1, 0, 1, 1
				if not rst7.eof then
				do until rst7.eof	
		%>
          <option value="<%=rst7("status")%>"><font face="Arial, Helvetica, sans-serif"><%=rst7("status")%></font></option>
          <%
					rst7.movenext
					loop
					end if
					rst7.close
				%>
        </select>
              </font></td>
                
				 <td width="11%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="percentcomp">
         <option value="25"><font face="Arial, Helvetica, sans-serif">25</font></option>
		  <option value="50"><font face="Arial, Helvetica, sans-serif">50</font></option>
		  <option value="75"><font face="Arial, Helvetica, sans-serif">75</font></option>
		  <option value="100"><font face="Arial, Helvetica, sans-serif">100</font></option>
        </select>
              </font></td>
              
              <td width="19%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="billdate" >
              </font></td>
			  <td width="18%"></td>
          </tr>
        </table>
        <table width="100%" border="0">
          <tr> 
            <td bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Comments</font></td>
          </tr>
          <tr> 
            <td> <font face="Arial, Helvetica, sans-serif"> 
              <textarea name="comments" rows="5" cols="75"></textarea>
              </font></td>
          </tr>
        </table>
          <font face="Arial, Helvetica, sans-serif"><i> 
          <input type="submit" name="saveit" value="SAVE" >
		  
          <input type="button" name="cancel" value="CANCEL" onclick='javascript:parent.document.location="oplogindex.asp"'>
          </i></font></div>
    </td>
  </tr>
</table>

</form>
<%
end if
%>
</body>
</html>
