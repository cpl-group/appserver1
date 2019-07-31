<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
job = Request.Querystring("job")
mkid = Request.Querystring("mkid")
cnum = Request.Querystring("cnum")
user = Session("login")
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rstC = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")

if trim(mkid)<>"" and trim(cnum)<>"" then
    rst1.open "SELECT * from customers where id="& cnum, cnn1
    if not(rst1.eof) then
        CustomerID = rst1("CustomerID")
        custname = rst1("ContactFirstName") &" "& rst1("ContactLastName")
        phone = rst1("PhoneNumber")
        fax = rst1("FaxNumber")
    end if
    rst1.close
end if 


'sqlstr = "select * from " & "[job log]" & " where " & "[entry id]" & "=" & job 

sqlstr = "select Distinct customers.companyname, [employees].[First Name] + ' ' + [employees].[Last Name] AS projmanager, [job log].* from employees join [job log] on (employees.id=[job log].manager) join customers on ([job log].customer=customers.customerid )"

'response.write sqlstr
rst1.Open sqlstr, cnn1, 0, 1, 1

dim contactname, PhoneNumber, FaxNumber, custnum
contactname = ""
PhoneNumber = ""
FaxNumber = ""
custnum = request("customer")
if trim(custnum)="" then custnum=request("cnum")
custnum = trim(custnum)
if trim(custnum)<>"" then
	rstC.open "select * from customers where customerid="&custnum&" order by companyname", cnn1
	if not rstC.eof then
		contactname = rstC("ContactFirstName")&" "&rstC("ContactLastName")
		PhoneNumber = rstC("PhoneNumber")
		FaxNumber = rstC("FaxNumber")
	end if
end if


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
<form name="form1" method="post" action="saverfp.asp">
<table width="100%" border="0">
  <tr> 
      <td bgcolor="#3399CC" height="30"> 
        <table width="100%" border="0" height="33">
          <tr> 
            <td width="73%" height="29"><b><i><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">New 
              RFP</font></i></b><font face="Arial, Helvetica, sans-serif"><b><i><font color="#FFFFFF"> 
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
              <td width="31%"><font face="Arial, Helvetica, sans-serif">Customer:</font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">Proposal 
                Type:</font></td>
              <td width="39%"><font face="Arial, Helvetica, sans-serif">Type:</font></td>
            </tr>
            <tr> 
              <td width="31%"> <font face="Arial, Helvetica, sans-serif" size="2"> 
                <select name="customer" onChange="document.location.href='newrfp.asp?customer='+this.value">
                  <%Set rst4 = Server.CreateObject("ADODB.recordset")
			sqlstr = "select distinct customerid, companyname from customers order by companyname"
   			rst4.Open sqlstr, cnn1, 0, 1, 1
			if not rst4.eof then
				do until rst4.eof	
		%>
                  <option value="<%=rst4("customerid")%>" <%if trim(rst4("customerid"))=(custnum) then response.write "SELECTED"%>><font face="Arial, Helvetica, sans-serif" size="2"><%=rst4("companyname")%></font></option>
                  <%
					rst4.movenext
					loop
					end if
					rst4.close
				%>
                </select>
                <input type="hidden" name="cid" value="<%=rst1("customer")%>">
                </font></td>
              <td width="30%"> <font face="Arial, Helvetica, sans-serif" size="2"> 
                1st : 
                <select name="cost">
                  <%Set rst3 = Server.CreateObject("ADODB.recordset")
			  str3="select * from rfptype order by id"
			  rst3.Open str3, cnn1, 0, 1, 1
			  do until rst3.eof 
			  %>
                  <option value="<%=rst3("rfptype")%>"><%=rst3("rfptype")%></option>
                  <%
			      
			  rst3.movenext
			  loop
			  rst3.close
			  %>
                </select>
                $ 
                <input type="text" name="amt" size="5" maxlength="10" value="0">
                </font></td>
              <td width="39%"> <font face="Arial, Helvetica, sans-serif" size="2"> 
                <select name="type1">
                  <%Set rst2 = Server.CreateObject("ADODB.recordset")
			sqlstr = "SELECT [Type ID]FROM [Genergy Entry Types]where [type id] like '%RFP%' ORDER BY [Type ID]"
   			rst2.Open sqlstr, cnn1, 0, 1, 1
			if not rst2.eof then
				do until rst2.eof	
		%>
                  <option value="<%=rst2("Type ID")%>"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst2("Type ID")%></font></option>
                  <%
					rst2.movenext
					loop
				
					end if
					rst2.close
					%>
                </select>
                <input type="hidden" name="entrytype" value="<%=rst1("entry type")%>">
                </font></td>
            </tr>
            <tr> 
              <td width="31%">&nbsp;</td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif" size="2">2nd 
                : 
                <select name="cost2">
                  <%Set rst3 = Server.CreateObject("ADODB.recordset")
			  str3="select * from rfptype order by id"
			  rst3.Open str3, cnn1, 0, 1, 1
			  do until rst3.eof 
			  %>
                  <option value="<%=rst3("rfptype")%>"><%=rst3("rfptype")%></option>
                  <%
			      
			  rst3.movenext
			  loop
			  rst3.close
			  %>
                </select>
                $ 
                <input type="text" name="amt2" size="5" maxlength="10" value="0" >
                </font></td>
              <td width="39%">&nbsp;</td>
            </tr>
          </table>
		  
		  <table width="100%" border="0">
            <tr> 
              <td width="31%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Contact Name:</font></td>
              <td width="30%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Customer Phone #:</font></td>
              <td width="39%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Customer Fax Number:</font></td>
            </tr>
            <tr> 
              <td width="31%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="hidden" name="contactname" size="40" maxlength="40" value="<%=contactname%>">
                <%=contactname%></font></td>
              <td width="30%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="hidden" name="customerphone" value="<%=PhoneNumber%>">
                <%=PhoneNumber%></font></td>
              <td width="39%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="hidden" name="customerfax" value="<%=FaxNumber%>">
                <%=FaxNumber%></font></td>
            </tr>
            <tr> 
              <td width="31%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Referred By</font></td>
              <td width="30%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Requested By Name:</font></td>
              <td width="39%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Requested By Phone Number:</font></td>
            </tr>
            <tr> 
              <td width="31%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="refby" >
                </font></td>
              <td width="30%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="reqname" size="40" maxlength="40">
                </font></td>
              <td width="39%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="reqphone">
                </font></td>
            </tr>
          </table>
	      <table width="100%" border="0">
            <tr bgcolor="#CCCCCC"> 
              <td width="31%"><font face="Arial, Helvetica, sans-serif">Floor 
                / Room</font></td>
            </tr>
            <tr> 
              <td width="31%" height="41"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="floorroom">
                  <%Set rst5 = Server.CreateObject("ADODB.recordset")
			sqlstr = "select * from floors"
   			rst5.Open sqlstr, cnn1, 0, 1, 1
			if not rst5.eof then
				do until rst5.eof	
		%>
                  <option value="<%=rst5("floor")%>"><font face="Arial, Helvetica, sans-serif"><%=rst5("floor")%></font></option>
                  <%
					rst5.movenext
					loop
					end if
					rst5.close
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
              <td width="14%"><font face="Arial, Helvetica, sans-serif">Entered 
                By</font></td>
              <td width="23%"><font face="Arial, Helvetica, sans-serif">Sales 
                Manager </font></td>
              <td width="34%"><font face="Arial, Helvetica, sans-serif">Start 
                Date (mm/dd/yyyy) </font></td>
				 
              <td width="29%"><font face="Arial, Helvetica, sans-serif">Estimated 
                Completion Date(mm/dd/yyyy) </font></td>
          </tr>
          <tr> 
              <td width="14%"><font face="Arial, Helvetica, sans-serif"> 
                <input type="hidden" name="EnteredBy" value="<%=user%>">
              <%=user%>
              </font></td>
              <td width="23%"> <font face="Arial, Helvetica, sans-serif"> 
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
              <td width="34%"><font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="stdate" value="<%=date()%>">
                </font></td>
				
              <td width="29%"><font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="enddate" value="<%=date()%>">
                </font></td>
          </tr>
          <tr bgcolor="#CCCCCC"> 
              <td width="14%"><font face="Arial, Helvetica, sans-serif">Current 
                Status</font></td>
              <td width="23%"><font face="Arial, Helvetica, sans-serif">Acceptance 
                Probability</font></td>
             
				
              <td width="29%">&nbsp;</td>
				
              <td width="0%"></td>
          </tr>
          <tr> 
           
			  <td width="14%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="status">
                  <%
				Set rst7 = Server.CreateObject("ADODB.recordset")

				sqlstr = "select status from status where job=0"
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
              <td width="23%">
                <input type="text" name="prob" >
              </td>
              <td width="34%">&nbsp;</td>
			  <td width="29%"></td>
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
		 <input type="hidden" name="mkid" value="<%=mkid%>">
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
