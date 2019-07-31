<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
input= Request.Querystring("findvar")
item= Request.Querystring("select")
completed = Request.Querystring("comp")
var= Request.Querystring("var")
	if isempty(input) then
				msg="Please enter search and click the FIND button to begin"
				 'Write a browser-side script to update another frame (named
				 'detail) within the same frameset that displays this page.
				Response.Write "<script>" & vbCrLf
			    Response.Write "parent.location = " & _
                Chr(34) & "rfpindex.asp?msg=" & msg & Chr(34) & vbCrLf
				Response.Write "</script>" & vbCrLf
	end if

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")

if completed="1" then
	if item = "salesmanager" then 
	
	sqlstr = "select * from " & "rfplog" & " where " & item & " = " & input 	
	else
	
	sqlstr = "select * from " & "rfplog" & " where " & item & " like '%" & input & "%'"
	
	end if
	if item="customer" and not isempty(var) then
		sqlstr="select distinct rfplog. customer as customer1, customers.companyname, [entry id], ChgOrderRefNum, [% completed], salesmanager, description from rfplog join customers on rfplog.customer=customers.customerid and companyname like '%" & var & "%'"
	end if	
else 

	if item = "salesmanager" then 
	
	sqlstr = "select * from " & "rfplog" & " where " & item & " =" & input & " and [current status] <> 'Proposal Accepted' and [current status] <> 'Proposal Rejected'  and [current status] <> 'Closed'" 	
	else
	
	sqlstr = "select * from " & "rfplog" & " where " & item & " like '%" & input & "%' and [current status] <> 'Proposal Accepted' and [current status] <> 'Proposal Rejected' and [current status] <> 'Closed'" 
	
	end if
		
    if item="customer" and not isempty(var) then
		sqlstr="select distinct rfplog. customer as customer1, customers.companyname, [entry id], ChgOrderRefNum, [% completed], salesmanager, description from rfplog join customers on rfplog.customer=customers.customerid and companyname like '%" & var & "%' and [current status] <> 'Proposal Accepted' and [current status] <> 'Proposal Rejected' and [current status] <> 'Closed'"
	end if

end if

rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic
'response.write sqlstr
'response.end
rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.EOF then 
	msg="Last search not found...please try again"
	Response.Write "<script>" & vbCrLf
	Response.Write "parent.location = " & _
    Chr(34) & "rfpindex.asp?msg=" & msg & Chr(34) & vbCrLf
	Response.Write "</script>" & vbCrLf
Else
x=0
%>

<script>
function updcust(custid){
	var  temp="updcustomer.asp?custid="+custid
	window.open(temp,"", "scrollbars=yes,width=600, height=300, status=no, menubar=no" );
}

</script>
<body bgcolor="#FFFFFF">
<form name="form1" method="post" action=""><table width="100%" border="0">
  <tr>
    <td bgcolor="#3399CC" height="36" width="13%"> 
	 
	    <div align="center"> <font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF" size="4"> 
          RFP LOG SEARCH RESULTS </font></b></font></div>
	</td>  
  </tr>
  <% if item="customer" then%>
 <tr>
      <td height="36" width="13%"> 
        <input type="button" name="Button" value="Update Customer" onclick="updcust('<%=rst1("customer")%>')">
      </td>
 </tr>
<%end if%>
  <tr>
   
    <td width="87%"> 
      <table width="100%" border="0">
        <tr bgcolor="#CCCCCC"> 
		
            <td bgcolor="#CCCCCC" width="9%"><font face="Arial, Helvetica, sans-serif" color="#000000">RFP 
              #</font></td>
          <td width="18%"><font face="Arial, Helvetica, sans-serif" color="#000000">Description</font></td>
          <td width="30%"><font face="Arial, Helvetica, sans-serif" color="#000000">Project 
            Manager</font></td>
          <td width="24%"><font face="Arial, Helvetica, sans-serif" color="#000000">% completed</font></td>
        </tr>
        <% While not rst1.EOF %>
        <tr> 
          <td width="9%"><font face="Arial, Helvetica, sans-serif"><a href=<%="rfpview.asp?rfp=" & rst1("entry id")%> ><%=rst1("entry id")%></a></font></td>
          <td width="18%"><font face="Arial, Helvetica, sans-serif"><%=rst1("description")%></font></td>
		  <%
		  manager=rst1("salesmanager")
		  strsql2="select [first name], [last name] from employees where id='" & manager &"'"
		  rst2.Open strsql2, cnn1, 0, 1, 1
		  if not rst2.eof then
		  %>
          <td width="30%"><font face="Arial, Helvetica, sans-serif"><%=rst2("first name")%>  
		  &nbsp<%=rst2("last name")%></font></td>
		  <%
		  end if
		  rst2.close
		  %>
          <td width="24%"><font face="Arial, Helvetica, sans-serif"><%=rst1("% completed")%></font></td>
        </tr>
        <%
		x=x+1
		rst1.movenext
		Wend
		%>
      </table>
    </td>
  </tr>
  
  <tr>
  
    <td bgcolor="#3399CC" width="13%"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF"><%=x%> Jobs 
        Found </font></b></font></div>
    </td>
  </tr>
 </table></form>
<%
end if
rst1.close
%>
</body>
