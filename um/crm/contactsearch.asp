<%@Language="VBScript"%>
<%
input= Request.Querystring("findvar")
item= Request.Querystring("select")
var= Request.Querystring("var")
	if isempty(input) then
				msg="Please enter search and click the FIND button to begin"
				 'Write a browser-side script to update another frame (named
				 'detail) within the same frameset that displays this page.
				Response.Write "<script>" & vbCrLf
			    Response.Write "parent.location = " & _
                Chr(34) & "contactindex.asp?msg=" & msg & Chr(34) & vbCrLf
				Response.Write "</script>" & vbCrLf
	end if

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"Intranet")


	if item = "rebny" then 
	
	sqlstr = "select * from contacts where rebny=1 and assc_member=0 and princ_member=0"
	else
	
    if item="boma" and not isempty(var) then
		if var="principal" then
		sqlstr="select * from contacts where  princ_members=1 and assc_members=0 and rebny=0"
			else
		sqlstr="select * from contacts where assc_memebers=1 and princ_members=0 and rebny=0"
	else
		sqlstr="select * from contacts where assc_members=1 and princ_members=1 and rebny=0"
	end if
	end if

end if

rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.EOF then 
	msg="Last search not found...please try again"
	Response.Write "<script>" & vbCrLf
	Response.Write "parent.location = " & _
    Chr(34) & "contactindex.asp?msg=" & msg & Chr(34) & vbCrLf
	Response.Write "</script>" & vbCrLf
Else
x=0
%>
<body bgcolor="#FFFFFF">
<form name="form1" method="post" action="">
<table width="100%" border="0">
  <tr>
    <td bgcolor="#3399CC" height="36" width="13%"> 
	 
	    <div align="center"> <font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF" size="4"> 
          CONTACT SEARCH RESULTS </font></b></font></div>
	</td>  
  </tr>
 
  <tr>
   
    <td width="87%"> 
      <table width="100%" border="0">
        <tr bgcolor="#CCCCCC"> 
		
            <td bgcolor="#CCCCCC" width="12%"><font face="Arial, Helvetica, sans-serif" color="#000000">Contact 
              Name </font></td>
            <td width="16%"><font face="Arial, Helvetica, sans-serif" color="#000000">Contact 
              Title </font></td>
            <td width="18%"><font face="Arial, Helvetica, sans-serif" color="#000000">Company 
              Name </font></td>
        </tr>
        <% While not rst1.EOF %>
        <tr> 
            <td width="12%"><font face="Arial, Helvetica, sans-serif"><a href=<%="contactview.asp?job=" & rst1("entry id")%> ><%=rst1("first_name")%><%=rst1("last_name")%></a></font></td>
            <td width="16%"><font face="Arial, Helvetica, sans-serif"><%=rst1("title")%></font></td>
			<td width="16%"><font face="Arial, Helvetica, sans-serif"><%=rst1("company")%></font></td>
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
  
      <td bgcolor="#3399CC" width="13%" height="19"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF"><%=x%> 
          Contacts Found </font></b></font></div>
    </td>
  </tr>
 </table></form>
<%
end if
rst1.close
%>
</body>
