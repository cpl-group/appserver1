<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">


<%@Language="VBScript"%>
<%
job= Request.Querystring("job")

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_main")

'sqlstr = "select * from " & "[job log]" & " where " & "[entry id]" & "=" & job 

sqlstr = "SELECT dbo.[Job Log].[Entry ID], dbo.[Job Log].[Entry Type], dbo.[Job Log].jobtype,dbo.[Job Log].amt,dbo.[Job Log].[recording date],dbo.[Job Log].[Recording Date], dbo.[Job Log].Building, dbo.[Job Log].[Floor/Room], dbo.[Job Log].Description, dbo.[Job Log].[Current Status], dbo.[Job Log].[% completed], dbo.Employees.[Last Name] + ' ' + dbo.Employees.[First Name] AS manager, dbo.Customers.CompanyName AS customer FROM dbo.[Job Log] INNER JOIN dbo.Employees ON dbo.[Job Log].Manager = dbo.Employees.ID INNER JOIN dbo.Customers ON dbo.[Job Log].Customer = dbo.Customers.CustomerID WHERE (dbo.[Job Log].[Current Status] NOT LIKE '%Closed%')and [entry type] NOT LIKE '%RFP%' order by [Job Log].[Entry ID]"

'response.write sqlstr
rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.EOF then
%>
<table width="100%" border="0">
  <tr>
    <td>ERROR</td>
  </tr>
</table>
<%
else

%>
<table width="100%" border="0">
  <tr>
    <td bgcolor="#3399CC" height="30"> <font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF" size="4">Genergy 
      Work In Progress</font></b></font> </td>
  </tr>
  <tr>
    <td> 
      <table width="100%" border="0" cellpadding="0" cellspacing="10">
        <tr> 
          <td  bgcolor="#CCCCCC" height="28" width="6%"><font face="Arial, Helvetica, sans-serif">Project 
            Manager </font></td>
          <td  bgcolor="#CCCCCC" height="28" width="4%"><font face="Arial, Helvetica, sans-serif">Job 
            # </font></td>
          <td bgcolor="#CCCCCC" height="28" width="7%"><font face="Arial, Helvetica, sans-serif">Recording 
            Date</font></td>
          <td  bgcolor="#CCCCCC" height="28" width="4%"><font face="Arial, Helvetica, sans-serif">Status</font></td>
          <td  bgcolor="#CCCCCC" height="28" width="15%"><font face="Arial, Helvetica, sans-serif">Type</font></td>
          <td bgcolor="#CCCCCC" height="28" width="6%"><font face="Arial, Helvetica, sans-serif"> 
            Billing Type</font></td>
			
          <td bgcolor="#CCCCCC" height="28" width="5%"><font face="Arial, Helvetica, sans-serif"> 
            Contract Amt</font></td>
          <td  bgcolor="#CCCCCC" height="28" width="5%"><font face="Arial, Helvetica, sans-serif">Floor</font></td>
          <td  bgcolor="#CCCCCC" height="28" width="14%"><font face="Arial, Helvetica, sans-serif">Customer</font></td>
          <td  bgcolor="#CCCCCC" height="28" width="22%"><font face="Arial, Helvetica, sans-serif">Description</font></td>
        </tr>
        <% While not rst1.EOF %>
        <tr align="left" valign="top"> 
          <td  height="37" width="6%"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("manager")%></font></td>
          <td  height="37" width="4%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><a href=<%="opslogview.asp?job=" & rst1("entry id")%>><%=rst1("entry id")%></a></font></div>
          </td>
          <td " height="37" width="7%"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("Recording date")%></font></td>
          <td  height="37" width="4%"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("Current Status")%></font></td>
          <td  height="37" width="15%"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("entry type")%></font></td>
          <td  height="37" width="6%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("jobtype")%></font></div>
          </td>
		   <td  height="37" width="5%"> 
		  <%if Session("opslog") =5 then  %>
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("amt")%></font></div>
			<%else%>
			 <div align="center"><font face="Arial, Helvetica, sans-serif" size="2">N/A</font></div>
			 <% end if%>
          </td>
          <td  height="37" width="5%"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("floor/room")%></font></div>
          </td>
          <td  height="37" width="14%"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("customer")%></font></td>
          <td width="22%" height="37"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("description")%></font></td>
        </tr>
        <% 
		rst1.movenext
		Wend
		%>
      </table>
    </td>
  </tr>
</table>
<p>&nbsp;</p>
<%
end if
%>
</body>
</html>