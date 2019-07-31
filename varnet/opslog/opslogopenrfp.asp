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

sqlstr = "SELECT rfplog.[Entry ID], rfplog.[Entry Type],rfplog.[probability],rfplog.[amt],rfplog.[amt2], rfplog.[Recording Date], rfplog.[Floor/Room], rfplog.Description, rfplog.[Current Status],  Employees.[Last Name] + ' ' + Employees.[First Name] AS manager,  Customers.CompanyName AS customer FROM rfplog LEFT OUTER JOIN Customers ON rfplog.Customer = Customers.CustomerID LEFT OUTER JOIN Employees ON rfplog.salesManager = Employees.ID WHERE (rfplog.[current status] <> 'Closed') and (rfplog.[current status] <> 'Proposal Accepted')and (rfplog.[current status] <> 'Proposal Rejected')order by rfplog.[Entry ID] "

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
          Open RFP</font></b></font>
      
    </td>
  </tr>
  <tr>
    <td>
      <table width="100%" border="0" cellpadding="0" cellspacing="10">
        <tr> 
          <td width="5%" bgcolor="#CCCCCC"><b><font face="Arial, Helvetica, sans-serif" size="2">Project 
            Manager</font></b></td>
          <td width="3%" bgcolor="#CCCCCC"><b><font face="Arial, Helvetica, sans-serif" size="2">Job 
            #</font></b></td>
          <td width="7%" bgcolor="#CCCCCC"><b><font face="Arial, Helvetica, sans-serif" size="2">Recording 
            Date</font></b></td>
          <td width="5%" bgcolor="#CCCCCC"> 
            <div align="center"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
              <% if Session("corp") > 4 then %>
              Primary Amount 
              <%end if%>
              </font></b></div>
          </td>
          <td width="5%" bgcolor="#CCCCCC"> 
            <div align="center"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
              <% if Session("corp") > 4 then %>
              Secondary Amount 
              <%end if%>
              </font></b></div>
          </td>
          <td width="7%" bgcolor="#CCCCCC"> 
            <div align="center"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
              <% if Session("corp") > 4 then %>
              A/P S 
              <%end if%>
              </font></b></div>
          </td>
          <td width="9%" bgcolor="#CCCCCC"><b><font face="Arial, Helvetica, sans-serif" size="2">Type</font></b></td>
          <td width="19%" bgcolor="#CCCCCC"><b><font face="Arial, Helvetica, sans-serif" size="2">Customer</font></b></td>
          <td width="28%" bgcolor="#CCCCCC"><b><font face="Arial, Helvetica, sans-serif" size="2">Description</font></b></td>
        </tr>
        <%dim amtp(5)
		amtp(1) = 0
		amtp(2) = 0
		amtp(3) = 0
		amtp(4) = 0
		amtp(5) = 0
		While not rst1.EOF %>
        <tr align="left" valign="top"> 
          <td width="5%" height="37"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("manager")%></font></td>
          <td width="3%" height="37"><font face="Arial, Helvetica, sans-serif" size="2"><a href=<%="rfpview.asp?rfp=" & rst1("entry id")%>><%=rst1("entry id")%></a></font></td>
          <td width="7%" height="37"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("recording date")%></font></div>
          </td>
          <td width="5%" height="37"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
              </font></b> 
              <% if Session("corp") > 4 then %>
              $<%=rst1("amt")%> 
              <%end if%>
              </font></div>
          </td>
          <td width="5%" height="37"> 
            <div align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
              </font></b> 
              <% if Session("corp") > 4 then %>
              $<%=rst1("amt2")%> 
              <%end if%>
              </font></div>
          </td>
          <td width="7%" height="37"> 
            <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
              </font></b> 
              <% if Session("corp") > 4 then %>
              <%=rst1("probability")%> 
              <%else %>
              <%=rst1("Current Status")%> 
              <%end if%>
              </font></div>
          </td>
          <td width="9%" height="37"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("entry type")%></font></td>
          <td width="19%" height="37"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("customer")%></font></td>
          <td width="28%" height="37"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("description")%></font></td>
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
<table align="center" border="0" cellspacing="10" cellpadding="0" style="font-family: Arial, Helvetica, sans-serif;font-size:13">
<tr style="font-weight:bold"><td></td>
	<td>Proposal Amt</td>
	<td>Probable Work</td>
	<td>Weighted Factor</td></tr>
<%
	rst1.close
	rst1.open "select probability,sum(amt) as amtsum, weight, sum(amt*weight) as probAmt from rfplog join percentage on prob=probability WHERE (rfplog.[current status] <> 'Closed')and (rfplog.[current status] <> 'Proposal Accepted')and (rfplog.[current status] <> 'Proposal Rejected') and [entry type] like '%RFP%' group by probability, weight order by probability", cnn1
	
	dim propAmtTotal, probAmtTotal
	propAmtTotal = 0
	probAmtTotal = 0
	do until rst1.eof
		response.write "<tr><td>"&rst1("probability")&"</td>"
		response.write "<td align=""right"">"&formatcurrency(rst1("amtsum"),2)&"</td>"
		response.write "<td align=""right"">"&formatcurrency(rst1("probAmt"),2)&"</td>"
		response.write "<td align=""right"">"&formatpercent(trim(rst1("weight")))&"</td></tr>"
		propAmtTotal = propAmtTotal + cDBL(rst1("amtsum"))
		probAmtTotal = probAmtTotal + cDBL(rst1("probAmt"))
		rst1.movenext
	loop
	rst1.close
	%>
	<tr><td>Total</td><td align="right"><b><%=formatcurrency(propAmtTotal)%></b></td><td align="right"><b><%=formatcurrency(probAmtTotal)%></b></td><td></td></tr>
	</table>
	<%
end if
%>
</body>
</html>