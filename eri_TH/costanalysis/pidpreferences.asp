<%
function checkit(name)
	if session(name)=1 then
		checkit = " CHECKED"
	else
		checkit = ""
	end if
end function
if request.form("Accept")="Accept" then
	session("Expenses")				=request.form("Expenses")
	session("Submeter")				=request.form("Submeter")
	session("ERI")					=request.form("ERI")
	session("Expense_Adjustments")	=request.form("Expense_Adjustments")
	session("Revenue_Adjustments")	=request.form("Revenue_Adjustments")
	session("Mac_Revenue")			=request.form("Mac_Revenue")
	session("PLP_Revenue")			=request.form("PLP_Revenue")
'	session("Net")					=request.form("Net")
end if
%>
<html>
<head>
<title></title>
<%if request.form("Accept")="Accept" then%>
<script>
parent.settabs(parent.storedPreferenceHottab)
parent.loadchart();
</script>
<%end if%>
</head>
<body bgcolor="#FFFFFF">
<table width="100%" border="0" bgcolor="#000000">
  <tr>
    <td><b><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF">Preferences</font></b></td>
  </tr>
</table>
<form action="preferences.asp" method="post">
  <table width="349">
    <tr> 
      <td width="27"><b><font face="Arial, Helvetica, sans-serif" size="2"></font></b></td>
      <td width="150"> 
        <div align="center"><b><font face="Arial, Helvetica, sans-serif" size="2">Expenses</font></b></div>
      </td>
      <td width="10"><b><font face="Arial, Helvetica, sans-serif" size="2"></font></b></td>
      <td width="156"> 
        <div align="center"><b><font face="Arial, Helvetica, sans-serif" size="2">Revenue</font></b></div>
      </td>
    </tr>
    <tr> 
      <td width="27"> 
        <input name="Expenses" value="1" type="checkbox"<%=checkit("Expenses")%>>
      </td>
      <td width="150"><font face="Arial, Helvetica, sans-serif" size="2">Expenses</font></td>
      <td width="10"> 
        <input name="Submeter" value="1" type="checkbox"<%=checkit("Submeter")%>>
      </td>
      <td width="156"><font face="Arial, Helvetica, sans-serif" size="2">Submeter</font></td>
    </tr>
    <tr> 
      <td width="27"> 
        <input name="Expense_Adjustments" value="1" type="checkbox"<%=checkit("Expense_Adjustments")%>>
      </td>
      <td width="150"><font face="Arial, Helvetica, sans-serif" size="2">Expense 
        Adjustments</font></td>
      <td width="10"> 
        <input name="ERI" value="1" type="checkbox"<%=checkit("ERI")%>>
      </td>
      <td width="156"><font face="Arial, Helvetica, sans-serif" size="2">ERI</font></td>
    </tr>
    <tr> 
      <td width="27">&nbsp; </td>
      <td width="150"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
      <td width="10"> 
        <input name="Revenue_Adjustments" value="1" type="checkbox"<%=checkit("Revenue_Adjustments")%>>
      </td>
      <td width="156"><font face="Arial, Helvetica, sans-serif" size="2">Revenue 
        Adjustments</font></td>
    </tr>
    <tr> 
      <td width="27">&nbsp; </td>
      <td width="150"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
      <td width="10"> 
        <input name="Mac_Revenue" value="1" type="checkbox"<%=checkit("Mac_Revenue")%>>
      </td>
      <td width="156"><font face="Arial, Helvetica, sans-serif" size="2">Mac Revenue</font></td>
    </tr>
    <tr> 
      <td width="27">&nbsp; </td>
      <td width="150"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
      <td width="10"> 
        <input name="PLP_Revenue" value="1" type="checkbox"<%=checkit("PLP_Revenue")%>>
      </td>
      <td width="156"><font face="Arial, Helvetica, sans-serif" size="2">PLP Revenue</font></td>
    </tr>
    <tr> 
      <td width="27">&nbsp; </td>
      <td width="150"> <font face="Arial, Helvetica, sans-serif" size="2"> 
        <input name="Accept" type="submit" value="Accept">
        </font></td>
      <td width="10">&nbsp;</td>
      <td width="156"><font face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
    </tr>
    <%
'<tr><td><input name="Net" value="1" type="checkbox"<%=checkit("Net")></td>
'	<td>Net</td></tr>
%>
  </table>
</form>

</body>
</html>
