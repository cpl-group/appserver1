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

cnn1.Open application("cnnstr_main")


%>

<form name="form1" method="post" action="savepo.asp">
  <table width="100%" border="0">
    <tr> 
      <td bgcolor="#3399CC" height="2"> 
        <table width="100%" border="0">
          <tr> 
            <td height="2"><i><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Details 
              for New PO: <%=ponum%> 
              <input type="hidden" name="job" value="<%=job%>">
              </font></b></i></td>
            <td height="2"> 
              <div align="right"><font face="Arial, Helvetica, sans-serif"><i> 
                <input type="button" name="Button2" value="BACK" onClick="Javascript:history.back()">
                </i></font></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td height="2" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">PO 
        Description</font></td>
    </tr>
    <tr> 
      <td height="2">
        <textarea name="description" cols="50" rows="3"></textarea>
      </td>
    </tr>
    <tr> 
      <td height="2"> 
        <div align="left"> 
          <table width="100%" border="0">
            <tr bgcolor="#CCCCCC"> 
              <td width="8%" height="28"><font face="Arial, Helvetica, sans-serif">Job 
                Number</font></td>
              <td width="11%" height="28"><font face="Arial, Helvetica, sans-serif">Date 
                (mm/dd/yyyy)</font></td>
              <td width="6%" height="28"><font face="Arial, Helvetica, sans-serif">Vendor:</font></td>
              <td width="14%" height="28"><font face="Arial, Helvetica, sans-serif">Job 
                Address:</font></td>
              <td width="14%" height="28"><font face="Arial, Helvetica, sans-serif">Ship 
                Address:</font></td>
              <td width="36%" height="28"><font face="Arial, Helvetica, sans-serif">Requistioner</font></td>
            </tr>
            <tr> 
              <td width="8%" height="32"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="jobnum">
                  <%Set rst2 = Server.CreateObject("ADODB.recordset")
			sqlstr = "select distinct [entry id] from [job log] order by [entry id] desc"
   			rst2.Open sqlstr, cnn1, 0, 1, 1
			if not rst2.eof then
				do until rst2.eof	
		%>
                  <option value="<%=rst2("entry id")%>"><font face="Arial, Helvetica, sans-serif"><%=rst2("entry id")%></font></option>
                  <%
					rst2.movenext
					loop
					end if 
					rst2.close
				%>
                </select>
                </FONT> 
              <td width="11%" height="32"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="podate" >
                </font></td>
              <td width="6%" height="32"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="vendor">
                </font></td>
              <td width="14%" height="32"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="jobaddr"  >
                </font></td>
              <td width="14%" height="32"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="shipaddr" >
                </font></td>
              <td width="36%" height="32"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="req">
                  <%Set rst3 = Server.CreateObject("ADODB.recordset")
			sqlstr = "select [last name]+', '+ [first name] as name, substring(username,7,20) as user1, active from employees order by [last name]"
   			rst3.Open sqlstr, cnn1, 0, 1, 1
			if not rst3.eof then
				do until rst3.eof	
					if rst3("active") then 
				   %>	
				 
                  <option value="<%=rst3("user1")%>"><font face="Arial, Helvetica, sans-serif"><%=rst3("name")%></font></option>
                  <%
				  	end if
					rst3.movenext
					loop
					end if
					rst3.close
				%>
                </select>
                </font></td>
            </tr>
          </table>
          <input type="submit" name="choice" value="Save">
          <font face="Arial, Helvetica, sans-serif"><i> 
          <input type="button" name="Button22" value="BACK" onClick="Javascript:history.back()">
          </i></font></div>
      </td>
    </tr>
  </table>
</form>

</body>
</html>
