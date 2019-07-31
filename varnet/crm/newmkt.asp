<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<%
user=Session("login")

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")

cnn1.Open application("cnnstr_main")


%>

<form name="form1" method="post" action="savemkt.asp">
  <table width="100%" border="0">
    <tr> 
      <td bgcolor="#3399CC" height="2"> 
        <table width="100%" border="0">
          <tr> 
            <td height="2"><i><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">New 
              Interaction Details: <%=ponum%> 
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
      <td height="2"><font face="Arial, Helvetica, sans-serif">Sales Manager : 
        <select name="manager">
          <%Set rst1 = Server.CreateObject("ADODB.recordset")
			sqlstr = "Select * from  salesmanagers order by manager"
   			rst1.Open sqlstr, cnn1, 0, 1, 1
			if not rst1.eof then
					do until rst1.eof
		%>
          <option value="<%=rst1("id") %>"><font face="Arial, Helvetica, sans-serif"><i><b><font color="#FFFFFF"><%=rst1("manager")%></font></b></i></font></option>
          <%
				 
					rst1.movenext
					loop
					end if
					rst1.close
				%>
        </select>
        </font></td>
    </tr>
    <tr> 
      <td height="2"> 
        <div align="left"> 
          <table width="100%" border="0">
            <tr bgcolor="#CCCCCC"> 
              <td width="26%"><font face="Arial, Helvetica, sans-serif">Customer 
                <input type="hidden" name="eb" value="<%=user%>">
                </font></td>
              <td width="26%"><font face="Arial, Helvetica, sans-serif">Situation</font></td>
            </tr>
            <tr> 
              <td width="26%" valign="top"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="cust">
                  <%Set rst1 = Server.CreateObject("ADODB.recordset")
			sqlstr = "Select first_name,last_name,company,id from contacts order by last_name"
   			rst1.Open sqlstr, cnn1, 0, 1, 1
			if not rst1.eof then
					do until rst1.eof
		%>
                  <option value="<%=rst1("id") %>"><font face="Arial, Helvetica, sans-serif"><i><b><font color="#FFFFFF"><%=rst1("last_name") %>, 
                  <%=rst1("first_name")%> (<%=rst1("company")%>)</font></b></i></font></option>
                  <%
				 
					rst1.movenext
					loop
					end if
					rst1.close
				%>
                </select>
                </font></td>
              <td width="26%" height="31" valign="top"> <font face="Arial, Helvetica, sans-serif"> 
                <textarea name="sit" cols="25" rows="3" wrap="PHYSICAL"></textarea>
                </font></td>
            </tr>
          </table>
          <input type="submit" name="choice" value="SAVE">
          <font face="Arial, Helvetica, sans-serif"><i> 
          <input type="button" name="Button22" value="BACK" onClick="Javascript:history.back()">
          </i></font></div>
      </td>
    </tr>
  </table>
</form>
</body>
</html>
