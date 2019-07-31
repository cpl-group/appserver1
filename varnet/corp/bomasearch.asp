<%@Language="VBScript"%>
<%

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=security;"

sqlstr = " Select * from tblIP"

rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1
%>
<body bgcolor="#FFFFFF">
<form name="form1" method="post" action=""><table width="100%" border="0">
  <tr>
    <td bgcolor="#3399CC" height="36" width="13%"> 
	 
	    <div align="center"> <font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF" size="4"> 
          SEARCH RESULTS </font></b></font></div>
	</td>  
  </tr>
  <tr>
   
    <td width="87%"> 
        <table width="100%" border="0">
          <tr bgcolor="#CCCCCC"> 
            <td bgcolor="#CCCCCC" width="9%"><b><font face="Arial, Helvetica, sans-serif" color="#000000" size="2">Email 
              Account </font></b></td>
            <td width="19%"><b><font face="Arial, Helvetica, sans-serif" color="#000000" size="2">Number 
              of Visits to BOMA Site</font></b></td>
            <td width="18%"><b><font face="Arial, Helvetica, sans-serif" color="#000000" size="2">IP 
              Address </font></b></td>
          </tr>
          <% While not rst1.EOF %>
          <tr> 
            <td width="9%"><font face="Arial, Helvetica, sans-serif"><%=rst1("email")%></a></font></td>
            <td width="19%"><font face="Arial, Helvetica, sans-serif"><%=rst1("count")%></font></td>
            <td width="18%"><font face="Arial, Helvetica, sans-serif"><%=rst1("ip")%></font></td>
          </tr>
          <%
		rst1.movenext
		Wend
		%>
        </table>
    </td>
  </tr>
   </table></form>
<%
rst1.close
%>
</body>
