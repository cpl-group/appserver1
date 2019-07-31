<!-- #include file="./adovbs.inc" -->
<% 
leaseid= Request("leaseid")
profiletype=Request("profiletype")
portfolio=Request("portfolio")
%> 
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function periodfill(leaseid,bldg){
	document.location.href="meterservices.asp?leaseid=" + leaseid + "&bldg=" + bldg;
}
function loadinvoice(ypid,lid){
			var temp= "invoice.asp?y=" + ypid + "&l=" + lid 
			document.frames.invoice.location.href=temp;
	}
function print_invoice() {

document.invoice.focus();
document.invoice.print();

}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<div align="center">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td bgcolor="#0099FF"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="4" color="#000000">Meter 
          Services </font></div>
      </td>
    </tr>
  </table>
  
  <form method="post" action="" name="lmp">
    <div align="left"> 
      <input type="hidden" name="bldg" value=<%=Request("bldg")%>>
     </div>
    <table width="306" border="0" cellspacing="0" cellpadding="0" align="center">
      <tr> 
        <td height="37" width="101"> 
          <div align="left"><font face="Arial, Helvetica, sans-serif" size="3">Tenant</font></div>
        </td>
        <td height="37" width="92"> 
          <div align="left"><font face="Arial, Helvetica, sans-serif" size="3">Period</font></div>
        </td>
      </tr>
      <tr> 
        <td height="56" width="101"> 
          <div align="left">
            <%
		Set cnn1 = Server.CreateObject("ADODB.Connection")
		Set rst1 = Server.CreateObject("ADODB.recordset")
		cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
		%>
            <div align="center"> 
              <div align="left"></div>
              <div align="left"><font face="Arial, Helvetica, sans-serif" size="3"> 
			  
                <select name="leaseid" onChange="periodfill(this.value,bldg.value)">
				<OPTGROUP label='Building Tenants'>
                  <%
				strsql = "SELECT DISTINCT tenantnum, Billingname, leaseutilityid FROM tblBillByPeriod where LeaseUtilityId= N'" & leaseid & "' order by tenantnum"
 
				rst1.Open strsql, cnn1, adOpenStatic
				
				if not rst1.eof then
				%>
                  <option value=<%=rst1("leaseutilityid")%> selected>[<%=rst1("tenantnum") %>] 
                  <%=rst1("billingname") %> </option>
                  <%
				 else %>
				  <option selected>Select Tenant</option>
				<%
				end if
				rst1.close
				strsql = "SELECT DISTINCT tenantnum, Billingname, leaseutilityid FROM tblBillByPeriod WHERE BldgNum = N'" & Request("bldg") & "'  order by tenantnum"
 
				rst1.Open strsql, cnn1, adOpenStatic		
				
				if not rst1.EOF then
				%>
                  <option value=<%=rst1("leaseutilityid")%>>[<%=rst1("tenantnum") %>] 
                  <%=rst1("billingname") %></option>
                  <%
				if portfolio="1" then
				leaseid= rst1("leaseutilityid")
				end if
				rst1.movenext
				end if
				
				do until rst1.EOF 
			    %>
                  <option value=<%=rst1("leaseutilityid")%>>[<%=rst1("tenantnum") %>] 
                  <%=rst1("billingname") %></option>
                  <%
				rst1.movenext
			   	loop
			%>
                </select>
                </font></div>
            </div>
        </div>
        </td>
        <td height="56" width="92">
		<input type="hidden" name="lid" value="<%=leaseid%>" >
          <div align="left"><font face="Arial, Helvetica, sans-serif" size="3">
            <select name="ypid" onChange="loadinvoice(this.value,lid.value)">
			<OPTGROUP label='Bill Period'>
             <%
				if leaseid <> "" then
				
				strsql = "SELECT ypid,datestart, dateend from tblbillbyperiod where leaseutilityid=" & Leaseid
				response.write strsql
		 		rst1.close
				rst1.Open strsql, cnn1, adOpenStatic
				%>
				<option selected>Select Period</option>
				<%
				do until rst1.eof
			%>
              <option value=<%=rst1("ypid")%>><font face="Arial, Helvetica, sans-serif" size="3"><%=rst1("datestart")-1 %> to <%=rst1("dateend")%></font></option>
              <% 
				rst1.movenext
				loop
				 else %>
				  <option selected>Select Tenant</option>
				<%
				End if 
			%>
            </select>
            </font></div>
        </td>
      </tr>
      <tr> 
        <td height="56" width="101"> 
          <div align="center"><font face="Arial, Helvetica, sans-serif" size="3"> 
            </font> 
            <div align="left"><font face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif">
              <input type="button" name="Button" value="Print Invoice" onclick="print_invoice()">
              </font></font></div>
          </div>
        </td>
        <td height="56" width="92"> 
          <div align="center"> 
            <div align="left"></div>
            <div align="left"><font face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif" size="3"> 
              </font></font></font><font face="Arial, Helvetica, sans-serif" size="3"> 
              </font></div>
          </div>
        </td>
      </tr>
    </table>
  </form>
  <p align="left"><IFRAME name="invoice" src="null.htm" width="100%" height="100%" scrolling="auto" marginwidth="8" marginheight="16">
    <%
rst1.close
%>
    </IFRAME></p>
</div>
</body>
</html>
