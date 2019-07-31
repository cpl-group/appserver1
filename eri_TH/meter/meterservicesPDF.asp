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
function tenantdetails(leaseid,bldg){
	document.frames.panel.location="leasehistoryPDF.asp?leaseid=" + leaseid + "&b=" + bldg;
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
</head><style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>


<body bgcolor="#FFFFFF" text="#000000">
<div align="left">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td bgcolor="#0099FF"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="4" color="#FFFFFF">Meter 
          Services </font></div>
      </td>
    </tr>
  </table>
  
  <form method="post" action="" name="lmp">
    <div align="left"> 
      <input type="hidden" name="bldg" value=<%=server.urlencode(Request("bldg"))%>>
     </div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
      <tr> 
        <td height="37" width="101"> 
          <div align="left"><font face="Arial, Helvetica, sans-serif" size="3">Tenant</font></div>
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
                <select name="leaseid" onChange="tenantdetails(this.value,bldg.value)">
                  <option>Select Tenant</option>
					<option value="0">All Tenants</option>
                  <OPTGROUP label='Building Tenants'> 
                  <%
				strsql = "SELECT DISTINCT tenantnum, Billingname, leaseutilityid FROM tblBillByPeriod where LeaseUtilityId= N'" & leaseid & "' order by tenantnum"
 
				rst1.Open strsql, cnn1, adOpenStatic
				
				if not rst1.eof then
					%><option value=<%=rst1("leaseutilityid")%> selected>[<%=rst1("tenantnum") %>] <%=rst1("billingname") %> </option><%
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
      </tr>
      <tr>
        <td height="300" width="101"><p align="left"><IFRAME name="panel" src="null.htm" width="100%" height="100%" scrolling="auto" marginwidth="8" marginheight="16"></IFRAME></p>
</td>
      </tr>
      <tr> 
        <td height="300" width="101"> 
          <div align="center"><font face="Arial, Helvetica, sans-serif" size="3"> 
            </font> 
            <div align="left"><font face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif"> 
              <p align="left"><IFRAME name="panel_2" src="null.htm" width="100%" height="100%" scrolling="auto" marginwidth="8" marginheight="16"></IFRAME></p>
              </font></font></div>
          </div>
        </td>
      </tr>
    </table>
  </form>
    
    
</div>
<%
rst1.close
%>  

</body>
</html>
