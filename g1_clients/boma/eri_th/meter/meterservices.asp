<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<% 
leaseid= Request("leaseid")
profiletype=Request("profiletype")
portfolio=Request("portfolio")
%> 
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">		
<script>
function tenantdetails(leaseid,bldg){
	document.frames.panel.location="leasehistory.asp?leaseid=" + leaseid + "&b=" + bldg;
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
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      
    <td bgcolor="#6699cc"><span class="standardheader"><font size="2">Meter Services</font> 
      </span> </td>
    </tr>
  </table>
  
  <form method="post" action="" name="lmp">
      <input type="hidden" name="bldg" value=<%=Request("bldg")%>>
    
  <table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
    <tr> 
      <td height="56" width="100">Select Tenant</td>
      <td width="1015">
        <%
		Set cnn1 = Server.CreateObject("ADODB.Connection")
		Set rst1 = Server.CreateObject("ADODB.recordset")
		cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
		%>
        <select name="leaseid" onChange="tenantdetails(this.value,bldg.value)"><optgroup label='Building Tenants'> 
          <%
				Dim Tenants(10)
				tenants(1) = "Demo Tenant 1"
				tenants(2) = "Demo Tenant 2"
				tenants(3) = "Demo Tenant 3"
				tenants(4) = "Demo Tenant 4"
				tenants(5) = "Demo Tenant 5"
				strsql = "SELECT DISTINCT tenantnum, Billingname, leaseutilityid FROM tblBillByPeriod where LeaseUtilityId= N'" & leaseid & "' order by tenantnum"
 
				rst1.Open strsql, cnn1, adOpenStatic
				
				if not rst1.eof then
				%>
          <option value=<%=rst1("leaseutilityid")%> selected>[<%=rst1("tenantnum") %>] 
          <%=rst1("billingname") %> </option>
          <%
				end if
				rst1.close
				democount = 1
				strsql = "SELECT DISTINCT tenantnum, Billingname, leaseutilityid FROM tblBillByPeriod WHERE BldgNum = N'" & Request("bldg") & "'  order by tenantnum"
 
				rst1.Open strsql, cnn1, adOpenStatic		
				
				if not rst1.EOF then
				%>
          <option value=<%=rst1("leaseutilityid")%>><%=Tenants(democount)%></option>
          <%
				if portfolio="1" then
				leaseid= rst1("leaseutilityid")
				end if
				rst1.movenext
				end if
				democount=democount + 1
				do until democount>5 
			    %>
          <option value=<%=rst1("leaseutilityid")%>><%=Tenants(democount)%></option>
          <%
				democount=democount + 1
				rst1.movenext
			   	loop
			%>
        </select></td>
    </tr>
    <tr> 
      <td height="300" colspan="2"> <IFRAME name="panel" src="null.htm" width="100%" height="100%" scrolling="auto" marginwidth="8" marginheight="16"></IFRAME> 
      </td>
    </tr>
    <tr> 
      <td height="300" colspan="2"> <IFRAME name="panel_2" src="null.htm" width="100%" height="100%" scrolling="auto" marginwidth="8" marginheight="16"></IFRAME></p> 
      </td>
    </tr>
  </table>
  </form>
    
    
<%
rst1.close
%>  

</body>
</html>
