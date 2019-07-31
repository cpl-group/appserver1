<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim leaseid, profiletype, portfolio, bldg
leaseid= Request("leaseid")
profiletype=Request("profiletype")
portfolio=Request("portfolio")
bldg = Request("bldg")
dim cnn1, rst1, strsql
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getLocalConnect(request("bldg"))
dim DBmainIP
DBmainIP = "["&application("superIP")&"].mainmodule.dbo."
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
</script><meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">
</head>


<body bgcolor="#FFFFFF" text="#000000">
<div align="left">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td bgcolor="#6699CC"> 
        <div align="center"><span class="standardheader">Meter 
          Services </span></div>
      </td>
    </tr>
  </table>
  
  <form method="post" action="" name="lmp">
    <div align="left"> 
      <input type="hidden" name="bldg" value=<%=server.urlencode(bldg)%>>
     </div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
      <tr> 
        <td height="37" width="101"> 
          <div align="left"><font size="2">Tenant</font></div>
        </td>
      </tr>
      <tr> 
        <td height="56" width="101"> 
          <div align="left"> 
            <div align="center"> 
              <div align="left"></div>
              <div align="left"><font face="Arial, Helvetica, sans-serif" size="3"> 
                <select name="leaseid" onChange="tenantdetails(this.value,bldg.value)">
                  <option value="0">Select Tenant</option>
      <%
				strsql = "SELECT distinct 'u'+ltrim(utilityid) as uout, utilitydisplay FROM tblbillbyperiod b, "&DBmainIP&"tblutility u WHERE u.utilityid=b.utility AND bldgnum='" & bldg & "' order by utilityDisplay"
				rst1.Open strsql, cnn1, adOpenStatic		
				do until rst1.EOF 
				  %><option value=<%=rst1("uout")%>>All <%=rst1("utilitydisplay")%> Tenants</option><%
  				rst1.movenext
		   	loop
        rst1.close
			%> 
            <OPTGROUP label='Building Tenants'> 
      <%
				strsql = "SELECT DISTINCT utilitydisplay, b.tenantnum, l.Billingname, b.leaseutilityid, leaseexpired FROM tblBillByPeriod b, tblleasesutilityprices lup, tblleases l, "&DBmainIP&"tblutility u WHERE lup.billingid=l.billingid and b.leaseutilityid=lup.leaseutilityid and u.utilityid=lup.utility and b.BldgNum = N'" & bldg & "'  order by leaseexpired, b.tenantnum"
				rst1.Open strsql, cnn1, adOpenStatic		
				do until rst1.EOF 
				  %><option value="<%=rst1("leaseutilityid")%>" <%if lcase(rst1("leaseexpired"))="true" then%>style="color: Gray;"<%end if%> <%if trim(leaseid)=trim(rst1("leaseutilityid")) then response.write " SELECTED"%>>[<%=rst1("tenantnum") %>] Demo Tenant (<%=rst1("utilitydisplay")%>)<%if rst1("leaseexpired")="True" Then%>expired<%end if%></option><%
  				rst1.movenext
		   	loop
        rst1.close
			%>
                </select>
                </font></div>
            </div>
          </div>
        </td>
      </tr>
      <tr>
        <td height="300">
			<p align="left"><IFRAME name="panel" src="/null.htm" width="100%" height="100%" scrolling="auto" marginwidth="8" marginheight="16"></IFRAME></p>
		</td>
      </tr>
      <tr> 
        <td height="300">
			<p align="left"><IFRAME name="panel_2" src="/null.htm" width="100%" height="100%" scrolling="auto" marginwidth="8" marginheight="16"></IFRAME></p>
        </td>
      </tr>
    </table>
  </form>
</div>
</body>
</html>
