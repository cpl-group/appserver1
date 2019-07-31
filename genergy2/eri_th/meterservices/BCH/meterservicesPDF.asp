<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim leaseid, profiletype, portfolio, bldg,demo
leaseid= Request("leaseid")
profiletype=Request("profiletype")
portfolio=Request("portfolio")
bldg = Request("bldg")
demo = Request("demo")

'Show Demo
if demo = "" then demo = false else demo = true end if

dim cnn1, rst1, strsql
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getLocalConnect(request("bldg"))

%>

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function tenantdetails(leaseid,bldg){
	document.frames.panel.location="leasehistoryPDF.asp?demo=<%=demo%>&leaseid=" + leaseid + "&b=" + bldg;
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
      <td bgcolor="#336699"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="4" color="#FFFFFF">Boston Childrens Hospital Custom Meter 
          Services </font></div>
      </td>
    </tr>
  </table>
  
  <form method="post" action="" name="lmp">
    <div align="left"> 
      <input type="hidden" name="bldg" value=<%=server.urlencode(bldg)%>>
    </div>
    <table width="100%" border="0" cellspacing="0" cellpadding="0" align="left">
      <tr> 
        <td height="37"> 
          <font face="Arial, Helvetica, sans-serif" size="3">Tenant List&nbsp;&nbsp;
            <%dim isTenantonly
							if trim(leaseid)<>"" then isTenantonly = true
				if not(isTenantonly) then%>
            <select name="leaseid" onChange="tenantdetails(this.value,bldg.value)">
              <option value="0">Select Tenant</option>
			  <option value="A">All Tenant Invoices</option>
              <%
										strsql = "SELECT distinct 'u'+ltrim(utilityid) as uout, utilitydisplay FROM tblbillbyperiod b, tblutility u WHERE u.utilityid=b.utility AND bldgnum='" & bldg & "' order by utilityDisplay"
	'									response.write strsql
	'									response.end
										rst1.Open strsql, cnn1, adOpenStatic		
										do until rst1.EOF 
										  %>
              <option value=<%=rst1("uout")%>>All <%=rst1("utilitydisplay")%> Tenants</option>
              <%
						  				rst1.movenext
								   	loop
						        rst1.close
									%>
              <optgroup label='Building Tenants'>
              <%
			end if
				Dim currentBillingname,currentbillingid
				
				strsql = "SELECT DISTINCT utilitydisplay, l.tenantnum, l.Billingname, lup.billingid, lup.leaseutilityid, leaseexpired FROM tblleasesutilityprices lup, tblleases l, tblutility u WHERE lup.billingid=l.billingid and u.utilityid=lup.utility and l.BldgNum = N'"&bldg&"' ORDER BY leaseexpired, l.tenantnum"
									response.write "<br>"&strsql
									'	response.end
				rst1.Open strsql, cnn1, adOpenStatic		
				do until rst1.EOF 
				
					if trim(currentbillingid) <> trim(rst1("billingid")) and currentbillingid <> "" then 
						%>
							<option value="A_<%=trim(currentBillingid)%>" style="color: #336699;">All <%=currentBillingname%> Invoice</option>
						<%
						
					end if

				  if not(isTenantonly) then %>
              <option value="<%=rst1("leaseutilityid")%>" <%if lcase(rst1("leaseexpired"))="true" then%>style="color: Gray;"<%end if%> <%if trim(leaseid)=trim(rst1("leaseutilityid")) then response.write " SELECTED"%>>[<%=rst1("tenantnum") %>]
              <%if demo then %>
              Demo Tenant
              <%else%>
              <%=rst1("billingname")%>
              <%end if%>
			  (<%=rst1("utilitydisplay")%>)
			  <%if rst1("leaseexpired")="True" Then%>
			  expired
			  <%end if%>
              </option>
              <%
					elseif trim(rst1("leaseutilityid"))=trim(leaseid) then%>
              <input type="hidden" value="<%=rst1("leaseutilityid")%>" name="leaseid">
              [<%=rst1("tenantnum") %>]&nbsp;<%=rst1("billingname")%>&nbsp;(<%=rst1("utilitydisplay")%>)
              <%end if
  				currentBillingname = rst1("billingname")
				currentBillingid = rst1("billingid")
				rst1.movenext
		   	loop
			
			
        rst1.close
			%>
				<option value="A_<%=trim(currentBillingid)%>" style="color: #336699;">All <%=currentBillingname%> Invoice</option>
            </select>
</font>
        </td>
      </tr>
      <tr>
        <td height="5">&nbsp;</td>
      </tr>
      <tr>
        <td height="10" valign="middle"><font face="Arial, Helvetica, sans-serif" size="3">Bill Period Summary | <strong><font size="1"><i>When Data Loads, click ANY BILL PERIOD ROW To View Meter Details</i></font></strong></font></td>
      </tr>
      <tr>
        <td height="300">
			<p align="left"><IFRAME name="panel" src="/null.htm" width="100%" height="100%" scrolling="auto" marginwidth="8" marginheight="16" frameborder="0" style="border:1px solid #336699"></IFRAME></p>
		</td>
      </tr>
      <tr>
        <td height="10"><font face="Arial, Helvetica, sans-serif" size="3">Bill Period Details</font></td>
      </tr>
      <tr> 
        <td height="300">
			<p align="left"><IFRAME name="panel_2" src="/null.htm" width="100%" height="100%" scrolling="auto" marginwidth="8" marginheight="16" frameborder="0" style="border:1px solid #336699"></IFRAME></p>
        </td>
      </tr>
    </table>
  </form>
</div>
<script>
<%if isTenantonly and trim(leaseid)<>"" then%>tenantdetails(<%=leaseid&",'"&bldg&"'"%>)<%end if%>
</script>
</body>
</html>
