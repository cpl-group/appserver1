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
'dim DBmainIP
'DBmainIP = "["&application("superIP")&"].mainmodule.dbo."
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
      <td bgcolor="#0099FF"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="4" color="#FFFFFF">Meter 
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
        <td height="37" width="101"> 
          <div align="left"><font face="Arial, Helvetica, sans-serif" size="3">Tenant</font></div>
        </td>
      </tr>
      <tr> 
        <td height="56"> 
          <div align="left"> 
            <div align="center"> 
              <div align="left"></div>
              <div align="left"><font face="Arial, Helvetica, sans-serif" size="3">
							<%dim isTenantonly
							if trim(leaseid)<>"" then isTenantonly = true
				if not(isTenantonly) then%>
                <select name="leaseid" onChange="tenantdetails(this.value,bldg.value)">
                  <option value="0">Select Tenant</option>
						      <%
										strsql = "SELECT distinct 'u'+ltrim(utilityid) as uout, utilitydisplay FROM tblbillbyperiod b,tblutility u WHERE u.utilityid=b.utility AND bldgnum='" & bldg & "' order by utilityDisplay"'dbmain b4 tblutility
										rst1.Open strsql, cnn1, adOpenStatic		
										do until rst1.EOF 
										  %><option value=<%=rst1("uout")%>>All <%=rst1("utilitydisplay")%> Tenants</option><%
						  				rst1.movenext
								   	loop
						        rst1.close
									%> 
            <OPTGROUP label='Building Tenants'> 
      <%
			end if
				strsql = "SELECT DISTINCT utilitydisplay, l.tenantnum, l.Billingname, lup.leaseutilityid, leaseexpired FROM tblleasesutilityprices lup, tblleases l, tblutility u WHERE lup.billingid=l.billingid and u.utilityid=lup.utility and l.BldgNum = N'"&bldg&"' ORDER BY leaseexpired, l.tenantnum"
				rst1.Open strsql, cnn1, adOpenStatic		
				do until rst1.EOF 
				  if not(isTenantonly) then %><option value="<%=rst1("leaseutilityid")%>" <%if lcase(rst1("leaseexpired"))="true" then%>style="color: Gray;"<%end if%> <%if trim(leaseid)=trim(rst1("leaseutilityid")) then response.write " SELECTED"%>>[<%=rst1("tenantnum") %>] <%if demo then %>Demo Tenant<%else%><%=rst1("billingname")%><%end if%> (<%=rst1("utilitydisplay")%>)<%if rst1("leaseexpired")="True" Then%>expired<%end if%></option><%
					elseif trim(rst1("leaseutilityid"))=trim(leaseid) then%><input type="hidden" value="<%=rst1("leaseutilityid")%>" name="leaseid">[<%=rst1("tenantnum") %>]&nbsp;<%=rst1("billingname")%>&nbsp;(<%=rst1("utilitydisplay")%>)<%end if
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
			<p align="left"><IFRAME name="panel" src="/mymain.htm" width="100%" height="100%" scrolling="auto" marginwidth="8" marginheight="16"></IFRAME></p>
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
<script>
<%if isTenantonly and trim(leaseid)<>"" then%>tenantdetails(<%=leaseid&",'"&bldg&"'"%>)<%end if%>
</script>
</body>
</html>
