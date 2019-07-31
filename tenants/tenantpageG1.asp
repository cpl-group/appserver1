<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim tenantnum, bldg, leaseid, utility, bperiod, byear,tenantname,msg,pid
'tenantnum = getXMLUserName() N.Ambo removed for g1console purposes

tenantnum = request("tenantnum") 'N.Ambo added 4/23/2009, the tenantnum will be stated in teh url link on the pgi
'bldg = getKeyValue("bldg")
bldg = request("bldg") 'N.Ambo added becuase bldg number will be passed in the query string
leaseid = request("leaseid")
pid= request("pid")

dim cnn1, rst1, strsql
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

'2/5/2008 N.Ambo added getConnect connection
if pid="" or pid = 0 then
    cnn1.Open getLocalConnect(bldg)
else
    cnn1.Open getConnect(pid,bldg,"billing")
end if

dim lastdate, lastypid

if trim(leaseid)="" then
  'strsql = "SELECT DISTINCT leaseutilityid, tenantname FROM tblBillByPeriod where tenantnum= '" & tenantnum & "' and bldgnum='"& bldg&"' and tenantnum in (SELECT  tenantnum FROM tblleases where tenantnum= '" & tenantnum & "'  and onlinebill=1)"
  strsql = "SELECT LUP.leaseutilityid, tL.tName as TenantName " & _
			" FROM tblLeases tL " & _
			" INNER JOIN tblLeasesUtilityPrices LUP ON tL.BillingId = LUP.BillingId " & _
			" WHERE tL.tenantnum= '" & tenantnum & "' and tL.bldgnum='"& bldg&"'" 
				
  rst1.Open strsql, cnn1, adOpenStatic
  if not rst1.eof then 
  	if rst1.RecordCount > 1 then
   Response.write "<b> SELECT TENANT" &"<BR>"
    while not rst1.EOF
  	'"<a href='tenantpage.asp?tenantnum="&rst1("tenantnum")&"&bldg="&bldg&"&leaseid="&rst1("leaseutilityid")&">"&rst1("tenantname")&"</a>"
	response.write "<LI>" & "<a href='tenantpage.asp?pid="&pid&"&bldg="&bldg&"&leaseid="&rst1("leaseutilityid")&"'>"&rst1("tenantname")&"</a>"&"<BR>"
   
    rst1.movenext
   wend
	response.end
   	else
	leaseid=rst1("leaseutilityid")
	tenantname = rst1("tenantname")
    end if
 end if 
 rst1.close
end if

rst1.open "SELECT * FROM tblleasesutilityprices lup WHERE leaseutilityid='"&leaseid&"'", cnn1
if not rst1.eof then 
utility = rst1("utility")
end if
rst1.close
%>
<html>
<head>
<title>gEnergyOne - Tenant Utiltiy Bill </title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function periodfill(leaseid,bldg){
	document.location.href="meterservices.asp?leaseid=" + leaseid + "&bldg=" + bldg;
}
function loadinvoice(ypid,lid){
  ypid = ypid.split("|");
  var byear = ypid[1]
  var bperiod = ypid[2]
  var ypid = ypid[0]
  var temp= "http://pdfmaker.genergyonline.com<%=getKeyValue("billlink")%>genergy2=true&devIP=<%=request.ServerVariables("SERVER_NAME")%>&building=<%=bldg%>&lid="+lid+"&byear="+byear+"&bperiod="+bperiod+"&y="+ypid+"&ypid="+ypid+"&l="+lid+"&detailed=false&utilityid=<%=utility%>"

  document.frames.invoice.location.href=temp;
}
</script>
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#3399cc" class="standardheader">
    <td>Available Bills for <%=tenantname%> (<%=tenantnum%>)</td>
  </tr>
</table>
<form method="post" action="" name="lmp">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="pid" value="<%=pid%>" >
  &nbsp;Select the Utility Type: 
  <select name="leaseid" onchange="submit()">
<%
'strsql = "SELECT distinct lup.leaseutilityid, u.utilitydisplay FROM tblleasesutilityprices lup, tblleases l, ["&application("superip")&"].mainmodule.dbo.tblutility u WHERE lup.billingid=l.billingid and u.utilityid=lup.utility and tenantnum='"&tenantnum&"'"
strsql = "SELECT distinct lup.leaseutilityid, u.utilitydisplay FROM tblleasesutilityprices lup, tblleases l, ["&application("Coreip")&"].dbcore.dbo.tblutility u WHERE lup.billingid=l.billingid and u.utilityid=lup.utility and tenantnum='"&tenantnum&"'"
'response.write strsql
'response.end
rst1.Open strsql, cnn1, adOpenStatic
do until rst1.eof
  %><option value="<%=rst1("leaseutilityid")%>"<%
  if int(leaseid)=rst1("leaseutilityid") then
    response.write " SELECTED"
  end if
  %>><%=rst1("utilitydisplay")%></option><%
  rst1.movenext
loop
rst1.close
%>
</select>
  <%
if leaseid <> "" then
  strsql = "SELECT distinct ypid,datestart, dateend, billyear, billperiod from tblbillbyperiod where leaseutilityid=" & Leaseid & " and billyear >= '2002' order by dateend desc"
  rst1.Open strsql, cnn1, adOpenStatic
  if not rst1.EOF then
    lastdate = rst1("dateend")
    lastypid = rst1("ypid")
    byear = rst1("billyear")
    bperiod = rst1("billperiod")
    %>
  Select Bill Dates: 
  <input type="hidden" name="lid" value="<%=leaseid%>" >
        <select name="ypid">
        <%
        do until rst1.eof
          %><option value="<%=rst1("ypid")%>|<%=rst1("billyear")%>|<%=rst1("billperiod")%>"><%=rst1("datestart")-1 %> to <%=rst1("dateend")%></option><%
          rst1.movenext
        loop%>
        </select>
  <input type="button" name="Button" value="View Invoice" onclick="loadinvoice(ypid.value,lid.value)">
</form>  
    <IFRAME name="invoice" src="http://pdfmaker.genergyonline.com<%=getKeyValue("billlink")%>genergy2=true&devIP=<%=request.ServerVariables("SERVER_NAME")%>&building=<%=bldg%>&lid=<%=leaseid%>&byear=<%=byear%>&bperiod=<%=bperiod%>&y=<%=lastypid%>&ypid=<%=lastypid%>&l=<%=leaseid%>&detailed=false&pid=<%=pid%>&utilityid=<%=utility%>" width="100%" height="90%" scrolling="auto" marginwidth="8" marginheight="16"></IFRAME>
  <%else
    msg="No bills currently available"
    %>
    <table border="0" cellspacing="0" cellpadding="0"><tr><td align="center"><b><%=msg%></b></td></tr></table>
    <%
  End if 
  
  
else
  msg="No bills currently available"
  %>
  <table border="0" cellspacing="0" cellpadding="0">
  <tr><td align="center"><b><%=msg%></b></td></tr>
  </table>
<%end if
%>
</body>
</html>
