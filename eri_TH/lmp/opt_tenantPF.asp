<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim b, profiletype, leaseid, user, m, luid
b = request.querystring("b")
m = request.querystring("m")
luid = request.querystring("luid")
leaseid= Request("leaseid")
profiletype=Request("profiletype")
user=session("loginemail")

if trim(leaseid)="" and trim(luid)<>"" then leaseid = luid
dim tenantmeter
dim cnn1, rst1
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_genergy1")
if session("roleid")=1 then
    if leaseid="" then 'need to get the leaseid before hand to display the tenant info properly
        rst1.open "SELECT LeaseUtilityId FROM tblLeases INNER JOIN tblLeasesUtilityPrices ON tblLeases.BillingId = tblLeasesUtilityPrices.BillingId WHERE tblLeases.TenantNum='"&session("userid")&"'", cnn1
        leaseid=rst1("LeaseUtilityId")
        rst1.close()
    end if
end if
%>

<html>
<head>
<title>Untitled Document</title>
<script>
function meterfill(leaseid,b, profiletype)
{   parent.document.forms[0].pd.value=""
    parent.document.forms[0].nd.value=""
	parent.document.forms[0].d.value=""
    var lmp = "";
    if(!leaseid) lmp=1;
	parent.document.forms[0].lmp.value=lmp
	parent.document.forms[0].tenantmeter.value=0
	parent.document.forms[0].luid.value=leaseid
	parent.document.forms[0].m.value='<%=m%>'
	if(parent.document.all['tab3'].style.backgroundColor=="#0099ff") parent.loadcalendar();else parent.loadchart();
  parent.openLoadBox('loadFrame2')
	document.location.href="opt_tenantPF.asp?luid=" + leaseid + "&b=" + b + "&m=<%=m%>" + "&profiletype=" + profiletype;
}

function loadmeter(meterid)
{   //parent.document.forms[0].d.value=""
    //parent.document.forms[0].pd.value=""
    //parent.document.forms[0].nd.value=""
    if(parent.document.all['tab3'].style.backgroundColor=="#0099ff") parent.loadcalendar();else parent.loadchart();
    parent.openLoadBox('loadFrame1')
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

<body bgcolor="#FFFFFF" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" onload="parent.closeLoadBox('loadFrame2');">
<form name="lmp" method="post" action="">
  <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000">
    <tr>
      <td width="48%"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#FFFFFF">Other 
        Load Profiles</font></b></font></td>
      <td width="52%">
        <div align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b><a href="javascript:document.location.href='<%="options2.asp?b=" & b & "&m=" & m &"&luid="& leaseid%>'" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Return To Options</a></b></font></div>
      </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100">
    <tr> 
      <td width="79%" valign="top"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="2">
          <tr valign="top"> 
            <td height="2" width="789"><font face="Arial, Helvetica, sans-serif" size="3" color="#000000">Select Profile Below:</font></td>
          </tr>
          <tr valign="top"> 
            <td height="14" width="789"> <font face="Arial, Helvetica, sans-serif" size="3" color="#000000"> 
              <input type="hidden" name="b" value="<%=b%>">
              <input type="hidden" name="profiletype" value="<%=profiletype%>">
              <input type="hidden" name="prev" value="prev">
              <input type="hidden" name="next" value="next">
    <%if session("roleid")=1 then' output just a tenant name
        response.write "<input value="""& luid &""" name=""leaseid"" type=""hidden"">"
'        response.write "<nobr><a href=""lmpload2.asp?b="& b &"&luid="& luid &"&m="& m &""">"
        response.write "<nobr>"
        dim rstrole
        Set rstrole = Server.CreateObject("ADODB.recordset")
        rstrole.open "select billingname from tblLeases where TenantNum='"&session("userid")&"'", cnn1
        response.write rstrole("billingname")
        rstrole.close()
        response.write "</nobr>"
    else 'create select box%>
        <select name="leaseid" onChange="meterfill(leaseid.value, b.value, profiletype.value)">
        <%
        dim strsql
        strsql = "SELECT DISTINCT tblLeases.TenantNum, tblLeases.tName, tblLeasesUtilityPrices.LeaseUtilityId FROM tblLeases INNER JOIN tblLeasesUtilityPrices ON tblLeases.BillingId = tblLeasesUtilityPrices.BillingId INNER JOIN pulse_" & b & " INNER JOIN Meters ON pulse_" & b & ".meterid = Meters.MeterId ON tblLeasesUtilityPrices.LeaseUtilityId = Meters.LeaseUtilityId WHERE (tblLeases.BldgNum = N'" & b & "') AND (Meters.PP <> 1) AND (meters.meterid = pulse_" & b & ".meterid) and leaseexpired = 0 GROUP BY tblLeases.BillingId, tblLeases.TenantNum, tblLeases.tName, tblLeasesUtilityPrices.LeaseUtilityId, Meters.MeterId"
        rst1.Open strsql, cnn1, adOpenStatic
        do until rst1.EOF
            response.write "<option value="""& rst1("leaseutilityid") &""""
            if trim(leaseid)=trim(rst1("leaseutilityid")) then response.write " selected"
            response.write ">"& rst1("tenantnum") &" - "& rst1("tname") &"</option>"
            rst1.movenext
        loop
        rst1.close
        response.write "</select>"
    end if
    response.write "<BR>"
dim cnn2, rst2



Set cnn2 = Server.CreateObject("ADODB.Connection")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn2.Open application("cnnstr_security")
rst2.open "SELECT Label, Link, Target, Active, tbladdonlinks.SID FROM tbladdonlinks JOIN tbladdons on tbladdons.SID=tbladdonlinks.SID WHERE userid='"&session("userid")&"' and CSID=4 ORDER BY listorder", cnn2
if rst2.eof then response.write "Client has no options"
do while not(rst2.eof)
    if rst2("SID")<>1 and rst2("SID")<>5 then response.write "<a href="""&rst2("Link")&"?m="& m &"&b="& b &"&luid="& leaseid &""" style=""color:black"" onMouseOver=""this.style.color='gray'"" onMouseOut=""this.style.color='Black'"" onclick=""parent.openLoadBox('loadFrame2');""><font size=""2"" face=""arial"">"&rst2("Label")&"</font></a><br>"
    rst2.movenext
loop
rst2.close
%>
              </font></td>
          </tr>
          &nbsp;<br>
        </table>
      </td>
      <td width="21%" valign="top" bgcolor="#CCCCCC"> 
	  <table width="100%">
		  <%if leaseid <> "" then %>
            <tr> 
              <td bgcolor="#0099FF" valign="top"> <div align="center"><font color="#000000" face="Arial, Helvetica, sans-serif" size="1"><b><font face="Arial, Helvetica, sans-serif" size="3"> 
                  </font><font color="#FFFFFF">Meter List</font><font face="Arial, Helvetica, sans-serif" size="3"> 
                  </font></b></font></div></td>
            </tr>
            <%end if%>
		</table>
        <div name="meterlist" style="border-width:1; width:100%; height:150;overflow:auto" > 
          <table border="0" cellpadding="0" cellspacing="0" width="101%" bgcolor="#FFFFFF">
            <%
				if leaseid <> "" then
    				'strsql = "SELECT meterid, meternum, lmnum from meters where (LeaseUtilityId=" & leaseid & "and online=1 and lmnum is not NULL) or (leaseUtilityId=" & leaseid & "and online=1 and EXISTS (select * from tblLeasesUtilityPrices where LeaseUtilityId=" & leaseid &  " and LoadProfile=1))order by meternum"
    				strsql = "SELECT DISTINCT meters.meterid, p.meterid as hasdata, meternum, lmnum from meters LEFT OUTER JOIN pulse_"& b &" p on meters.meterid=p.meterid where (LeaseUtilityId=" & leaseid & "and online=1 and lmnum is not NULL and meters.pp<>1) or (leaseUtilityId=" & leaseid & "and online=1 and EXISTS (select * from tblLeasesUtilityPrices where LeaseUtilityId=" & leaseid &  " and LoadProfile=1 and meters.pp<>1)) and meters.pp<>1 order by meternum"
'    				response.write strsql
'					response.end
					rst1.Open strsql, cnn1, adOpenStatic
	    			do until rst1.eof
        				dim italic, italicx 'this is for italisizing the meter link if it has data
	    			    italic=""
	    			    italicx=""
	    			    if IsNumeric(trim(rst1("hasdata"))) then 'if hasdata has a meter number in it then pulse has a data set for that meter number
	    			        italic="<u>"
	    			        italicx="</u>"
	    			    end if%>
            <tr> 
              <td onMouseOver="this.style.background='lightgreen'" onMouseOut="this.style.background='white'" height="19"><a <%if IsNumeric(trim(rst1("hasdata"))) then%>onClick="parent.document.forms[0].m.value='<%=rst1("meterid")%>';parent.document.forms[0].b.value='<%=b%>';parent.document.forms[0].tenantmeter.value='1';parent.document.forms[0].luid.value='<%=luid%>';loadmeter(<%=rst1("meterid")%>);" <%end if%> target="lmp" style="text-decoration:none; color:black; font-size:12px; font-family:arial"><%= italic &rst1("meternum")& italicx %></a></td>
            </tr>
            <%rst1.movenext
		    		loop
				End if%>
          </table>
          
        </div>
        <div align="center"><a href="javascript:meterfill(document.forms['lmp'].leaseid.value, document.forms['lmp'].b.value, document.forms['lmp'].profiletype.value)" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'"><font color="#000000" face="Arial, Helvetica, sans-serif" size="1"><b><u>View 
          Aggregated Profile</u></b></font></a><br>
        </div></td>
    </tr>
	<tr><tr colspan="2"></tr></tr>
    <tr>
      <td height="20" width="789" valign="bottom" bgcolor="#000000"> 
        <div align="left">&nbsp;</div>
      </td>
      <td width="21%" valign="top" bgcolor="#000000">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>