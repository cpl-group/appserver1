<!-- #include file="./adovbs.inc" -->
<% 
leaseid= Request("leaseid")
profiletype=Request("profiletype")
user=session("loginemail")
%> 
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function meterfill(leaseid,bldg, profiletype){

	document.location.href="lmp2.asp?leaseid=" + leaseid + "&bldg=" + bldg + "&profiletype=" + profiletype;
	
}
function loadmeter(){
			var id =document.forms.lmp.meter.value
			var bldg=document.forms.lmp.bldg.value
			if (id == ""){
				alert("Please Select a Tenant")
			}else{
			var temp= "lmpindex.asp?m=" + id +"&b=" + bldg 
			document.frames.lmp.location.href=temp;
			}
	}
function loadlmp(){
			var id =document.forms.lmp.bldglmp.value
			var bldg=document.forms.lmp.bldg.value
			var temp= "lmpindex.asp?m=" + id +"&b=" + bldg + "&lmp=1"
			document.frames.lmp.location.href=temp;
	}

function AgrTenant(){
			var id =document.forms.lmp.leaseid.value
			var bldg=document.forms.lmp.bldg.value
			if (id == ""){
				alert("Please Select a Tenant")
			}else{
			var temp= "lmpindex.asp?luid=" + id +"&b=" + bldg 
			document.frames.lmp.location.href=temp;
			}
	}
	
function navigate(direc,bldg,lid,profiletype){
    var currdate = new Date()
	var user=document.forms[0].user.value
	//alert(document.frames.lmp.datefield.value)
	if (direc == "prev") {
	
		currdate=new Date(currdate).valueOf() - (86400000)
		}else{
		
		currdate=new Date(currdate).valueOf() + (86400000)
	    }
	
	currdate = new Date(currdate)
	document.forms[0].currentid.value=lid
	var script =document.forms[0].currentscript.value
	currdate = (currdate.getMonth() + 1) + "/" + currdate.getDate() + "/" + currdate.getFullYear()
	
		var temp = "http://www.genergy.com/cgi-bin/" + script + "?start=01:00&end=24:00&interval=60&graph=bar3d&bldg=" + bldg + "&lid="+ lid+"&dsn=sqlserverg1&user="+ user
	//document.lmp.navigate(temp)
	//document.frames.lmp.location.href=temp;
    document.frames.lmp.ok()
}
</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000" alink="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#0099FF"> 
      <div align="center"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0" width="118" height="20">
          <param name=movie value="text1.swf">
          <param name=quality value=high>
          <param name="BGCOLOR" value="#0099FF">
          <param name="SCALE" value="exactfit">
          <embed src="text1.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" scale="exactfit" width="118" height="20" bgcolor="#0099FF">
          </embed> 
        </object></div>
    </td>
  </tr>
</table>
<form name="lmp" method="post" action="">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
        <table width="335" border="0" cellspacing="0" cellpadding="2" align="left">
          <tr> 
            <td height="56" width="69"><b><font face="Arial, Helvetica, sans-serif" size="3">Tenants</font></b></td>
            <td height="56" width="117"> 
              <%
		Set cnn1 = Server.CreateObject("ADODB.Connection")
		Set rst1 = Server.CreateObject("ADODB.recordset")
		cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
		%>
              <font face="Arial, Helvetica, sans-serif" size="3"> 
              <input type="hidden" name="bldg" value="<%=Request("bldg")%>">
              <input type="hidden" name="profiletype" value="<%=profiletype %>">
              <input type="hidden" name="prev" value="prev">
              <input type="hidden" name="next" value="next">
              <select name="leaseid" onChange="meterfill(this.value,bldg.value,profiletype.value)">
                <%
			  if leaseid <> "" then
				strsql = "SELECT tblLeases.BillingName, tblLeasesUtilityPrices.LeaseUtilityId, tblLeases.TenantNum FROM tblLeases INNER JOIN tblLeasesUtilityPrices ON tblLeases.BillingId = tblLeasesUtilityPrices.BillingId WHERE (tblLeasesUtilityPrices.LeaseUtilityId= N'" & leaseid & "') order by tblLeases.billingname"
				rst1.Open strsql, cnn1, adOpenStatic
	
				
				if not rst1.eof then
				%>
                <option value=<%=rst1("leaseutilityid")%> selected>Demo Tenant 
                </option>
                <%	
				end if
				
				rst1.close
				
				else
				
				%>
                <option selected>Select Tenant</option>
                <%
				
				end if
				
				if leaseid <> "" then
				
				strsql = "SELECT DISTINCT tblLeases.TenantNum, tblLeases.tName, tblLeasesUtilityPrices.LeaseUtilityId FROM tblLeases INNER JOIN tblLeasesUtilityPrices ON tblLeases.BillingId = tblLeasesUtilityPrices.BillingId INNER JOIN pulse_" & Request("bldg") & " INNER JOIN Meters ON pulse_" & Request("bldg") & ".meterid = Meters.MeterId ON tblLeasesUtilityPrices.LeaseUtilityId = Meters.LeaseUtilityId WHERE (tblLeases.BldgNum = N'" & Request("bldg") & "') AND (Meters.PP <> 1) AND (meters.meterid = pulse_" & Request("bldg") & ".meterid)and tblLeasesUtilityPrices.LeaseUtilityId <> " & leaseid & " GROUP BY tblLeases.BillingId, tblLeases.TenantNum, tblLeases.tName, tblLeasesUtilityPrices.LeaseUtilityId, Meters.MeterId"
				else
				
				strsql = "SELECT DISTINCT tblLeases.TenantNum, tblLeases.tName, tblLeasesUtilityPrices.LeaseUtilityId FROM tblLeases INNER JOIN tblLeasesUtilityPrices ON tblLeases.BillingId = tblLeasesUtilityPrices.BillingId INNER JOIN pulse_" & Request("bldg") & " INNER JOIN Meters ON pulse_" & Request("bldg") & ".meterid = Meters.MeterId ON tblLeasesUtilityPrices.LeaseUtilityId = Meters.LeaseUtilityId WHERE (tblLeases.BldgNum = N'" & Request("bldg") & "') AND (Meters.PP <> 1) AND (meters.meterid = pulse_" & Request("bldg") & ".meterid) GROUP BY tblLeases.BillingId, tblLeases.TenantNum, tblLeases.tName, tblLeasesUtilityPrices.LeaseUtilityId, Meters.MeterId"

				end if 
  response.write strsql
				rst1.Open strsql, cnn1, adOpenStatic		
				Dim Tenants(10)
				tenants(1) = "Demo Tenant 1"
				tenants(2) = "Demo Tenant 2"
				tenants(3) = "Demo Tenant 3"
				tenants(4) = "Demo Tenant 4"
				tenants(5) = "Demo Tenant 5"
				democount = 1
				do until democount > 5 or rst1.EOF
			    %>
                <option value=<%=rst1("leaseutilityid")%>><%=Tenants(democount)%></option>
                <%
				democount = democount + 1
				rst1.movenext
			   	loop
							
				
			%>
              </select>
              </font></td>
            <td height="56" width="55"><font face="Arial, Helvetica, sans-serif" size="2"><b><font size="3">Meters</font></b></font></td>
            <td height="56" width="55"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0" width="47" height="16">
                <param name=movie value="meters.swf">
                <param name=quality value=high>
                <param name="BGCOLOR" value="">
                <param name="SCALE" value="exactfit">
                <embed src="meters.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" scale="exactfit" width="47" height="16" bgcolor="">
                </embed> 
              </object></td>
            <td height="56" width="81"> <font face="Arial, Helvetica, sans-serif" size="3"> 
              <%
				strsql = "SELECT meterid, lmnum from meters where bldgnum='" & Request("bldg") & "' and pp=1 order by meternum"
				
		 		rst1.close
				rst1.Open strsql, cnn1, adOpenStatic
			    if not rst1.eof then
				lmpavail = 1
				%>
              <input type="hidden" name="bldglmp" value="<%=rst1("meterid")%>" >
              <% End if 
					%>
              <select name="meter" onChange="loadmeter()">
                <%rst1.close
				if leaseid <> "" then
				strsql = "SELECT meterid, meternum, lmnum from meters where (LeaseUtilityId=" & leaseid & "and online=1 and lmnum is not NULL) or (leaseUtilityId=" & leaseid & "and online=1 and EXISTS (select * from tblLeasesUtilityPrices where LeaseUtilityId=" & leaseid &  " and LoadProfile=1))order by meternum"

				rst1.Open strsql, cnn1, adOpenStatic
				do until rst1.eof
				if rst1("lmnum") <> "pt" then 
				
			%>
                <option value="<%=rst1("meterid")%>"><%=rst1("meternum") %></option>
                <% 
				else
			%>
                <option value="<%=rst1("meterid")%>"><%=rst1("meternum") %></option>
                <%
				end if 
				rst1.movenext
				loop
				End if 
			%>
              </select>
              </font></td>
          </tr>
        </table>
        
      </td>
  </tr>
  <tr>
    <td>
        <table width="300" border="0" align="left" cellpadding="0" cellspacing="0">
          <tr valign="bottom"> 
            <td width="101"> <font face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif" size="3"> 
              <%if lmpavail = 1 then %>
              <a href="javascript:loadlmp()" style="text-decoration:none;" onMouseOver="this.style.color
        = 'gray'"; onMouseOut="this.style.color = 'Black'"><b><font size="2">Building 
              Profile</font></b></a> 
              <%end if %>
              </font></font></font></td>
            <td width="98"><font face="Arial, Helvetica, sans-serif" size="2"><b> | <a href="javascript:loadmeter()" style="text-decoration:none;" onMouseOver="this.style.color= 'gray'"; onMouseOut="this.style.color = 'Black'">Meter 
              Profile</a></b></font></td>
            <td width="101"> <b><font face="Arial, Helvetica, sans-serif" size="2">| 
              <a href="javascript:AgrTenant()" style="text-decoration:none;" onMouseOver="this.style.color= 'gray'"; onMouseOut="this.style.color = 'Black'">Tenant Profile</a></font></b></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</form>
<p><IFRAME name="lmp" src="null.htm" width="100%" height="100%" scrolling="auto" marginwidth="8" marginheight="16"></IFRAME>        
</p></body>
</html>
