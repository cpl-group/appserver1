<!-- #include file="./adovbs.inc" -->
<% 
leaseid= Request("leaseid")
profiletype=Request("profiletype")
%> 
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function meterfill(leaseid,bldg, profiletype){

	document.location.href="lmp.asp?leaseid=" + leaseid + "&bldg=" + bldg + "&profiletype=" + profiletype;
}
function loadmeter(id,bldg){
			var sa=id.split("_")
			var id = sa[0]
			var script=sa[1]
			var now=new Date()
			document.forms[0].datefield.value=now.toLocaleString()
			document.forms[0].currentid.value=id
			if (script=="pt"){
				script="ptstart.cgi"
				document.forms[0].currentscript.value="ptstart.cgi"				
				dsn="sqlserverg1"
			}else{
				script="meterstart.cgi"
				document.forms[0].currentscript.value="meterstart.cgi"
				dsn="rm_lucy"								
			}
			var temp= "http://www.genergy.com/cgi-bin/" + script + "?bldg=" + bldg + "&lid=" + id +"&dsn=" + dsn
			document.frames.lmp.location.href=temp;
		
	}
function AgrTenant(type, bldg, lid){
			var now=new Date()
			document.forms[0].currentid.value=lid
			document.forms[0].datefield.value=now.toLocaleString()
			if (type=="pt") {
				script="ptagrstart.cgi"
				document.forms[0].currentscript.value="ptagrstart.cgi"
			} else {
				script="agrstart.cgi"
				document.forms[0].currentscript.value="agrstart.cgi"
			}
			dsn="sqlserverg1"
			var temp = "http://www.genergy.com/cgi-bin/" + script + "?bldg=" + bldg + "&lid=" + lid +"&dsn=" + dsn
			document.frames.lmp.location.href=temp;
	}
	
function navigate(direc, datevalue,bldg,lid,profiletype){
	var currdate = new Date(datevalue)
	if (direc == "prev") {
	
		currdate=new Date(currdate).valueOf() - (1 * 86400000)
		}else{
		
		currdate=new Date(currdate).valueOf() + (1 * 86400000)
	    }
	
	currdate = new Date(currdate)
	document.forms[0].datefield.value=currdate.toLocaleString()
	document.forms[0].currentid.value=lid
	var script =document.forms[0].currentscript.value
	currdate = (currdate.getMonth() + 1) + "/" + currdate.getDate() + "/" + currdate.getFullYear()
	
			
	var temp = "http://www.genergy.com/cgi-bin/" + script + "?datefield=" + currdate + "&start=01:00&end=24:00&interval=60&graph=bar3d&bldg=" + bldg + "&lid="+ lid+"&dsn=sqlserverg1"
	document.frames.lmp.location.href=temp;

}
function set_date_field(){
	var now=new Date()
	document.forms[0].datefield.value=now.toLocaleString()
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


<body bgcolor="#FFFFFF" text="#000000" onload="set_date_field()">
<div align="center">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td bgcolor="#0099FF"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="4" color="#000000">Load 
          Management Profiles</font></div>
      </td>
    </tr>
  </table>
  
  <form method="post" action="" name="lmp">
    <div align="left"> 
      <table width="306" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr> 
          <td height="37" width="101"> 
            <div align="left"><font face="Arial, Helvetica, sans-serif" size="3">Tenant</font></div>
          </td>
          <td height="37" width="92"> 
            <div align="left"><font face="Arial, Helvetica, sans-serif" size="3">Meter</font></div>
          </td>
        </tr>
        <tr> 
          <td height="56" width="101"> 
            <%
		Set cnn1 = Server.CreateObject("ADODB.Connection")
		Set rst1 = Server.CreateObject("ADODB.recordset")
		cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
		%>
            <div align="center"> 
              <div align="left"></div>
              <div align="left"> <font face="Arial, Helvetica, sans-serif" size="3"> 
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
                  <option value=<%=rst1("leaseutilityid")%> selected>[<%=rst1("tenantnum") %>] 
                  <%=rst1("billingname") %> </option>
                  <%	
				end if
				
				rst1.close
				
				else
				
				%>
                  <option selected>Select Tenant </option>
                  <%
				
				end if
				
				if leaseid <> "" then
				
				strsql = "SELECT DISTINCT tblLeases.BillingName, tblLeasesUtilityPrices.loadprofile,tblLeasesUtilityPrices.LeaseUtilityId, Meters.Online, tblLeases.TenantNum FROM tblLeases INNER JOIN tblLeasesUtilityPrices ON tblLeases.BillingId = tblLeasesUtilityPrices.BillingId INNER JOIN Meters ON tblLeasesUtilityPrices.LeaseUtilityId = Meters.LeaseUtilityId WHERE (tblLeases.BldgNum = N'" & Request("bldg") & "') AND (Meters.Online = 1) and tblLeasesUtilityPrices.loadprofile=1 and tblLeasesUtilityPrices.LeaseUtilityId <> " & leaseid & " order by tblleases.billingname"
				
				else
				
				strsql = "SELECT DISTINCT tblLeases.BillingName, tblLeasesUtilityPrices.loadprofile,tblLeasesUtilityPrices.LeaseUtilityId, Meters.Online, tblLeases.TenantNum FROM tblLeases INNER JOIN tblLeasesUtilityPrices ON tblLeases.BillingId = tblLeasesUtilityPrices.BillingId INNER JOIN Meters ON tblLeasesUtilityPrices.LeaseUtilityId = Meters.LeaseUtilityId WHERE (tblLeases.BldgNum = N'" & Request("bldg") & "') AND (Meters.Online = 1) and tblLeasesUtilityPrices.loadprofile=1 order by tblleases.billingname"

				end if 
 				
				rst1.Open strsql, cnn1, adOpenStatic		
				
				do until rst1.EOF 
			    %>
                  <option value=<%=rst1("leaseutilityid")%>><%=rst1("tenantnum") %> 
                  - <%=rst1("billingname")%></option>
                  <%
				rst1.movenext
			   	loop
							
				
			%>
                </select>
                </font></div>
            </div>
          </td>
          <td height="56" width="92"><font face="Arial, Helvetica, sans-serif" size="3"> 
            <%
				strsql = "SELECT meterid, g1onlinedate from meters where bldgnum='" & Request("bldg") & "' and pp=1 order by meternum"
		 		rst1.close
				rst1.Open strsql, cnn1, adOpenStatic
			    if not rst1.eof then
				%>
            <input type="hidden" name="bldglmp" value="<%=rst1("meterid")%>_<%=profiletype%>" >
            <% End if 
					%>
            <select name="meter" onChange="loadmeter(this.value,bldg.value)">
              <%rst1.close
				if leaseid <> "" then
				strsql = "SELECT meterid, meternum, g1onlinedate from meters where (LeaseUtilityId=" & leaseid & "and online=1 and g1onlinedate<>'1/1/1900' ) or (leaseUtilityId=" & leaseid & "and online=1 and EXISTS (select * from tblLeasesUtilityPrices where LeaseUtilityId=" & leaseid &  " and LoadProfile=1))order by meternum"

				rst1.Open strsql, cnn1, adOpenStatic
				do until rst1.eof
				if rst1("g1onlinedate") then 
				
			%>
              <option value="<%=rst1("meterid")%>_lm"><font face="Arial, Helvetica, sans-serif" size="3"><%=rst1("meternum") %></font></option>
              <% 
				else
			%>
              <option value="<%=rst1("meterid")%>_pt"><font face="Arial, Helvetica, sans-serif" size="3"><%=rst1("meternum") %></font></option>
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
      <div align="center"></div>
      <table width="300" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><font face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif" size="3">
            <input type="button" name="buildglmp" value="Building LMP" onClick="loadmeter(bldglmp.value, bldg.value)">
            </font></font></font></td>
          <td>
            <input type="button" name="Button2" value="Meter LMP" onClick="loadmeter(meter.value,bldg.value)" >
          </td>
          <td>
            <input type="button" name="Button" value="Aggregate LMP" onClick="AgrTenant(profiletype.value, bldg.value, leaseid.value)" >
          </td>
        </tr>
      </table>
      <table width="402" border="0" align="center">
        <tr valign="middle"> 
          <td>&nbsp;</td>
          <td>&nbsp;</td>
          <td>&nbsp;</td>
        </tr>
        <tr valign="middle"> 
          <td> 
		  <input type="hidden" name="currentid" value="">
  		  <input type="hidden" name="currentscript" value="">
            <input type="button" name="Button32" value="&lt;&lt;&lt;" onClick="navigate(prev.value, datefield.value,bldg.value,currentid.value,profiletype.value)">
          </td>
          <td> 
            <div align="center"> 
              <input type="text" name="datefield"  size="50">
            </div>
          </td>
          <td> 
            <input type="button" name="Button3" value="&gt;&gt;&gt;" onClick="navigate(next.value, datefield.value,bldg.value,currentid.value,profiletype.value)">
          </td>
        </tr>
      </table>
      
    </div>
    </form>
  <p align="left"><IFRAME name="lmp" src="null.htm" width="100%" height="100%" scrolling="auto" marginwidth="8" marginheight="16"></IFRAME></p>
        </div>
</body>
</html>
