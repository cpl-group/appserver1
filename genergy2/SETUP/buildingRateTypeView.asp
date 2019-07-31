<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
	dim pid, bldg, utility, TenantRate, procname, str, strmsg, strQuery
	'pid = secureRequest("pid")
	'bldg = secureRequest("bldg")
	strmsg=""
	pid = request.QueryString("pid")
	bldg = request.QueryString("bldg")
	
	if pid="" then
		pid = request("PID")
	End if
	if bldg="" then
		bldg = request("BuildingNumber")
	End if
	procname = request("procname")
	TenantRate = request("TenantRate")
	
	if secureRequest("Utility")="" then
		utility=request("Utility")
	else
		utility=secureRequest("Utility")
	End if
	
	if utility="" Then
		utility=2
	End If
	
	dim cnn1, rst1, strsql, rst2
	set cnn1 = server.createobject("ADODB.connection")
	set rst1 = server.createobject("ADODB.recordset")
	cnn1.open getLocalConnect(bldg)

	dim bldgname, portfolioname
	if trim(bldg)<>"" then
		rst1.Open "SELECT bldgname, name, strt FROM buildings b join portfolio p on b.portfolioid=p.id WHERE bldgnum='"&bldg&"'", cnn1
		if not rst1.EOF then
			bldgname = rst1("strt")
			portfolioname = rst1("name")
		end if
		rst1.close
	end if
	
	if Request("action")="Save" then
		dim rst,rstcheck,strCheck,rst4
		set rst	= server.createobject("ADODB.Recordset")
		set rst4 = server.createobject("ADODB.Recordset")
		set rstcheck = server.createobject("ADODB.Recordset")
		'Checking if the rate type exists
		strCheck="Select procname, RateTenant From BuildingRateTypes Where bldgNum ='" & bldg & "' AND UtilityId=" & utility	
		rstcheck.Open strCheck, getConnect(pid,bldg,"billing")
		if rstcheck.EOF then	'if not then
			'Add the rate type 
			strQuery="INSERT INTO [BuildingRateTypes] ([bldgNum], [utilityId], [procname], [rateTenant]) " &_
					"VALUES ('" & bldg &"', "& utility &", '"& procname &"', '"& TenantRate &"')"
		
			rst.open strQuery, getConnect(pid,bldg,"billing")
			set rst=nothing
		
			strmsg="Rate Type Added!"
		else
		    'Insert record in BuildingRateTypesHistory
					
			strQuery="INSERT INTO [BuildingRateTypesHistory] ([bldgNum], [utilityId], [procname], [rateTenant], [DateReplaced]) " &_
					"Select bldgNum, utilityId, procname, RateTenant, getdate() From BuildingRateTypes Where bldgNum ='" & bldg & "' AND UtilityId=" & utility
					
			rst4.open strQuery, getConnect(pid,bldg,"billing")
			set rst4=nothing
			
			'update the rate type first
			strQuery="UPDATE [BuildingRateTypes] set [procname]='" & procname &"', [rateTenant]='" & TenantRate &_
						"' WHERE bldgNum ='" & bldg & "' AND UtilityId=" & utility
			rst.open strQuery, getConnect(pid,bldg,"billing")
			set rst=nothing
			'UPDATE rate type in tblLeasesUtilityPrices
			
			dim rstds,strds
			Set rstds=Server.CreateObject("ADODB.Recordset")
			strds="UPDATE tblLeasesUtilityPrices set RateTenant='" & TenantRate &"', ProcName='" & procname &_
						"' WHERE billingId in (select billingId From tblLeases Where BldgNum='" & bldg &"') " &_
						"AND utility=" & utility &" AND UseBldgRate=1"
			rstds.open strds, getConnect(pid,bldg,"billing")
			set rstds=nothing
			
					
			strmsg="Rate Type Updated!"
			'Response.Write(strmsg)	
		End if	
		rstcheck.Close 
		
	End if
%>
<html>
<head>
	<link rel="Stylesheet" href="/genergy2/SETUP/setup.css" type="text/css">
	<title>Add/Edit Building Level Rate Type</title>
	<script>
		function closeWinda(){ 
		window.close()
		}
		
		function checkform(frm){
			var err = "";
			if(frm.procname.value=='') err+="Select Proc name\n";
			if(frm.utility.value=='') err+="Select Utility name\n"
			if(frm.TenantRate.value=='') err+="Select Tenant Rate\n"
			if(err=="") 
				return true 
			else alert(err);
				return false;
		}
	</script>
</head>

<body bgcolor="#eeeeee" topmargin=0 leftmargin=0 marginwidth="0" marginheight="0">
	<form action="buildingRateTypeView.asp">
		<table width="100%" border="0" cellpadding="3" cellspacing="0" align="center">
			<tr bgcolor="#3399cc">
				<td><font color='white'>Add/Edit Building Level Rate Type for  </font>&nbsp;<b>Building # <%=bldg%>(<%=bldgname%>)</b></td>
			</tr>
		</table>
		 &nbsp;
		<div id="Datasourceinfo" style="BORDER-RIGHT: #cccccc 1px solid; PADDING-RIGHT: 3px; BORDER-TOP: #cccccc 1px solid; DISPLAY: inline; PADDING-LEFT: 3px; PADDING-BOTTOM: 3px; BORDER-LEFT: #cccccc 1px solid; WIDTH: 98%; PADDING-TOP: 3px; BORDER-BOTTOM: #cccccc 1px solid"> 
		  <table width="100%" border="0" cellpadding="3" cellspacing="0">
		    <tr>
            <td align="left" colspan=3>&nbsp; <b>Add building level rate type: </b></td>
        </tr>
        <tr bgcolor="#eeeeee"> 
            <td align="right"><span class="standard">Utility Type</span></td>
            <td colspan=2><select name="utility" onchange="return utilityChanged(this.value, utility);">
          <%
			rst1.open "SELECT * FROM tblutility ORDER BY utilitydisplay", getConnect(pid,bldg,"dbCore")
			do until rst1.eof
				%>
          <option value="<%=rst1("utilityid")%>"<%if trim(utility)=trim(rst1("utilityid")) then response.write " SELECTED"%>><%=rst1("utilitydisplay")%></option>
          <%
				rst1.movenext
			loop
			rst1.close
			%>
        </select> </td>
		</tr>
		<tr>
        <td align="right"><span class="standard">Account Rate</span></td>
        <td colspan=2><select name="TenantRate">
          <%
			rst1.open "SELECT * FROM ratetypes WHERE regionid in (SELECT region FROM buildings WHERE bldgnum='"& bldg &"') ORDER BY type", getConnect(pid,bldg,"billing")
			do until rst1.eof
				%>
          <option value="<%=rst1("id")%>"<%if trim(tenantrate)=trim(rst1("id")) then response.write " SELECTED"%>><%=rst1("type")%></option>
          <%
				rst1.movenext
			loop
			rst1.close
			%>
        </select> </td>
		</tr>
       <tr>
        <td align="right"><span class="standard">Rate Function</span></td>
        <td colspan=2><select name="procname">
          <%
			rst1.open "SELECT * FROM functiontypes ORDER BY description", getConnect(pid,bldg,"dbCore")
			do until rst1.eof
				%>
          <option value="<%=rst1("id")%>"<%if trim(procname)=trim(rst1("id")) then response.write " SELECTED"%>><%=rst1("description")%></option>
          <%
				rst1.movenext
			loop
			rst1.close
			%>
        </select> </td>
        </tr>
        <tr>
            <td colspan=2></td>
			<td>
				<input name="action" style="BORDER-RIGHT: #ffffff 1px outset; BORDER-TOP: #ffffff 1px outset; BORDER-LEFT: #ffffff 1px outset; CURSOR: hand; COLOR: #336699; BORDER-BOTTOM: #ffffff 1px outset; BACKGROUND-COLOR: #eeeeee" type="submit" value="Save" >&nbsp;
					<input name="close"  style="BORDER-RIGHT: #ffffff 1px outset; BORDER-TOP: #ffffff 1px outset; BORDER-LEFT: #ffffff 1px outset; CURSOR: hand; COLOR: #336699; BORDER-BOTTOM: #ffffff 1px outset; BACKGROUND-COLOR: #eeeeee" type="button" value="Cancel" onclick="javascript:closeWinda();">
			</td>
        </tr>
		  </table>
		</div>
		 &nbsp;
		<TABLE width="100%" border="0" cellpadding="3" cellspacing="0" align="center">  
		  <TR>
		    <TD align=absmiddle bgColor=#3399cc><font color='white'>Current Rate Type Set</font></TD>
		  </TR>
		  <tr><td>
		  <table>
		  <tr bgcolor="#3399cc">
					<td><span class="standard">utilitydisplay</span></td>
					<td><span class="standard">AccountRate</span></td>
					<td><span class="standard">RateFunction</span></td>
			</tr>	
		    <% 
		    str="" 
		    Set rst2 = Server.CreateObject("ADODB.recordset")
				str="select utilitydisplay, type as AccountRate, description as RateFunction " &_
					"From BuildingRateTypes b Inner Join tblutility u on b.utilityId=u.utilityid " &_
					"Inner Join ratetypes r on b.RateTenant=r.id " &_
					"Inner Join functiontypes f on b.ProcName=f.id Where bldgnum ='"& bldg &"'"
				rst2.Open str, getConnect(0,0,"billing"), 0, 1, 1	
				do until rst2.eof%>
				<tr bgcolor="#cccccc">
					<td><span class="standard"><%=rst2("utilitydisplay")%></span></td>
					<td><span class="standard"><%=rst2("AccountRate")%></span></td>
					<td><span class="standard"><%=rst2("RateFunction")%></span></td>
				</tr>	
				<%
				rst2.movenext
				loop
			rst2.close
			%>
			</table></td></tr>
		</TABLE>
		<TABLE width="100%" border="0" cellpadding="3" cellspacing="0" align="center">  
		  <TR>
		    <TD align=absmiddle bgColor=#3399cc><font color='white'>Rate Type Set History</font></TD>
		  </TR>
		  <tr><td>
		  <table>
		  <tr bgcolor="#3399cc">
					<td><span class="standard">utilitydisplay</span></td>
					<td><span class="standard">AccountRate</span></td>
					<td><span class="standard">RateFunction</span></td>
					<td><span class="standard">DateReplaced</span></td>
			</tr>	
		    <% 
		    str="" 
		    Set rst2 = Server.CreateObject("ADODB.recordset")
				str="select utilitydisplay, type as AccountRate, description as RateFunction, DateReplaced " &_
					"From BuildingRateTypesHistory b Inner Join tblutility u on b.utilityId=u.utilityid " &_
					"Inner Join ratetypes r on b.RateTenant=r.id " &_
					"Inner Join functiontypes f on b.ProcName=f.id Where bldgnum ='"& bldg &"'"
				rst2.Open str, getConnect(0,0,"billing"), 0, 1, 1	
				do until rst2.eof%>
				<tr bgcolor="#cccccc">
					<td><span class="standard"><%=rst2("utilitydisplay")%></span></td>
					<td><span class="standard"><%=rst2("AccountRate")%></span></td>
					<td><span class="standard"><%=rst2("RateFunction")%></span></td>
					<td><span class="standard"><%=rst2("DateReplaced")%></span></td>
				</tr>	
				<%
				rst2.movenext
				loop
			rst2.close
			%>
			</table></td></tr>
		</TABLE>
		 <p>
		 <input type="hidden" name="PID" value="<%=pid%>">
		  
		  <input type="hidden" name="BuildingNumber" value="<%=bldg%>">
		  
		</p>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<p>&nbsp; </p>
		<p>&nbsp;</p>
		<p>&nbsp; </p>
	</form>
	<TABLE cellSpacing=0 cellPadding=0 width="100%" align=center  border=0 style="FONT-SIZE: smaller; BOTTOM: 0px; TEXT-ALIGN: center">
		<TR>
			<TD align=absmiddle bgColor=#3399cc><FONT color=white><%=strmsg%></FONT></TD>
		</TR>
	</TABLE>
</body>