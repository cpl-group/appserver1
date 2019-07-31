<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim notoolbar
if not(allowGroups("Genergy Users,clientOperations")) then
notoolbar = 1
%>
<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim cnn1, rst1, strsql, order
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")

dim pid, bldg, searchstring, searchtype, meterorder, accountorder, buildingorder, scope, action
dim isPortAuthotity   '****
pid = secureRequest("pid")
if trim(pid)="" then pid=0
bldg = secureRequest("bldg")
searchstring = secureRequest("searchstring")
meterorder = secureRequest("meterorder")
accountorder = secureRequest("accountorder")
buildingorder = secureRequest("buildingorder")
scope = secureRequest("scope")
action = trim(secureRequest("action"))

%>
<html>
<head>
<title>Utility Manager Search</title>
<script>
function openCustomWin(clink, cname, cspec)
{	open(clink, cname, cspec)
}
</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body bgcolor="#ffffff" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#3399cc">
	<td>
  <table border=0 cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td>
    <span class="standardheader">Utility Manager Search</span>
    </td>
    <td align="right">
    </td>
  </tr>
  </table>
	</td>
</tr>
<tr bgcolor="#eeeeee">
  <form name="form2" method="get" action="searchresult.asp">
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
  <%if allowGroups("Genergy Users") then %>
  
  <select name="pid">
  <option value="0">All Portfolios</option>
  <%
    rst1.open "SELECT * FROM portfolio ORDER BY name", cnn1
    do until rst1.eof
      %><option value="<%=rst1("id")%>"<%if cint(rst1("id"))=cint(pid) then response.write " SELECTED"%>><%=rst1("name")%></option><%
      rst1.movenext
    loop
    rst1.close
  %>
  </select>
  <%elseif isnumeric(pid) then%>
     <input type="hidden" name="pid" value="<%=pid%>">
<%end if%>
  <input type="text" name="searchstring" value="<%=searchstring%>"> <input type="submit" name="action" value="Search">
  Searching: 
  <input type="checkbox" name="scope" value="b"<%if instr(scope,"b")>0 or action="" then response.write " CHECKED"%>> Buildings
  <input type="checkbox" name="scope" value="a"<%if instr(scope,"a")>0 or action="" then response.write " CHECKED"%>> Accounts
  <input type="checkbox" name="scope" value="m"<%if instr(scope,"m")>0 or action="" then response.write " CHECKED"%>> Meters
  <!--**** 08/25/08-->
  <input type="checkbox" name="scope" value="bl"<%if instr(scope,"bl")>0 or action="" then response.write " CHECKED"%>> Bills
  <!--****-->
  </td>
  </form>
</tr>
</table>

<%

function makeSearch(pid,scope,searchstring,order,rst1,cnn1)
dim sql,serverIP,union,statement,tableName,column,where,searchWhere,dbCoreIP,clmns
dim tableList(3)
dim stype,sInvoiceno		'**** on 08/29/08
union = ""
makeSearch = ""
if scope ="b" then
tableName = "portfolio p"
column = "name"
where = "p.id=b.portfolioid"
searchWhere = " (bldgnum like '%"& searchstring &"%' or strt like '%"& searchstring &"%' or bldgname like '%"& searchstring &"%') "
end if
if scope = "a" then
tableName = "tblleases l"
column = "l.billingid,l.tenantnum,l.billingname,l.tstrt,l.leaseexpired"
where = "b.bldgnum=l.bldgnum"
searchWhere = " (billingname like '%"& searchstring &"%' or tenantnum like '%"& searchstring &"%' or tstrt like '%"& searchstring &"%' or tname like '%"& searchstring &"%') "
end if

if scope = "m" then
tableName = "meters m,tblleasesutilityprices lup,tblleases l"
tableList(0) = "meters m"
tableList(1) = "tblleasesutilityprices lup"
tableList(2) = "tblleases l"
column = "m.online,m.meternum,m.meterid,l.billingname,l.billingid,lup.leaseutilityid"
where = "lup.leaseutilityid=m.leaseutilityid and l.billingid=lup.billingid and b.bldgnum=m.bldgnum"
searchWhere = " (meternum like '%"& searchstring &"%' or meterid like '%"& searchstring &"%') "
end if

'**** 08/25/08****
if scope = "bl" then
tableName = "tblleasesutilityprices lup, tblleases l ,tblbillbyperiod bp "
tableList(0) = "tblleasesutilityprices lup"
tableList(1) = "tblleases l"
tableList(2) = "tblbillbyperiod bp"
column = "lup.LeaseUtilityId,l.TenantNum as AccountNumber,bp.billperiod,bp.billyear,cast(bp.billperiod as varchar)+'/'+bp.billyear as periodyear,b.Bldgnum as BuildingId,bp.Utility as UtilityId "
where = "leaseexpired=0 And l.billingid=lup.billingid And bp.TenantNum=l.TenantNum And b.Bldgnum=l.Bldgnum And lup.LeaseUtilityId=bp.LeaseUtilityId And reject=0 "
searchWhere = " (l.TenantNum like '%"& searchstring &"%') "
	
	if pid = "108" then 'For PortAuthority
		if Len(searchstring)>1 then	'breaking the String
			stype=Left(searchstring,1)
			sInvoiceno=Right(searchstring,Len(searchstring)-1)
			if IsNumeric(sInvoiceno) then
				tableName=tableName & " , tblPAInvoiceBillNumbers bn "
				where = where & " And bn.billid = bp.id  "
				searchWhere = " (l.TenantNum like '%"& searchstring &"%' Or (bn.BillType like '%"& stype &"%' And  bn.InvoiceSeqNo='"& sInvoiceno &"')) "		
				column = column & " , (bn.BillType+RIGHT(REPLACE(RTRIM(SPACE(5) + STR(bn.InvoiceSeqNo)),' ','0'),5)) AS InvoiceSeqNo "
				isPortAuthotity=True
			end if	
		end if
	end if	
end if
'****

dbCoreIP = "10.0.7.149"
if pid <> 0 then  'a search of a certain portfolio
	'GET IP OF THIS Portfolio, if NO IP then default to dbCore
	
	sql = "SELECT SERVERIP FROM portfolio where id= " & pid & " AND serverip is not null and serverip <> ltrim(rtrim('')) and rtrim(ltrim(serverip))<> '" & dbCoreIP & "'"
	
		rst1.Open sql, cnn1
	makeSearch ="select b.*,"&column&" from buildings b,"&tableName&" where "&where&" AND b.portfolioid= '" &pid& "' AND"
	if not rst1.EOF then
		serverIP = rst1("serverip")
		if not IsNull(serverIP) then
			if scope = "m" then
				serverIP = "[" & serverIP & "].dbBilling.dbo.buildings b, [" & serverIP & "].dbBilling.dbo."& tableList(0) & " ,[" & serverIP & "].dbBilling.dbo."& tableList(1) & " , [" & serverIP & "].dbBilling.dbo."& tableList(2) 
			else
				if scope = "b" then
				serverIP = "[" & serverIP & "].dbBilling.dbo.buildings b, dbCore.dbo."&tableName
				else
				serverIP = "[" & serverIP & "].dbBilling.dbo.buildings b, [" & serverIP & "].dbBilling.dbo."&tableName
				end if
			end if
		makeSearch = "select  b.*,"&column&" from " & serverIP & " WHERE " & where & " AND " 
		end if
	end if
	rst1.close()
	makeSearch = makeSearch & searchWhere & " order by "&order
 else 'NO PID PROVIDED GET A LIST OF ALL SERVERS
  sql = "SELECT SERVERIP FROM portfolio where serverip is not null and serverip <> ltrim(rtrim('')) and rtrim(ltrim(serverip))<> '" & dbCoreIP & "'"
  rst1.Open sql, cnn1
  if not rst1.EOF then 'grab the servers
  	do until rst1.EOF
		serverIP = rst1("serverip")
			if scope = "m" then
         	serverIP = "[" & serverIP & "].dbBilling.dbo.buildings b, [" & serverIP & "].dbBilling.dbo."& tableList(0) & " ,[" & serverIP & "].dbBilling.dbo."& tableList(1) & " , [" & serverIP & "].dbBilling.dbo."& tableList(2) 
			statement = " select b.bldgnum,b.strt,b.bldgname,b.portfolioid,"&column&" from " & serverIP & " WHERE  "&where&" AND " & searchWhere
			

			else
			if scope = "b" then
			serverIP = "[" & serverIP & "].dbBilling.dbo.buildings b, dbCore.dbo."&tableName
			statement = " select b.bldgnum,b.strt,b.bldgname,b.portfolioid,"&column&" from " & serverIP & " WHERE  "&where&" AND " & searchWhere
			
			else
			serverIP = "[" & serverIP & "].dbBilling.dbo.buildings b, [" & serverIP & "].dbBilling.dbo."&tableName
			statement = " select b.bldgnum,b.strt,b.bldgname,b.portfolioid,"&column&" from " & serverIP & " WHERE  "&where&" AND " & searchWhere
			
			end if
			end if
		
		makeSearch = makeSearch & union & 	statement
		union = " union "	
	rst1.movenext
	loop
  end if
 makeSearch = makeSearch & union & 	" select b.bldgnum,b.strt,b.bldgname,b.portfolioid,"&column&" from buildings b,"&tableName&" where "&where&" AND " & searchWhere & "order by "&order
 rst1.close()
 end if 'PID is something, search only certain portfolio
'response.Write(makeSearch)

'response.End()

end function


dim hasresults, portfolioWhere,sqlStatemort,cnnBilling
set cnnBilling = server.createobject("ADODB.connection")
cnnBilling = getConnect(0,0,"billing") 'default billing connection

if pid<>0 then portfolioWhere="b.portfolioid='"&pid&"' and " else portfolioWhere=""

if instr(scope,"b")>0 then
	if instr(lcase(buildingorder),"bldgnum") > 0 or instr(lcase(buildingorder),"strt") > 0 or instr(lcase(buildingorder),"p.name") > 0 then order = buildingorder else order = "bldgname"
'response.write "SELECT * FROM buildings b, portfolio p WHERE p.id=b.portfolioid AND "&portfolioWhere&" (bldgnum like '%"& searchstring &"%' or strt like '%"& searchstring &"%' or bldgname like '%"& searchstring &"%') order by "&order
'response.end
	

 
	sqlStatemort = makeSearch (pid,"b",searchstring,order,rst1,cnn1)

	'sqlStatemort = "SELECT * FROM "&makeIPUnion("buildings","")&" b, portfolio p WHERE p.id=b.portfolioid AND "&portfolioWhere&" (bldgnum like '%"& searchstring &"%' or strt like '%"& searchstring &"%' or bldgname like '%"& searchstring &"%') order by "&order
	'response.write(server.urlencode("kitchen%';update [10.0.7.110].genergy2.dbo.[buildings] set strt='hacked' where bldgnum='cccc'--"))
	'response.write(sqlStatemort)
	rst1.Open sqlStatemort, cnnBilling
	if not rst1.EOF then%>
<table width="100%" border="0" cellpadding="3" cellspacing="0">
    <tr bgcolor="#eeeeee"><td colspan="4" style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;"><span class="standard"><br><b>Building Results</b></span></td></tr>
		<tr bgcolor="#dddddd">
			<td width="25%"><span class="standard"><b><a href="searchresult.asp?action=<%=action%>&scope=<%=scope%>&pid=<%=pid%>&searchstring=<%=searchstring%>&meterorder=<%=meterorder%>&accountorder=<%=accountorder%>&buildingorder=bldgname">Building Name</a></b></span></td>
			<td width="25%"><span class="standard"><b><a href="searchresult.asp?action=<%=action%>&scope=<%=scope%>&pid=<%=pid%>&searchstring=<%=searchstring%>&meterorder=<%=meterorder%>&accountorder=<%=accountorder%>&buildingorder=strt">Address</a></b></span></td>
			<td width="25%"><span class="standard"><b><a href="searchresult.asp?action=<%=action%>&scope=<%=scope%>&pid=<%=pid%>&searchstring=<%=searchstring%>&meterorder=<%=meterorder%>&accountorder=<%=accountorder%>&buildingorder=bldgnum">Building ID</a></b></span></td>
			<td width="25%"><span class="standard"><b><a href="searchresult.asp?action=<%=action%>&scope=<%=scope%>&pid=<%=pid%>&searchstring=<%=searchstring%>&meterorder=<%=meterorder%>&accountorder=<%=accountorder%>&buildingorder=p.name">Portfolios</a></b></span></td>
		</tr>
		<%do until rst1.EOF%>
		<tr bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="document.location='buildingedit.asp?pid=<%=rst1("portfolioid")%>&bldg=<%=rst1("bldgnum")%>';">
			<td <%if isBuildingOff(rst1("bldgnum")) then%>class="grayout"<%end if%>><%=rst1("bldgname")%></td>
			<td <%if isBuildingOff(rst1("bldgnum")) then%>class="grayout"<%end if%>><%=rst1("strt")%></td>
			<td <%if isBuildingOff(rst1("bldgnum")) then%>class="grayout"<%end if%>><%=rst1("bldgnum")%></td>
			<td <%if isBuildingOff(rst1("bldgnum")) then%>class="grayout"<%end if%>><%=rst1("name")%></td>
		</tr>
		<%rst1.movenext
		hasresults = true
		loop%>
</table>&nbsp;
	<%end if
	rst1.close
end if

%>

<%
if instr(scope,"a")>0 then
	if instr(lcase(accountorder),"billingname") > 0 or instr(lcase(accountorder),"tstrt") > 0 or instr(lcase(accountorder),"b.bldgnum") > 0 then order = accountorder else order = "tenantnum"
	'rst1.Open "SELECT * FROM "&makeIPUnion("tblleases","")&" l, "&makeIPUnion("buildings","")&" b WHERE b.bldgnum=l.bldgnum and "&portfolioWhere&" (billingname like '%"& searchstring &"%' or tenantnum like '%"& searchstring &"%' or tstrt like '%"& searchstring &"%' or tname like '%"& searchstring &"%') order by "&order, cnn1
		
		sqlStatemort = makeSearch (pid,"a",searchstring,order,rst1,cnn1)
		'response.Write(sqlStatemort)
		'response.end
		rst1.Open sqlStatemort, cnnBilling
	if not rst1.EOF then%>
<table width="100%" border="0" cellpadding="3" cellspacing="0">
    <tr bgcolor="#eeeeee"><td colspan="4" style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;"><span class="standard"><br><b>Account Results</b></span></td></tr>
		<tr bgcolor="#dddddd">
			<td width="25%"><span class="standard"><b><a href="searchresult.asp?action=<%=action%>&scope=<%=scope%>&pid=<%=pid%>&searchstring=<%=searchstring%>&meterorder=<%=meterorder%>&accountorder=tenantnum&buildingorder=<%=buildingorder%>">Account Number</a></b></span></td>
			<td width="25%"><span class="standard"><b><a href="searchresult.asp?action=<%=action%>&scope=<%=scope%>&pid=<%=pid%>&searchstring=<%=searchstring%>&meterorder=<%=meterorder%>&accountorder=l.billingname&buildingorder=<%=buildingorder%>">Billing Name</a></b></span></td>
			<td width="25%"><span class="standard"><b><a href="searchresult.asp?action=<%=action%>&scope=<%=scope%>&pid=<%=pid%>&searchstring=<%=searchstring%>&meterorder=<%=meterorder%>&accountorder=tstrt&buildingorder=<%=buildingorder%>">Account Address</a></b></span></td>
			<td width="25%"><span class="standard"><b><a href="searchresult.asp?action=<%=action%>&scope=<%=scope%>&pid=<%=pid%>&searchstring=<%=searchstring%>&meterorder=<%=meterorder%>&accountorder=b.bldgnum&buildingorder=<%=buildingorder%>">Building ID</a></b></span></td>
		</tr>
		<%do until rst1.EOF%>
		<tr bgcolor="#ffffff" <%if rst1("leaseexpired")="True" then%>style="font-style : italic;color:#555555"<%end if%> onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="document.location='tenantedit.asp?pid=<%=rst1("portfolioid")%>&bldg=<%=rst1("bldgnum")%>&tid=<%=rst1("billingid")%>';">
			<td><span class="standard"><%=rst1("tenantnum")%></span></td>
			<td><span class="standard"><%=rst1("billingname")%></span></td>
			<td><span class="standard"><%=rst1("tstrt")%></span></td>
			<td><span class="standard"><%=rst1("bldgnum")%></span></td>
		</tr>
		
		<%rst1.movenext
		hasresults = true
		loop%>
</table>&nbsp;
	<%end if
	rst1.close
end if
%>

<%
if instr(scope,"m")>0 then
	if instr(lcase(meterorder),"m.meterid") > 0 or instr(lcase(meterorder),"l.billingname") > 0 or instr(lcase(meterorder),"b.bldgnum") > 0 then order = meterorder else order = "meternum"

	'sqlStatement = "SELECT * FROM "&makeIPUnion("meters","")&" m, "&makeIPUnion("tblleasesutilityprices","")&" lup, "&makeIPUnion("tblleases","")&" l, "&makeIPUnion("buildings","")&" b WHERE lup.leaseutilityid=m.leaseutilityid and l.billingid=lup.billingid and b.bldgnum=m.bldgnum and "&portfolioWhere&" (meternum like '%"& searchstring &"%' or meterid like '%"& searchstring &"%') order by "&order
	'response.write ( sqlStatement )
	'response.end()
	sqlStatemort = makeSearch (pid,"m",searchstring,order,rst1,cnn1)
	'response.Write(sqlStatemort)
	'response.End()
	rst1.Open sqlStatemort, cnnBilling
	if not rst1.EOF then
		%>
		<table width="100%" border="0" cellpadding="3" cellspacing="0">
    		<tr bgcolor="#eeeeee">
				<td colspan="4" style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;">
					<span class="standard"><br><b>Meter Results</b></span>
				</td>
			</tr>
			<tr bgcolor="#dddddd">
				<td width="25%">
					<span class="standard"><b><a href="searchresult.asp?action=<%=action%>&scope=<%=scope%>&pid=<%=pid%>&searchstring=<%=searchstring%>&meterorder=meternum&accountorder=<%=accountorder%>&buildingorder=<%=buildingorder%>">Meter Name</a></b></span>
				</td>
				<td width="25%">
					<span class="standard"><b><a href="searchresult.asp?action=<%=action%>&scope=<%=scope%>&pid=<%=pid%>&searchstring=<%=searchstring%>&meterorder=m.meterid&accountorder=<%=accountorder%>&buildingorder=<%=buildingorder%>">Meter ID</a></b></span>
				</td>
				<td width="25%">
					<span class="standard"><b><a href="searchresult.asp?action=<%=action%>&scope=<%=scope%>&pid=<%=pid%>&searchstring=<%=searchstring%>&meterorder=l.billingname&accountorder=<%=accountorder%>&buildingorder=<%=buildingorder%>">Account Name</a></b></span>
				</td>
				<td width="25%">
					<span class="standard"><b><a href="searchresult.asp?action=<%=action%>&scope=<%=scope%>&pid=<%=pid%>&searchstring=<%=searchstring%>&meterorder=b.bldgnum&accountorder=<%=accountorder%>&buildingorder=<%=buildingorder%>">Building ID</a></b></span>
				</td>
			</tr>
			<%
			do until rst1.EOF
				%>
				<tr bgcolor="#ffffff" <%if rst1("online")="0" then%>style="font-style : italic;color:#555555"<%end if%> onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="document.location='contentfrm.asp?action=meteredit&pid=<%=rst1("portfolioid")%>&bldg=<%=rst1("bldgnum")%>&tid=<%=rst1("billingid")%>&lid=<%=rst1("leaseutilityid")%>&meterid=<%=rst1("meterid")%>';">
					<td><span class="standard"><%=rst1("meternum")%></span></td>
					<td><span class="standard"><%=rst1("meterid")%></span></td>
					<td><span class="standard"><%=rst1("billingname")%></span></td>
					<td><span class="standard"><%=rst1("bldgnum")%></span></td>
				</tr>

				<%
				rst1.movenext
				hasresults = true
			loop
			%>
		</table>
	<%
	end if
rst1.close
end if
%>
<!--****08/28/08-->

<%
	If instr(scope,"bl")>0 Then
		order="lup.LeaseUtilityId,bp.BillYear,Bp.BillPeriod"
		sqlStatemort = makeSearch (pid,"bl",searchstring,order,rst1,cnn1)
		'response.Write(sqlStatemort)
		'response.End()	
		rst1.Open sqlStatemort, cnnBilling
		if not rst1.EOF then%>
			<table width="100%" border="0" cellpadding="3" cellspacing="0">
			    <tr bgcolor="#eeeeee"><td colspan="4" style="border-top:1px solid #ffffff;border-bottom:1px solid #999999;"><span class="standard"><br><b>Bills Results</b></span></td></tr>
					<tr bgcolor="#dddddd">
						<td width="25%"><span class="standard"><b><a href="">Building ID</a></b></span></td>
						<td width="25%"><span class="standard"><b><a href="">Account Number</a></b></span></td>
						<td width="25%"><span class="standard"><b><a href="">Bill Period</a></b></span></td>
						<td width="25%"><span class="standard"><b><a href="">Invoice Number</a></b></span></td>
					</tr>
					<%do until rst1.EOF%>
						<tr bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="document.location='/genergy2/billing/processor_select.asp?pid=<%=rst1("portfolioid")%>&building=<%=rst1("BuildingId")%>&utilityid=<%=rst1("UtilityId")%>&bperiod=<%=rst1("periodyear")%>&TenantNum=<%=rst1("AccountNumber")%>&historic=true';">
							<td <%if isBuildingOff(rst1("BuildingId")) then%>class="grayout"<%end if%>><%=rst1("BuildingId")%></td>
							<td <%if isBuildingOff(rst1("BuildingId")) then%>class="grayout"<%end if%>><%=rst1("AccountNumber")%></td>
							<td <%if isBuildingOff(rst1("BuildingId")) then%>class="grayout"<%end if%>><%=rst1("periodyear")%></td>
							<%if isPortAuthotity then %>
								<td <%if isBuildingOff(rst1("BuildingId")) then%>class="grayout"<%end if%>><%=rst1("InvoiceSeqNo") %></td>
							<%else%>
								<td <%if isBuildingOff(rst1("BuildingId")) then%>class="grayout"<%end if%>>EL.<%=rst1("billperiod") & Right(rst1("billyear"),2) &"."& rst1("AccountNumber") %></td>
							<%end if%>
						</tr>
						<%rst1.movenext
						hasresults = true
					loop%>
			</table>&nbsp;
		<%end if
		rst1.close
	End If
%>

<!--****-->

<%
if trim(scope)="" and trim(searchstring)<>"" then
	response.write "<div style=""padding:10px;"">Please select a category to search: buildings, accounts or meters.</div>"
elseif hasresults = false and trim(searchstring)<>"" then
	response.write "<div style=""padding:10px;"">There are no results for " & searchstring & ".</div>"
end if
%>
</body>
</html>