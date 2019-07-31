<%
function getUMInfo(ticketfortype, ticketfor, bldg)
	dim rsUMInfo, sqlUMInfo
	set rsUMInfo = server.createobject("ADODB.recordset")
	select case UCASE(ticketfortype)
	case "PID"
		rsUMInfo.ActiveConnection = getMainConnect(ticketfor)
		sqlUMInfo = "SELECT 'Opened for portfolio '+name+'. ' as info "&_
					"FROM portfolio WHERE id='"&ticketfor&"'"
	case "BLDGNUM"
		rsUMInfo.ActiveConnection = getLocalConnect(ticketfor)
		sqlUMInfo = "SELECT 'Opened for building '+bldgname+', portfolio '+p.name+'. ' as info "&_
					"FROM buildings b, ["&application("CoreIP")&"].dbCore.dbo.portfolio p  WHERE b.portfolioid=p.id and b.bldgnum='"&ticketfor&"'"
	case "TID"
		if request("bldg")<>"" then
			rsUMInfo.ActiveConnection = getLocalConnect(bldg)
			sqlUMInfo = "SELECT 'Opened for account '+BillingName+' ('+tenantnum+'), building '+bldgname+'('+convert(varchar(7),b.bldgnum)+'), portfolio '+p.name+'. ' as info "&_
						"FROM tblleases l, Buildings b, ["&application("coreIP")&"].dbCore.dbo.portfolio p WHERE '"&split(getBuildingIP(bldgnum),"\")(1)&"-' + ltrim(convert(varchar, billingid))='"&ticketfor&"' and b.portfolioid=p.id and b.bldgnum=l.bldgnum"
		end if
	case "METERID"
		if request("bldg")<>"" then
			rsUMInfo.ActiveConnection = getLocalConnect(bldg)
			sqlUMInfo = "SELECT 'Opened for meter '+meternum+' ('+convert(varchar(10),meterid)+'), account '+BillingName+' ('+tenantnum+'), building '+bldgname+', portfolio '+p.name+'. ' as info "&_
						"FROM meters m, Buildings b, tblleasesutilityprices lup, tblleases l, ["&application("CoreIP")&"].dbCore.dbo.portfolio p WHERE '"&split(getBuildingIP(bldgnum),"\")(1)&"-' + ltrim(convert(varchar, meterid))='"&ticketfor&"' and b.portfolioid=p.id and m.bldgnum=b.bldgnum and lup.leaseutilityid=m.leaseutilityid and l.billingid=lup.billingid"
		end if
	end select
	if sqlUMInfo <> "" then
		rsUMInfo.open sqlUMInfo
		if not rsUMInfo.eof then
			getUMInfo = rsUMInfo("info")
		end if
		rsUMInfo.close
	end if
end function
%>