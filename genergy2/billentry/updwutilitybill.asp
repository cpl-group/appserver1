<%option explicit
dim numDecimalPlaces
numDecimalPlaces = 5
%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldg, acctid, ypid, id, watercharge, sewer, totalamt, totalccf, avgcost, bp, by, action, avgdailyusage,utility
dim usageCharge, eeCharge,auCharge, opCharge, capCharge, tempCharge,  adjustments, totalTons, totalTonHrs, onPeakTonHrs, offPeakTonHrs, penalty, totalBillAmt, salesTax, miscCharge, miscChargeDesc, note
bldg = request("bldg")
action=Request.form("action")
acctid=Request.form("acctid")
ypid=Request.form("ypid")
id=Request.form("id")
utility = request("utility")
bp=Request.form("bp")
by=Request.form("by")
note = replace(left(request("note"), 500),"'","''")

dim cnn1, strsql
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getLocalConnect(bldg)

if utility = "4" or utility="3" then
	watercharge=Request.form("watercharge")
	sewer=Request.form("sewer")
	totalamt=Request.Form("totalamt")
	totalccf=Request.form("totalccf")
	avgcost=Request.Form("avgcost")

	avgdailyusage = Request("avgdailyusage")
	if action="UPDATE" then
		strsql = "update utilitybill_coldwater set watercharge='" & watercharge & "', totalbillamt='" & totalamt & "',sewercharge='" & sewer & "',totalccf='" & totalccf & "',avgcost='" & avgcost & "', avgdailyusage='"&avgdailyusage&"', note='"&note&"'  where id='" & id &"'"
	elseif action="SAVE" then
		strsql = "insert into utilitybill_coldwater (Ypid, acctid, watercharge, totalbillamt, sewercharge, totalccf, avgcost, avgdailyusage, note) values ("&ypid&", '"&acctid&"', '"& watercharge &"', '"& totalamt &"', '"& sewer &"', '"& totalccf &"', '"& avgcost &"', '"&avgdailyusage&"', '"&note&"')"
	end if
	
	
elseif utility = "6" then

	' some of these values may be nulls, so we have to check to see if they are empty strings.  if they are we will get them ready for insertion into the sql statement

	if trim(request("eeCharge")) 	 <> "" then eeCharge 		= "'" & request("eeCharge")		& "'" else eeCharge		= "0" end if
	if trim(request("opCharge")) 	 <> "" then opCharge 		= "'" & request("opCharge")		& "'" else opCharge 	= "0" end if
	if trim(request("tempCharge"))	 <> "" then tempCharge 		= "'" & request("tempCharge")	& "'" else tempCharge 	= "0" end if
	if trim(request("adjustments"))  <> "" then adjustments 	= "'" & request("adjustments")	& "'" else adjustments 	= "0" end if
	if trim(request("onPeakTonHrs")) <> "" then onPeakTonHrs 	= "'" & request("onPeakTonHrs")	& "'" else onPeakTonHrs = "0" end if
	if trim(request("offPeakTonHrs"))<> "" then offPeakTonHrs	= "'" & request("offPeakTonHrs")& "'" else offPeakTonHrs= "0" end if
	if trim(request("penalty")) 	 <> "" then penalty 		= "'" & request("penalty")		& "'" else penalty 		= "0" end if
	if trim(request("totalBillAmt")) <> "" then totalBillAmt	= "'" & request("totalBillAmt")	& "'" else totalBillAmt	= "0" end if
	if trim(request("taxPercent"))	 <> "" then salesTax 		= "'" & request("taxPercent")	& "'" else salesTax 	= "0" end if
	if trim(request("miscChargeDesc"))<> "" then miscChargeDesc  = "'" & request("MiscChargeDesc")&"'" else miscChargeDesc="0" end if
	if trim(request("miscCharge"))	 <> "" then miscCharge		= "'" & request("miscCharge")	& "'" else miscCharge	= "0" end if
	' these ones get the quotes after we do math to get per unit costs
	if trim(request("usageCharge"))  <> "" then usageCharge	= request("usageCharge")	else usageCharge 	= "0" end if
	if trim(request("capCharge"))	 <> "" then capCharge 	= request("capCharge")		else capCharge 		= "0" end if
'	if trim(request("auCharge")) 	 <> "" then auCharge 	= request("auCharge")		else auCharge 		= "0" end if
	if trim(request("totalTons")) 	 <> "" then totalTons	= request("totalTons")		else totalTons		= "0" end if
	if trim(request("totalTonHrs"))	 <> "" then totalTonHrs	= request("totalTonHrs")	else totalTonHrs 	= "0" end if
	if trim(request("unitCostTonh"))	 <> "" then unitCostTonh	= request("unitCostTonh")	else unitCostTonh 	= "0" end if
	
	' calculate the per unit costs
	dim unitCostTons, unitCostTonh', unitCostAdjTonh
	if totalTons = "null" or capCharge="null" then
		unitCostTons = "null"
	else
		if cdbl(totalTons) <> 0 then
			unitCostTons = "'" & (formatnumber(cdbl(capCharge) / cdbl(totalTons),numDecimalPlaces)) & "'"
		else
			unitCostTons = "'0'"
		end if
	end if
	
'	if totalTonHrs="null" then'auCharge="null" or 
'		unitCostAdjTonh = "null"
'	else
'		if cdbl(totalTonHrs) <> 0 then
'			unitCostAdjTonh = (formatnumber(cdbl(usageCharge) / cdbl(totalTonHrs),numDecimalPlaces)) & "'"
'		else
'			unitCostAdjTonh = "'0'"
'		end if
'	end if
	
'	if usageCharge="null" or totalTonHrs="null" then
'		unitCostTonh = "null"
'	else
'		if cdbl(totalTonHrs) <> 0 then
'			unitCostTonh ="'" &  formatnumber(cdbl(usageCharge) / cdbl(totalTonHrs),numDecimalPlaces) & "'"  
'		else
'			unitCostTonh = "'0'"
'		end if
'	end if
	
	' finished doing math, add quotes
	if usageCharge <> "null" then usageCharge = "'" & usageCharge & "'"
	if capCharge <> "null" then capCharge = "'" & capCharge & "'"
'	if auCharge <> "null" then auCharge = "'" & auCharge & "'"
	if totalTons <> "null" then totalTons = "'" & totalTons & "'"
	if totalTonHrs <> "null" then totalTonHrs = "'" & totalTonHrs & "'"
	
	if action="UPDATE" then
		strsql="update utilitybill_chilledwater set unit_credit=0, usagecharge=" & usagecharge& ", eecharge=" & eecharge& ", opCharge=" & opCharge& ", capCharge=" & capCharge& ", tempCharge=" & tempCharge& ", adjustments=" & adjustments& ", totalTons=" & totalTons& ",totalTonHrs=" & totalTonHrs& ",onPeakTonH=" & onPeakTonHrs& ",offPeakTonH=" & offPeakTonHrs& ", totalBillAmt=" & totalBillAmt& ", penalty_kwh=" & penalty& ", unitCostTons=" & unitCostTons& ", unitCostTonh=" & unitCostTonh& ", salesTax=" & salesTax & ", miscCharge =" & miscCharge & ", miscChargeDesc = "&miscChargeDesc&", note='"&note&"'  WHERE ypid="&YPID&" AND acctid='"&acctid&"' "'unitCostAdjTonh=" & unitCostAdjTonh& ", 
	elseif action="SAVE" then
		strsql="insert into utilitybill_chilledwater (unit_credit, ypid, acctid, usagecharge, eecharge, opCharge, capCharge, tempCharge, adjustments, totalTons,totalTonHrs,onPeakTonH,offPeakTonH, totalBillAmt, penalty_kwh, unitCostTons, unitCostTonh, salesTax, miscCharge, miscChargeDesc, note) values (0,'" & ypid & "','" & acctid & "'," & usagecharge & "," & eecharge& "," & opCharge & "," & capCharge& "," & tempCharge& "," & adjustments& "," & totalTons & "," & totalTonHrs & "," & onPeakTonHrs & "," & offPeakTonHrs & "," & totalBillAmt & "," & penalty & "," & unitCostTons& "," & unitCostTonh & "," & salesTax & "," & miscCharge & "," & miscChargeDesc & ", '"&note&"') "
	end if
end if
'response.write strsql
'response.end

cnn1.execute strsql
set cnn1=nothing

%><script>
document.location = "acctwdetail.asp?bldg=<%=bldg%>&acctid=<%=acctid%>&ypid=<%=ypid%>&bp=<%=bp%>&by=<%=by%>&utility=<%=utility%>"
</script>

