<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'2/21/2008 N.Ambo amended admin fee format to show correct percentage value (multiplied adminfee by 100)
'3/12/2008 N.Ambo amended sql statement which counts the meters per lease
'3/31/2009 N.Ambo added code to accomodate for new chiller utilities
'4/27/2009 K.Cheng added limit of 15 characters to meter location field on the bill (line 487)
'4/28/2009 N.Ambo reverted back to old accounting asp page dataoutput.asp until ready for SAP changes
'6/12/2009 Michelle T. Modify code line 653 to show HTML display of sewer charge for water bills and 
'and reformat energydetail and demanddetail in the same case condition to display 6 places after decimal on bills only demand charges for electricity.
Server.ScriptTimeout = 60*15
dim leaseid, ypid, building, pid, byear, bperiod, utilityid, detailed, meterbreakdown, calcintpeak, SJPproperties, ShowUsageDetails, ShowDemandDetails,onlinereview,totdeliverychrgs, maxmeters, textheader, masterTotal,demo, reject,billid
dim pdfsession,startCount,endCount,logo,meterinfo,unittype,serviceclass,serviceclassid,currBillsPageCount, invoiceNum, sInvoiceNo, energydetail, demanddetail

leaseid = trim(Request("l"))
ypid = trim(request("y"))
utilityid = trim(request("utilityid"))
building = trim(request("building"))
pid = trim(request("pid"))
byear = trim(request("byear"))
bperiod = trim(request("bperiod"))
logo = trim(request("logo"))
textheader = trim(request("textheader"))
demo = request("demo")
currBillsPageCount = request("currBillsPageCount")
invoiceNum = request("billid")

if trim(request("reject"))="" then reject = 0 else reject = 1
if trim(request("detailed"))="" then detailed = "false" else detailed = trim(request("detailed"))

if trim(request("SJPproperties"))="true" then SJPproperties = true else SJPproperties = false
if trim(request("meterbreakdown"))="no" then meterbreakdown = false else meterbreakdown = true
if trim(request("summaryusage"))="true" then showusagedetails= true else showusagedetails = false
if trim(request("summarydemand"))="true" then showdemanddetails = true else showdemanddetails = false
'response.Write(isempty(session("xmlUserObj")))
'response.End()

'if pdfsession = "pdffdp" then
Dim isPDF
isPDF = false
if ((request.servervariables("HTTP_REFERER")="Webster://Internal/315" and isempty(session("xmlUserObj"))) OR (pdfsession = "pdffdp"))  then 'this is for pdf sessions
 loadNewXML("activepdf")
loadIps(0)
isPDF = true
end if

dim cnn1, rst1, rst2, rst3,rst4,rst5, bldgrs, usagelabel, sql
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
Set rst3 = Server.CreateObject("ADODB.recordset")
Set rst4 = Server.CreateObject("ADODB.recordset")
Set rst5 = Server.CreateObject("ADODB.recordset")
'if trim(getbuildingIP(building))="0" then response.redirect "/eri_th/pdfmaker/genergyInvoice.asp?"&request.servervariables("QUERY_STRING") else cnn1.Open getLocalConnect(building)
cnn1.Open getLocalConnect(building)
usagelabel = "used"
dim DBmainIP
DBmainIP = ""
if utilityid="" and leaseid<>"" then
  rst1.open "SELECT utility FROM tblleasesutilityprices WHERE leaseutilityid="&leaseid, cnn1
  if not rst1.eof then utilityid = cint(rst1("utility"))
  rst1.close
  set rst1 = nothing
end if

'select case utilityid
'case 3, 10,6,1,4
	'select case utilityid
	'case 6,1,4
		'usagedivisor = 1
	'case else
		'usagedivisor = 100
 ' End select 
%>


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Untitled Document</title>
</style>
</head>

<body bgcolor="white">
<!--tenant-->

<%
' 690%
if leaseid<>"" and ypid<>"" and utilityid <> "" then
	showtenantbill leaseid,"", "","",byear,bperiod,building,meterbreakdown,utilityid, ypid
	'footer ypid,leaseid, utilityid
elseif building<>"" then
    sql = "SELECT b.id as billid, b.leaseutilityid, ypid, b.utility, extusg FROM tblbillbyperiod b WHERE reject=0 and bldgnum='"&building&"' and billyear="&byear&" and billperiod="&bperiod
    if isnumeric(utilityid) then sql = sql & " and utility="&utilityid
    sql = sql & "  ORDER BY TenantName"
		
	'response.Write(sql & " <p> ")
	dim blgdrset
	Set blgdrset = Server.CreateObject("ADODB.recordset")
	blgdrset.open sql, cnn1
	do until blgdrset.eof
		dim templid, tempypid, temputility
		templid = trim(blgdrset("leaseutilityid"))
		tempypid = trim(blgdrset("ypid"))
		temputility = trim(blgdrset("utility"))
		
		Dim metercountInfo
		Set metercountInfo = Server.CreateObject("ADODB.recordset")
		billid = blgdrset("billid")
		'metercountInfo.open "select count(*) as metercount from tblmetersbyperiod tm,buildings b,meters m where tm.bldgnum =b.bldgnum and tm.meternum=m.meternum and b.bldgnum = m.bldgnum and bill_id="&billid, cnn1
		'3/12/2008 N.Ambo replaced with statement below so that join will be on meterid to avoid duplicates
		metercountInfo.open "select count(*) as metercount from tblmetersbyperiod tm,buildings b,meters m where tm.bldgnum =b.bldgnum and tm.meterid=m.meterid and b.bldgnum = m.bldgnum and bill_id="&billid, cnn1
		
		dim tempMaxPageCount
		tempMaxPageCount = metercountInfo("metercount") \ 40 + 1
		
		if metercountInfo("metercount") > 5 then
			tempMaxPageCount = tempMaxPageCount + 1
		end if
		
		metercountInfo.Close()

		if CInt(tempMaxPageCount) = CInt(currBillsPageCount) then
			showtenantbill templid,"", "","",byear,bperiod,building,meterbreakdown,temputility, tempypid
			'footer tempypid,templid, temputility
			%>
			<WxPrinter PageBreak>
			<%
		end if

		blgdrset.movenext
	loop
	blgdrset.Close()
end if
%>
</body>
</html>

<%
sub showtenantbill(leaseid, billid, flag, meterlist,byear,bperiod,building,meterbreakdown,utilityid, ypid)


		
dim cnn1, rst1, rst2, rst3, bldgrs, usagelabel, sql,usagedivisor,tot_kwhusedon,tot_kwhusedoff,tot_kwhusedint,tot_kwhused,rst555,utilityname
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
Set rst3 = Server.CreateObject("ADODB.recordset")
Set rst555 = Server.CreateObject("ADODB.recordset")
'if trim(getbuildingIP(building))="0" then response.redirect "/eri_th/pdfmaker/genergyInvoice.asp?"&request.servervariables("QUERY_STRING") else cnn1.Open getLocalConnect(building)
cnn1.Open getLocalConnect(building)
Set rst555 = Server.CreateObject("ADODB.recordset")
'get utility name
rst555.open "SELECT utility FROM tblutility WHERE utilityid="&utilityid, getConnect(pid,building,"billing")
if not rst555.eof then
	utilityname = rst555("utility")
end if
rst555.close
'select case utilityid
'case 3, 10,6,1,4

'usagedivisor = 1

'case else

'usagedivisor = 100

 ' end select 
usagelabel="used"
'Electricity class="style1"
' background="file:///J|/appserver1/eri_TH/pdfMaker/PA_Logos.jpg"
'<td width="690" height="76" background="file:///J|/appserver1/eri_TH/pdfMaker/PA_Logos.jpg">
'<img src="oucooling.gif" alt="OUCooling / A TRIGEN*CINERGY SOLUTIONS Service" width="203" height="123" border="0">
'<td width="690" height="76" <img src="PA_Logos.jpg" width="203" height="123" border="0">
%>
<table width="100%" align="center" ID="Table1">
<tr>
<td align="center">
<table width="690" border="0" cellpadding="0" cellspacing="0" ID="Table2">
  <tr>
	<td width="690" height="50" background="file://J|//testserver1/eri_TH/pdfMaker/PA_Logos.jpg">
      <font color="#FFFFFF"><h3>Utility Invoice - <%=utilityname%></h3></font>
    </td>
  </tr>
  <% 
	if (invoiceNum = "") then
	    sql = "SELECT bp.id as billid " & _
			    "FROM tblbillbyPeriod bp " & _
			    " WHERE reject=0 and bp.ypid="&ypid&" and bp.leaseutilityid="&leaseid 
    			
	    rst2.open sql, cnn1
	    if not rst2.eof then
  		    billid = rst2("billid")
  	    end if	
      	
	    rst2.close
	 else
	    billid = invoiceNum
	 end if 

	if utilityid=2 then
		sql = "SELECT bp.id as billid, RIGHT(REPLACE(RTRIM(SPACE(5) + STR(BN.InvoiceSeqNo)),' ','0'),5) AS InvoiceSeqNo, BN.BillType " & _
			    "FROM tblbillbyPeriod bp, tblPAInvoiceBillNumbers BN " & _
				" WHERE bp.ypid="&ypid&" and bp.leaseutilityid="&leaseid & _
					"and BN.billid = bp.id and BN.billid = " + cstr(billid)

	elseif utilityid=3 or utilityid =10 then
			sql = "SELECT bp.id as billid, RIGHT(REPLACE(RTRIM(SPACE(5) + STR(BN.InvoiceSeqNo)),' ','0'),5) AS InvoiceSeqNo, BN.BillType " & _
			    "FROM tblbillbyPeriod bp, tblPAWaterBillNumbers BN " & _
				" WHERE bp.ypid="&ypid&" and bp.leaseutilityid="&leaseid & _
					"and BN.billid = bp.id and BN.billid = " + cstr(billid)

	'3/31/2009 N.Ambo added new chiller utilities
	elseif utilityid =19 or utilityid=18 or utilityid=20 then
			sql = "SELECT bp.id as billid, RIGHT(REPLACE(RTRIM(SPACE(5) + STR(BN.InvoiceSeqNo)),' ','0'),5) AS InvoiceSeqNo, BN.BillType " & _
			    "FROM tblbillbyPeriod bp, tblPAChWaterBillNumbers BN " & _
				" WHERE bp.ypid="&ypid&" and bp.leaseutilityid="&leaseid & _
					"and BN.billid = bp.id and BN.billid = " + cstr(billid)
		'end if
	end if	  

  rst2.open sql, cnn1
  sInvoiceNo = "0000000" ' Indicates No invoice number assigned
  if not rst2.eof then
	If not isNull(rst2("BillType")) Then 	
		sInvoiceNo = rst2("BillType") &  rst2("InvoiceSeqNo")
	Else
		sInvoiceNo = "0000000" ' Indicates an Error While Generating The Invoice Number
	End If
end if

'response.write sql &"<BR>"
'response.write "sInvoiceNo" & sInvoiceNo
'response.end

'3/31/2009 N.Ambo included utility types for chiller
select case utilityid
case 3, 10,6,1,4,18,19,20
	select case utilityid
	  case 6,1,4
		usagedivisor = 1
	case else
		usagedivisor = 100
     end select                                                            
'water                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           

 sql = "SELECT isnull(r.adjustment,0) as adjustment ,isnull(r.FuelADJ,0) as usefuelADJ,ContactPhone,ContactName,b.bldgname,isnull(b.btzip,'') as billzip, b.portfolioid, isnull(b.btstrt,'') as billto, r.addonfee as myaddonfee, isnull(r.energydetail,'0') as energydetail, isnull(r.demanddetail,'0') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress,isnull(r.adminfee,0) as adminfee,isnull(r.servicefee,0) as servicefeenonull, isnull(r.fueladj,0) as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(r.adjustment, 0) as adjustmentnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.[id] as rtid, datediff(day, datestart,dateend)+1 as days, (select case count(distinct isnull(addonfee,0)) when 0 then 0 else 1 end as aoncnt from tblmetersbyperiod where bill_id=r.id group by leaseutilityid,ypid,bill_id) as showaddonfee, r.datestart-1 as datebegin,r.invoice_note as invoiceNote, isnull(rate_servicefee_dollar,0) as rateservicefee_dollar, r.*, l.onlinebill,l.corpStreet,l.corpCity,l.corpState,l.corpZip,l.corpCountry,pa.acctnumber,pa.leasenumber,pa.seqnumber  FROM tblbillbyperiod r, tblleases l, buildings b,ratetypes rt,custom_PABT pa WHERE pa.acctnumber=l.billingid and r.ratetenant=rt.id AND b.bldgnum=r.bldgnum and l.billingid = (SELECT billingid FROM tblleasesutilityprices lup WHERE lup.leaseutilityid=r.leaseutilityid) and r.id="&billid
    
   case else
//	this original query was altered for PA tenant summary to match bill summary -Michelle T.
//sql = "SELECT isnull(r.adjustment,0) as adjustment ,isnull(r.FuelADJ,0) as usefuelADJ,ContactPhone,ContactName,b.bldgname,isnull(b.btzip,'') as billzip, b.portfolioid,isnull(r.btstrt,'') as billto, isnull(r.energydetail,'') as energydetail, isnull(r.demanddetail,'') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress, r.fueladj as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(r.adjustment, 0) as adjustmentnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.id as rtid, datediff(day, datestart,dateend) as days, case rt.[type] when 'AVG Cost 1' then round(avgkwh,6) when 'AVG COST 2' then round(unitcostkwh,6) else ' ' end as akwhdisplay, case rt.[type] when 'AVG COST 2' then round(isnull(tunitcostkw,0),6) else ' ' end as akwdisplay, case when Totalkw=0 then 0 else ((Totalkwh/Totalkw)/(datediff(day, ypiddatestart,ypiddateend)+1)*24) end as loadfactor, isnull(r.adminfee,0) as adminfee, isnull(r.adminfeedollar,0) as adminfeedollar, r.billperiod, r.billyear, r.datestart, r.dateend, isnull(r.servicefee,0) as servicefeenonull,r.addonfee as myaddonfee, r.unit_credit, isnull(r.subTotal,0) as subTotal, r.tenantname, lup.calcintpeak, r.*, l.onlinebill,l.corpStreet,l.corpCity,l.corpState,l.corpZip,l.corpCountry,pa.acctnumber,pa.leasenumber,pa.seqnumber FROM rpt_bill_summary r, tblleases l, buildings b, dbo.ratetypes rt, tblleasesutilityprices lup,custom_PABT pa WHERE pa.acctnumber=l.billingid and r.[type]=rt.id AND b.bldgnum=r.bldgnum and lup.leaseutilityid=r.leaseutilityid and l.billingid=lup.billingid and r.billid="&billid

'electricty
sql = "SELECT isnull(r.adjustment,0) as adjustment ,isnull(r.FuelADJ,0) as usefuelADJ,ContactPhone,ContactName,b.bldgname,isnull(b.btzip,'') as billzip, b.portfolioid,isnull(r.btstrt,'') as billto, isnull(r.energydetail,'') as energydetail, isnull(r.demanddetail,'') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress, r.fueladj as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(r.adjustment, 0) as adjustmentnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.id as rtid, datediff(day, datestart,dateend)+1 as days, case rt.[type] when 'AVG Cost 1' then round(avgkwh,6) when 'AVG COST 2' then round(unitcostkwh,6) else ' ' end as akwhdisplay, case rt.[type] when 'AVG COST 2' then round(isnull(tunitcostkw,0),6) else ' ' end as akwdisplay, case when Totalkw=0 then 0 else ((Totalkwh/Totalkw)/(datediff(day, ypiddatestart,ypiddateend)+1)*24) end as loadfactor, isnull(r.adminfee,0) as adminfee, isnull(r.adminfeedollar,0) as adminfeedollar, r.billperiod, r.billyear, r.datestart-1 as datebegin, r.dateend, isnull(r.servicefee,0) as servicefeenonull,r.addonfee as myaddonfee, r.unit_credit, isnull(r.subTotal,0) as subTotal, r.tenantname, lup.calcintpeak, r.*, l.onlinebill,l.corpStreet,l.corpCity,l.corpState,l.corpZip,l.corpCountry,pa.acctnumber,pa.leasenumber,pa.seqnumber FROM rpt_bill_summary r, tblleases l, buildings b, dbo.ratetypes rt, tblleasesutilityprices lup,custom_PABT pa WHERE pa.acctnumber=l.billingid and r.[type]=rt.id AND b.bldgnum=r.bldgnum and lup.leaseutilityid=r.leaseutilityid and l.billingid=lup.billingid and r.billid="&billid

end select

'response.write sql
'response.end
rst3.open sql, cnn1, 2
 if not rst3.eof then
if utilityid=2 then
 if rst3("calcintpeak") then calcintpeak = true else calcintpeak = false
else
calcintpeak = false
end if

pid = rst3("portfolioid")
unittype = rst3("unittype")
serviceclass = rst3("rt")
serviceclassid = rst3("rtid")
energydetail = rst3("energydetail")
demanddetail = rst3("demanddetail")
dim acctnumber,seqnumber,leasenumber,corpStreet,corpCity,corpState,corpZip
acctnumber=rst3("acctnumber")
seqnumber=rst3("seqnumber")
leasenumber=rst3("leasenumber")
corpStreet=rst3("corpStreet")
corpCity=rst3("corpCity")&","
corpState=rst3("corpState")&","
corpZip=rst3("corpZip")

if corpStreet="" then
corpStreet=replace(rst3("tstrt"),vbNewLine,"<br>")
corpCity=rst3("tcity")&","&  rst3("tstate")&","&  rst3("tzip")
corpState=""
corpZip=""


end if


%>
  <tr>
    <td height="30"><table width="100%" height="30" border="0" cellpadding="0" cellspacing="0" ID="Table3">
      <tr>
        <td height="20"><font face="Arial" size="2"><b>Account No.</b></font></td>
        <td height="20"><font face="Arial" size="2"><b>Sequence No.</b></font></td>
        <td height="20"><font face="Arial" size="2"><b>Lease No.</b></font></td>
        <td height="20"><font face="Arial" size="2"><b>Invoice No.</b></font></td>
        <td height="20"><font face="Arial" size="2"><b>Invoice Date</b></font></td>
      </tr>
      <tr>
        <td height="20"><font face="Arial" size="2"><%=rst3("tenantnum")%></font></td>
        <td height="20"><font face="Arial" size="2"><%=seqnumber%></font></td>
        <td height="20"><font face="Arial" size="2"><%=leasenumber%></font></td>
        <td height="20"><font face="Arial" size="2"><%=sInvoiceNo%></font></td>
        <td height="20"><font face="Arial" size="2"><%if isdate(rst3("postdate")) then response.write formatdatetime(rst3("postdate"),2)%></font></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="100" width="690"><table width="100%" height="100" border="0" ID="Table4">
      
      <tr>
        <td height="50" valign="top"><table width="100%" height="50" border="0" cellpadding="0" cellspacing="0" ID="Table5">
          <tr>
            <td height="20"><font face="Arial" size="2"><b>Name:</b></font></td>
            <td height="20"><font face="Arial" size="2"><%=rst3("tenantname")%></font></td>
          </tr>
          <tr>
            <td height="20" valign="top"><font face="Arial" size="2"><b>Corporate Address:</b></font></td>
            <td height="20"><font face="Arial" size="2"><%=corpstreet%></font></td>
          </tr>
          <tr>
            <td height="20">&nbsp;</td>
            <td height="20"><font face="Arial" size="2"><%=corpcity&corpstate&corpzip%></font></td>
            <!--41st Floor-->
          </tr>
          <tr>
            <td height="20">&nbsp;</td>
            <td height="20"><font face="Arial" size="2">&nbsp;</font></td>
          </tr>
          <tr>
            <td height="20">&nbsp;</td>
            <td height="20"><font face="Arial" size="2">&nbsp;</font></td>
          </tr>
        </table></td>
        <td valign="top" width="50%"><table width="100%"  height="50" border="0" cellpadding="0" cellspacing="0" ID="Table6">
          <tr>
            <td height="20">&nbsp;</td>
            <td height="20">&nbsp;</td>
          </tr>
          <tr>
            <td height="20"><font face="Arial" size="2"><b>Billing Address:</b></font></td>
            <td height="20"><font face="Arial" size="2"><%=rst3("tenantname")%></font></td> 
          </tr>
          <tr>
            <td height="20">&nbsp;</td>
            <td height="20"><font face="Arial" size="2"><%=replace(rst3("tstrt"),vbNewLine,"<br>")%></font></td>
          </tr>
          <tr> 
            <td height="20">&nbsp;</td>
            <td height="20"><font face="Arial" size="2"><%=rst3("tcity")&","&  rst3("tstate")&","&  rst3("tzip")%></font></td>
          </tr>
          <tr>
            <td height="20"><font face="Arial" size="2"><!--Attn to:--></font></td>
            <td height="20"><font face="Arial" size="2"><!--Bill Payyer--></font></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="200"><table width="680" height="170" border="0" cellpadding="0" cellspacing="2" ID="Table7">
      
      <tr>
        <td height="10" colspan="11" valign="top"><font face="Arial" size="2"><b><table width="597" height="20" border="0" ID="Table8">
          
          <tr>
            <td width="36" valign="top"><font face="Arial" size="2">Period:</font></td>
            <td width="41" valign="top"><font face="Arial" size="2"><%=rst3("billyear")%>/<%=rst3("billperiod")%></font></td>
            <td width="33" valign="top"><font face="Arial" size="2">From:</font></td>
            <td width="60" valign="top"><font face="Arial" size="2"><%=rst3("datebegin")%></font></td>
            <td width="18" valign="top"><font face="Arial" size="2">To:</font></td>
            <td width="60" valign="top"><font face="Arial" size="2"><%=rst3("dateend")%></font></td>
            <td width="200" align="center" valign="top"><font face="Arial" size="2">Number of days in BillPeriod. Days:</font></td>
            <td width="160" align="left" valign="top"><font face="Arial" size="2"><%=rst3("dateend")-rst3("datebegin")%></font></td>         
            </tr>
          
        </table></b></font></td>
        </tr>
      
		<%
		dim rst6,extusage
		Set rst6 = Server.CreateObject("ADODB.recordset")
		if extusage = "" then 
	     sql = "select extusg from tblbillbyperiod where id = " & billid
		rst6.open sql, cnn1, 2
	
			if not rst6.eof and trim(rst6("extusg")) <> "" then 
				extusage = rst6("extusg")
			else 
				extusage = false
			end if 
			rst6.close	
		end if

		Set meterinfo = Server.CreateObject("ADODB.recordset")
		'if flag then 
		'meterinfo.open "select * from tblmetersbyperiod where meternum not in ("&meterlist&") and bill_id="&billid, cnn1
		'else
		
		'2/25/2008 nambo ameneded te statement to join on meterid instead of meternum because duplicates were being produced
		meterinfo.open "select * from tblmetersbyperiod tm,buildings b,meters m where tm.bldgnum =b.bldgnum and tm.meterid=m.meterid and b.bldgnum = m.bldgnum and bill_id="&billid&" order by tm.meternum", cnn1
		
		'end if 
		 
		dim tot_onpeak,	tot_offpeak,metercnt,tot_demand,tot_demandoff_p,tot_demand_p,tot_demand_c
		metercnt=0
		tot_onpeak 		=0
		tot_offpeak		=0
		'tot_intpeak		=0
		'tot_demandint_p	=0
		tot_demandoff_p	=0 
		tot_kwhused		=0
		tot_kwhusedoff	=0
		tot_kwhusedint	=0
		tot_demand_p	=0
		tot_demand_c	=0
		dim AndoMeter
		dim maxMeterCount
 		maxMeterCount = 40
 		
		while not meterinfo.EOF 
		metercnt = metercnt + 1
		
		'check if we need to move to a new page and if so repeat the header
		if metercnt mod maxMeterCount = 1 and (isPDF or metercnt = 1) then
			if metercnt <> 1 then
				%>
					<tr>
						<td colspan="10" align="right">Continues On The Next Page<WxPrinter PageBreak></td>
					</tr>
				<%
			end if		
			%>
			<tr>
			<td width="8%" class="style9">&nbsp;</td>
				<td width="7%" class="style9">&nbsp;</td>
				<td width="10%" class="style9">&nbsp;</td>
				<td width="11%" class="style9">&nbsp;</td>
				<td colspan="2" align="center"><font face="Arial" size="2"><b>READINGS</b></font></td>
				<%  if  utilityid = 10 or utilityid = 3 then %>
				<td colspan="3" align="center"><font face="Arial" size="2"><b>CONSUMPTION</b></font></td>
				<% else %>
				<td colspan="3" align="center"><font face="Arial" size="2"><b>CONSUMPTION</b></font></td>
				<td align="center"><font face="Arial" size="2"><b>DEMAND</b></font></td>
				<% end if %>  
				</tr>
			<tr>
				<td><font face="Arial" size="2">Meter No.</font></td>
				<td><font face="Arial" size="2">Facility</font></td>
				<td align="center"><font face="Arial" size="2">Location</font></td>
				<td align="center"><font face="Arial" size="2">Multiplier</font></td>
				<td width="11%" align="center"><font face="Arial" size="2">Previous</font></td>
				<td width="9%" align="center"><font face="Arial" size="2">Current</font></td>
				<% if  utilityid = 10 or utilityid = 3 then %>
				<td colspan="3" width="20%" align='center'><font face="Arial" size="2">CCF Used</font></td>
				<td>&nbsp;</td>
				<% else %>
				<td width="11%" align="center"><font face="Arial" size="2">On Peak</font></td>
				<td width="11%" align="center"><font face="Arial" size="2">Off Peak</font></td>
				<td width="14%" align="center"><font face="Arial" size="2">Total Usage</font></td>
				<td width="8%" align="center"><font face="Arial" size="2">KW</font></td>
				<% end if %>    
			</tr>
			<%
		end if
		
		
		%>
      <tr>
        <td width="63" align="left" nowrap><font face="Arial" size="2"><%=meterinfo("meternum")%></font></td>
        <td width="47" align ="left" nowrap><font face="Arial" size="2"><%=meterinfo("bldgname")%></font></td>
       <!-- <td class="style9" nowrap>&nbsp;87</td>-->
        <td width="72" align="center" nowrap><font face="Arial" size="2"><%=left(meterinfo("location"),15)%></font></td>    
        <td width="61" align="center" nowrap><font face="Arial" size="2"><%=formatnumber(meterinfo("manualmultiplier"),2)%></font></td>
        <td width="73" nowrap align="center"><font face="Arial" size="2"><%=Formatnumber(meterinfo("rawprevious"),2)%></font></td>
        <td width="72" nowrap align="center"><font face="Arial" size="2"><%=Formatnumber(meterinfo("rawcurrent"),2)%></font></td>
        <% if utilityid = 10 or utilityid = 3 then %>
		<td class="style9">&nbsp;</td>
		<td><font face="Arial" size="2"><%=Formatnumber(meterinfo("used"),2)%></font></td>
        <td class="style9">&nbsp;</td>
		 <% else %>
		<td width="57" nowrap align="center"><font face="Arial" size="2"><%=Formatnumber(meterinfo("onpeak"),2)%></font></td>
        <td width="77" nowrap align="center"><font face="Arial" size="2"><%=Formatnumber(meterinfo("offpeak"),2)%></font></td>
        <td width="91" nowrap align="center"><font face="Arial" size="2"><%=Formatnumber(meterinfo(usagelabel),2)%></font></td>
        <td width="55" nowrap align="center"><font face="Arial" size="2"><%=Formatnumber(meterinfo("demand_P"),2)%></font></td>
		 <% end if %> 
	  </tr>
       <%
		
		

		tot_demand = tot_demand +cdbl(meterinfo("demand_P"))
		tot_onpeak=tot_onpeak+Formatnumber(meterinfo("onpeak"),2)
		tot_offpeak= tot_offpeak + formatnumber(cdbl(meterinfo("offpeak")),2)
		AndoMeter= Andometer + meterinfo("meternum")+","
		
		'testing totalusage and other stuff
		if calcintpeak then
		'tot_intpeak= tot_intpeak+ formatnumber(cdbl(meterinfo("intpeak")),2)
		tot_demandoff_p= tot_demandoff_p + formatnumber(cdbl(meterinfo("demand_off")),2)
		tot_demandint_p= tot_demandint_p + formatnumber(cdbl(meterinfo("demand_int")),2)
	end if 
	'tot_offpeak= tot_offpeak + formatnumber(cdbl(meterinfo("offpeak")),2)
	'tot_onpeak = tot_onpeak + cdbl(rst1("onpeak"))
	if utilityid = 3 or utilityid = 10 then
		tot_kwhused= tot_kwhused + (cdbl(meterinfo(usagelabel))/usagedivisor)
	else 
		if extusage then 
			if meterinfo("mextusg") then 
				tot_kwhusedon 	= tot_kwhusedon + (cdbl(meterinfo("used"))/usagedivisor)
				tot_kwhusedoff 	= tot_kwhusedoff + (cdbl(meterinfo("usedoff"))/usagedivisor)
				tot_kwhusedint 	= tot_kwhusedint + (cdbl(meterinfo("usedint"))/usagedivisor)
				tot_kwhused		= tot_kwhusedon + tot_kwhusedoff + tot_kwhusedint
			else
				tot_kwhusedon 	= tot_kwhusedon + (cdbl(meterinfo("onpeak"))/usagedivisor)
				tot_kwhusedoff 	= tot_kwhusedoff + (cdbl(meterinfo("offpeak"))/usagedivisor)
				tot_kwhusedint 	= tot_kwhusedint + (cdbl(meterinfo("intpeak"))/usagedivisor)
				tot_kwhused		= tot_kwhusedon + tot_kwhusedoff + tot_kwhusedint
			end if
		else
			tot_kwhused = tot_kwhused + (cdbl(meterinfo(usagelabel)))
			end if
	end if
	'testing anodmic
		
		
		' you have to keep track of the meters, if there are more than five..then you
		' break it apart here..
		' close the table before
		' and open the new table.
		
		
	meterinfo.movenext ' move to the next record. and see if it was the last below.
	
		
		'if metercnt  >5 then
		'half Formatnumber(tot_onpeak,2),Formatnumber(tot_offpeak,2)
		'showtenantbill leaseid, billid, 1, AndoMeter
		'exit while
		'exit sub
		'end if
		wend
		'Andomic(AndoMeter)
		%>
    
	  <tr>
        <td><font face="Arial" size="2">Totals</font></td>
        <!--<td class="style9">&nbsp;</td>-->
        <td class="style9">&nbsp;</td>
        <td class="style9">&nbsp;</td>
        <td class="style9">&nbsp;</td>
        <td class="style9">&nbsp;</td>
        <td class="style9">&nbsp;</td>
       <%  if  utilityid = 10 or utilityid = 3 then%>
	   <td class="style9" colspan="3" align='center' style="border-top:1px solid black" ><%=Formatnumber(tot_kwhused,2)%></td>
	    <% else %>
	    <td class="style9" align="center" style="border-top:1px solid black" ><%=Formatnumber(tot_onpeak,2)%></td>
        <td class="style9" align="center" style="border-top:1px solid black" ><%=Formatnumber(tot_offpeak,2)%></td>
        <td class="style9" align="center" style="border-top:1px solid black"><%=Formatnumber(tot_kwhused,2)%></td>
        <td width="55" align="center" class="style9" style="border-top:1px solid black"><%=Formatnumber(tot_demand,2)%></td>
     	 <% end if %>    
	  </tr>
    </table>
	<%
		if metercnt > 5 then
			%><WxPrinter PageBreak><%
		end if
	%>
    </td>
  </tr>
<%
dim ELFac,akwhdisplay,akwdisplay,rt,adminfeecalc
'fuel
'if isnumeric(rst3("fAdj")) and not(isnull(rst3("fAdj"))) and trim(rst3("fAdj"))<>"0" then ELFac=formatnumber(rst3("fAdj"), 6) else ELFac= "NA"
if isnumeric(rst3("fuel")) and not(isnull(rst3("fuel"))) and trim(rst3("fuel"))<>"0" then ELFac=formatnumber(rst3("fuel"), 6) else ELFac= "NA"
if utilityid = 2 then
if isnumeric(rst3("akwhdisplay")) and cdbl(rst3("akwhdisplay"))<>0 then akwhdisplay =FormatCurrency(rst3("akwhdisplay"),6) else akwhdisplay= "NA"
if isnumeric(rst3("akwdisplay")) and cdbl(rst3("akwdisplay"))<>0 then akwdisplay =FormatCurrency(rst3("akwdisplay"),6) else akwdisplay ="NA"
end if
'if isnumeric(rst3("unit_credit")) and trim(rst3("unit_credit"))<>"0" then rt=rst3("unit_credit") else rt="NA"
rt = rst3("rt")

'if trim(rst3("Adminfeedollar"))<>"" then  adminfeecalc = cdbl(rst3("Adminfeedollar")) else adminfeecalc = 0
'usefuelADJ
middle rst3("subtotal"),rst3("fuel"),rst3("servicefee"),rst3("Adminfee"),rst3("Adminfeedollar"),rst3("Taxnonull"),rst3("totalamtnonull"),rt,leaseid,byear,bperiod,billid,building,unittype,meterbreakdown,calcintpeak,ElFac,akwhdisplay,akwdisplay,rst3("energydetail"),rst3("demanddetail"),utilityid,rst3("adjustment"), rst3("credit")

end if

end sub

sub middle (subtotal,FuelAdj,servicefee,Adminfee,subtotal2,Taxnonull,totalamtnonull,rt,leaseid,billyear, _
		billperiod,billid,building,unittype,meterbreakdown,calcintpeak,ElFac,akwhdisplay,akwdisplay,energydetail, _
		demanddetail,utilityid,adjustment,creditamt)
dim hidedemand,rst6,extusage
if FuelAdj = "" or isnull(FuelAdj) then
FuelAdj = 0
end if
'Set rst6 = Server.CreateObject("ADODB.recordset")
hidedemand=""
'if extusage = "" then 
	'sql = "select extusg from tblbillbyperiod where id = " & billid
	'rst6.open sql, cnn1, 2
	
	'if not rst6.eof and trim(rst6("extusg")) <> "" then 
		'extusage = rst6("extusg")
	'else 
		'extusage = false
	'end if 
	'rst6.close	
'end if
if isnull(energydetail)  then
energydetail=0
end if

%>
<tr>
    <td height="100">
	<table width="100%" height="100" border="0" style="border:2px solid #999999" ID="Table19">
      <tr>
        <td height="20px" align="center"  valign="top"><div align="left"><font face="Arial" size="2">*Readings with a suffix "E" are estimated; otherwise readings are actual. Reasons for any Estimated are shown below.</font></div><br />DETAILED BILLED AREA / ESTIMATION NOTES
        </td>
       </tr>
       <%
        if (detailed = "false") then
        %>
       <tr>
        <td height="80px" align="left" valign="top">
        <%
			Dim commentinfo
			Set commentinfo = Server.CreateObject("ADODB.recordset")
			commentinfo.open "select distinct Note from Misc_Inv_Credit mic, tblmetersbyperiod tm where tm.LeaseUtilityId =mic.LeaseUtilityId and tm.BillYear =mic.BillYear and tm.BillPeriod =mic.BillPeriod and bill_id="&billid, cnn1
			
			while not commentinfo.EOF
				%>
				<%=commentinfo("Note")%><br/>
				<%
				commentinfo.MoveNext()
			wend
        %>        
        </td>
      </tr>
      <% else %>
          <tr>
            <td>
                <table width="100%" align="center" border="0">
                 <tr><td><%if trim(energydetail)<>"" then%><%if ucase(trim(serviceclass)) <> "SC4 RA1" and ucase(trim(serviceclass)) <> "SC4 RA2" then%>Consumption<%else%>Invoice<%end if%> Details:<%end if%></td>
	                   
	                   <%if (utilityid = 3) or (utilityid=10) then%>
	                <td><%if trim(demanddetail)<>"" then%>Sewer Charges: <%end if%></td>
	                    <%else %>
	                <td><%if trim(demanddetail)<>"" then%>Demand Details :<%end if%></td><%end if%>
	                                </tr>
                <tr><td valign="top"><%if energydetail<>"" then%><%=replace(replace(energydetail,"|","<br>")," ","&nbsp;")%><%end if%></td>
                    <td valign="top"><%if demanddetail<>"" then%><%=replace(replace(demanddetail,"|","<br>")," ","&nbsp;")%><%end if%></td></tr>
              
              </table>
            </td>    
          </tr>
      <% end if %>
    </table></td>
</tr>
  <tr>
    <td height="100"><table width="100%" height="100" border="0" ID="Table20">
      <tr>
        <td>
        <table width="100%" height="100" border="0" bgcolor="#999999" style="border:2px solid black" ID="Table21">
          <tr>
            <td><font face="Arial" size="2">Service Fee</font></td>
            <td><font face="Arial" size="2"><%=Formatcurrency(servicefee,2)%>/per meter</font></td>
          </tr>
          <tr>
            <td><font face="Arial" size="2">Admin Fee</font></td>
            <td><font face="Arial" size="2"><%=Formatnumber(Adminfee*100,2)%>%</font></td>
          </tr>
          
         <% if utilityid = 2 then%>
		  <tr>
            <td><font face="Arial" size="2">EL Adjust Factor</font></td>
            <td><font face="Arial" size="2"><%=ElFac%></font></td>
          </tr>
		  <tr>
            <td><font face="Arial" size="2">Average Cost KWH</font></td>
            <td><font face="Arial" size="2"><%=akwhdisplay%></font></td>
          </tr>
          <tr>
            <td><font face="Arial" size="2">Average Cost KW</font></td>
            <td><font face="Arial" size="2"><%=akwdisplay%></font></td>
          </tr>
		  <% '3/31/2009 N.Ambo added exclusions for new chiller utilities
		  elseif (utilityid <> 19 and utilityid <> 18 and utilityid <> 20) then 
		   dim utilitylabel
			Select Case Utilityid
				case 3
					utilitylabel="CW per CCF"
				case 10
					utilitylabel="HW per CCF"
				end Select 
		  		   %>
		  
		  <tr>
            <td><font face="Arial" size="2"><%=utilitylabel%></font></td>
            <td><font face="Arial" size="2"><%=FormatCurrency(energydetail,6)%></font></td>
          </tr>
          <tr>
            <td><font face="Arial" size="2">SC per CCF</font></td>
            <td><font face="Arial" size="2"><%=FormatCurrency(demanddetail,6)%></font></td>
          </tr> 
		  
          <% end if%>
		  <tr>
            <td><font face="Arial" size="2">Rate</font></td>
            <td><font face="Arial" size="2"><%=rt%></font></td>
          </tr>
        </table>
        </td>
        <td>
        <table width="100%" height="100" border="0" style="border:2px solid #666666" ID="Table22">
         <tr>
            <td align="center" valign="middle">Extended Area for Credit notes </td>
          </tr>
        </table>
        </td>
        <td><table width="100%" height="100" border="0" ID="Table23">
          <!--<tr>
            <td><font face="Arial" size="2">SubTotal</font></td>
            <td align="right"><font face="Arial" size="2"><'%=Formatcurrency(Subtotal,2)%></font></td>
          </tr>-->
          <% if utilityid = 2 then %>
		  <tr>
            <td><font face="Arial" size="2">Fuel Adj</font></td>
            <td align="right"><font face="Arial" size="2"><%=Formatcurrency(FuelAdj,2)%></font></td>
          </tr>
          <% end if %>
          <% if Cdbl(creditamt) <> 0 then %>
          <tr>
            <td><font face="Arial" size="2">Adjustment</font></td>
            <td align="right"><font face="Arial" size="2">(<%=Formatcurrency(creditamt,2)%>)</font></td>
          </tr>
          <% else %>
		  <tr>
            <td><font face="Arial" size="2">Adjustment</font></td>
            <td align="right"><font face="Arial" size="2"><%=Formatcurrency(adjustment,2)%></font></td>
          </tr>
          <% end if %>
          <tr>
            <td><font face="Arial" size="2">Service Fee</font></td>
            <td align="right"><font face="Arial" size="2"><%=Formatcurrency(ServiceFee,2)%></font></td>
          </tr>
			
          <tr style="border-bottom:2px solid black">
            <td><font face="Arial" size="2">Admin Fee</font></td> 
            <td align="right"><font face="Arial" size="2"><%=formatcurrency(cdbl(Subtotal2),2)%></font>
          </tr>
          
          <tr>
            <td><font face="Arial" size="2">Sub Total</font></td>
            <td align="right"><font face="Arial" size="2"><%=Formatcurrency(Subtotal,2)%></font></td>
          </tr>
          <tr style="border-bottom:2px solid black">
            <td><font face="Arial" size="2">Sales Tax</font></td>
            <td align="right"><font face="Arial" size="2"><%=Formatcurrency(Taxnonull,2)%></font></td>
          </tr>
          <tr>
            <td><font face="Arial" size="2"><strong>Total Due </strong>- Pay This amount -&gt;</font></td>
            <td align="right"><font face="Arial" size="2"><strong><%=Formatcurrency(totalamtnonull,2)%></strong></font></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td><img src="http://<%=request.servervariables("SERVER_NAME")%>/genergy2/invoices/MakeChartyrly.asp?genergy2=<%=trim(request("genergy2"))%>&lid=<%=leaseid%>&by=<%=billyear%>&bp=<%=billperiod%>&billid=<%=billid%>&hidedemand=<%=hidedemand%>&building=<%=building%>&unittype=<%=unittype%><%if extusage then %>&includepeaks=false&extusg=true<%else%>&includepeaks=<%=meterbreakdown%><%end if%>&calcintpeak=<%=calcintpeak%>" width="600" height="175"></td>
  </tr>
</table>
</td>
</tr>
</table>
  
  <%
  end sub
  %>