<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim leaseid, ypid, utilityid, building, byear,bperiod
dim currpage, totpages, maxPageCount, page1Flag
leaseid = trim(Request("l"))
ypid = trim(request("y"))
utilityid = trim(request("utilityid"))
building = trim(request("building"))
currpage = trim(request("currpage"))
totpages = trim(request("totpages"))
byear = trim(request("byear"))
bperiod = trim(request("bperiod"))
maxPageCount = trim(request("maxPageCount"))
page1Flag = 0

if building="BATH" and utilityid=2 and byear=2010 and bperiod=3 then
	maxPageCount = 2
end if

if (building="LGA" or building="JFK" or building="PABT") and utilityid=2 and byear=2010 and bperiod=4 and maxPageCount=1 then
	maxPageCount = 2
	page1Flag = 1
end if

Dim invoiceIndex
invoiceIndex = Session("invoiceCounter")
if not (invoiceIndex <> "") then
	invoiceIndex = 1
end if

Dim isPDF
isPDF = false
'if ((request.servervariables("HTTP_REFERER")="Webster://Internal/315" and isempty(session("xmlUserObj")))  ) then 'this is for pdf sessions
 loadNewXML("activepdf")
loadIps(0)
isPDF = true
'end if

dim cnn1, rst4,rst5
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst4 = Server.CreateObject("ADODB.recordset")
Set rst5 = Server.CreateObject("ADODB.recordset")
'if trim(getbuildingIP(building))="0" then response.redirect "/eri_th/pdfmaker/genergyInvoice.asp?"&request.servervariables("QUERY_STRING") else cnn1.Open getLocalConnect(building)
cnn1.Open getLocalConnect(building)

%>
<html>
<body>
<%
dim condition
if maxPageCount <> "" then
	condition = currpage mod maxPageCount = 0
else
	condition = currpage mod totpages = 0
end if

if condition then
	if ypid <> "" and leaseid <> "" and utilityid <> "" then
		footer ypid,leaseid, utilityid
	elseif building <> "" then
		dim sql
		sql = "SELECT b.id as billid, b.leaseutilityid, ypid, b.utility, extusg FROM tblbillbyperiod b WHERE reject=0 and bldgnum='"&building&"' and billyear="&byear&" and billperiod="&bperiod
		if isnumeric(utilityid) then 
			sql = sql & " and utility="&utilityid 
		end if
		sql = sql & "  ORDER BY TenantName"
			
		'response.Write(sql & " <p> ")
		dim blgdrset
		Set blgdrset = Server.CreateObject("ADODB.recordset")
		blgdrset.open sql, cnn1
		Dim currIndex
		currIndex = 0
		do until blgdrset.eof
			currIndex = currIndex + 1
		
			leaseid = trim(blgdrset("leaseutilityid"))
			ypid = trim(blgdrset("ypid"))
			utilityid = trim(blgdrset("utility"))
			
			Dim metercountInfo, billid
			Set metercountInfo = Server.CreateObject("ADODB.recordset")
			billid = blgdrset("billid")
			metercountInfo.open "select count(*) as metercount from tblmetersbyperiod tm,buildings b,meters m where tm.bldgnum =b.bldgnum and tm.meterid=m.meterid and b.bldgnum = m.bldgnum and bill_id="&billid, cnn1
			
			dim tempMaxPageCount
			tempMaxPageCount = metercountInfo("metercount") \ 40 + 1
			
			if metercountInfo("metercount") > 5 then
				tempMaxPageCount = tempMaxPageCount + 1
			end if
			
			metercountInfo.Close()
			if building="BATH" and utilityid=2 and byear=2010 and bperiod=3 then
				tempMaxPageCount = 2
			end if
			if (building="LGA" or building="JFK" or building="PABT") and utilityid=2 and byear=2010 and bperiod=4 then
				if tempMaxPageCount=1 and page1Flag=1 then
					tempMaxPageCount = 2
					maxPageCount = 2
				elseif tempMaxPageCount=2 and page1Flag=1 then
					maxPageCount = 1
				end if
		   end if
			
			if currIndex = invoiceIndex and CInt(tempMaxPageCount) <> CInt(maxPageCount) then
				invoiceIndex = invoiceIndex + 1
			end if

			if currIndex = invoiceIndex then
				footer ypid,leaseid, utilityid
			end if 
			
			blgdrset.movenext
		loop
		blgdrset.Close()
	end if
	Session("invoiceCounter") = invoiceIndex + 1
end if
%>
</body>
</html>
<%
sub footer(ypid,leaseid,utilityid)
dim footerBillid,sql
Dim sInvoiceNo

sql = "SELECT bp.id as billid " & _
	  "FROM tblbillbyPeriod bp " & _
	  "WHERE reject=0 and bp.ypid="&ypid&" and bp.leaseutilityid="&leaseid 
			
	rst4.open sql, cnn1
		if not rst4.eof then
			footerBillid = rst4("billid")
		end if
	rst4.Close
	
	if utilityid = 2 then
		sql = "SELECT bp.id as billid , RIGHT(REPLACE(RTRIM(SPACE(5) + STR(BN.InvoiceSeqNo)),' ','0'),5) AS InvoiceSeqNo, BN.BillType " & _
			  "FROM tblbillbyPeriod bp, tblPAInvoiceBillNumbers BN  " & _
			  "WHERE reject=0 and bp.ypid="&ypid&" and bp.leaseutilityid="&leaseid & _
					" and BN.billid = bp.id"
	elseif utilityid=3 or utilityid=10 then
			sql = "SELECT bp.id as billid , RIGHT(REPLACE(RTRIM(SPACE(5) + STR(BN.InvoiceSeqNo)),' ','0'),5) AS InvoiceSeqNo, BN.BillType " & _
			  "FROM tblbillbyPeriod bp, tblPAWaterBillNumbers BN  " & _
			  "WHERE reject=0 and bp.ypid="&ypid&" and bp.leaseutilityid="&leaseid & _
					" and BN.billid = bp.id"
	'3/31/2009 N.Ambo added new chiller utilities				
	'4/28/2009 N.Ambo modfied new chiller utilities
	elseif utilityid =18  or utilityid=19  or utilityid=20 then
			sql = "SELECT bp.id as billid, RIGHT(REPLACE(RTRIM(SPACE(5) + STR(BN.InvoiceSeqNo)),' ','0'),5) AS InvoiceSeqNo, BN.BillType " & _
			    "FROM tblbillbyPeriod bp, tblPAChWaterBillNumbers BN " & _
				" WHERE bp.ypid="&ypid&" and bp.leaseutilityid="&leaseid & _
					"and BN.billid = bp.id"
			
	end if	  

	
	rst4.open sql, cnn1
	sInvoiceNo = "0000000"
	if not rst4.eof then
		If not isNull(rst4("BillType")) Then 	
			sInvoiceNo = rst4("BillType") &  rst4("InvoiceSeqNo")
		Else
			sInvoiceNo = "0000000" ' Indicates an Error While Generating The Invoice Number
		End If

	end if

	rst4.Close()

select case utilityid
case 3, 10,6,1,4,19,18,20  '3/30/2009 N.Ambo added chiller utilities
	select case utilityid
	  case 6,1,4
		'usagedivisor = 1
	case else
		'usagedivisor = 100
     end select 
'water and stuff
sql = "SELECT isnull(r.adjustment,0) as adjustment ,isnull(r.FuelADJ,0) as usefuelADJ,ContactPhone,ContactName,b.bldgname,isnull(b.btzip,'') as billzip, b.portfolioid, isnull(b.btstrt,'') as billto, r.addonfee as myaddonfee, isnull(r.energydetail,'0') as energydetail, isnull(r.demanddetail,'0') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress,isnull(r.adminfee,0) as adminfee,isnull(r.servicefee,0) as servicefeenonull, isnull(r.fueladj,0) as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(r.adjustment, 0) as adjustmentnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.[id] as rtid, datediff(day, datestart,dateend)+1 as days, (select case count(distinct isnull(addonfee,0)) when 0 then 0 else 1 end as aoncnt from tblmetersbyperiod where bill_id=r.id group by leaseutilityid,ypid,bill_id) as showaddonfee, r.invoice_note as invoiceNote, isnull(rate_servicefee_dollar,0) as rateservicefee_dollar, r.*, l.onlinebill,l.corpStreet,l.corpCity,l.corpState,l.corpZip,l.corpCountry,pa.acctnumber,pa.leasenumber,pa.seqnumber  FROM tblbillbyperiod r, tblleases l, buildings b,ratetypes rt,custom_PABT pa WHERE pa.acctnumber=l.billingid and r.ratetenant=rt.id AND b.bldgnum=r.bldgnum and l.billingid = (SELECT billingid FROM tblleasesutilityprices lup WHERE lup.leaseutilityid=r.leaseutilityid) and r.id="&footerBillid
   
   case else

'sql electricty
sql = "SELECT isnull(r.adjustment,0) as adjustment ,isnull(r.FuelADJ,0) as usefuelADJ,ContactPhone,ContactName,b.bldgname,isnull(b.btzip,'') as billzip, b.portfolioid,isnull(r.btstrt,'') as billto, isnull(r.energydetail,'') as energydetail, isnull(r.demanddetail,'') as demanddetail, r.utility as unittype, isnull(Totalamt,0) as Totalamtnonull, isnull(tax,0) as taxnonull, isnull(energy,0) as energynonull, isnull(r.tstrt,'') as billingaddress, r.fueladj as fadj, isnull(demand,0) as demandnonull, isnull(credit,0) as creditnonull, isnull(r.adjustment, 0) as adjustmentnonull, isnull(credit,0) as credit, rt.[type] as rt,rt.id as rtid, datediff(day, datestart,dateend)+1 as days, case rt.[type] when 'AVG Cost 1' then round(avgkwh,6) when 'AVG COST 2' then round(unitcostkwh,6) else ' ' end as akwhdisplay, case rt.[type] when 'AVG COST 2' then round(isnull(tunitcostkw,0),6) else ' ' end as akwdisplay, case when Totalkw=0 then 0 else ((Totalkwh/Totalkw)/(datediff(day, ypiddatestart,ypiddateend)+1)*24) end as loadfactor, isnull(r.adminfee,0) as adminfee, isnull(r.adminfeedollar,0) as adminfeedollar, r.billperiod, r.billyear, r.datestart, r.dateend, isnull(r.servicefee,0) as servicefeenonull,r.addonfee as myaddonfee, r.unit_credit, isnull(r.subTotal,0) as subTotal, r.tenantname, lup.calcintpeak, r.*, l.onlinebill,l.corpStreet,l.corpCity,l.corpState,l.corpZip,l.corpCountry,pa.acctnumber,pa.leasenumber,pa.seqnumber FROM rpt_bill_summary r, tblleases l, buildings b, dbo.ratetypes rt, tblleasesutilityprices lup,custom_PABT pa WHERE pa.acctnumber=l.billingid and r.[type]=rt.id AND b.bldgnum=r.bldgnum and lup.leaseutilityid=r.leaseutilityid and l.billingid=lup.billingid and r.billid="&footerBillid
end select

dim ContactName, ContactPhone,pid
'response.write sql
'response.end
rst5.open sql, cnn1, 2
 if not rst5.eof then
dim acctnumber,seqnumber,leasenumber,corpStreet,corpCity,corpState,corpZip,FuelAdj
acctnumber=rst5("acctnumber")
seqnumber=rst5("seqnumber")
leasenumber=rst5("leasenumber")
corpStreet=rst5("corpStreet")
pid = rst5("portfolioid")

corpCity=rst5("corpCity")&","
corpState=rst5("corpState")&","
corpZip=rst5("corpZip")
ContactName=rst5("ContactName")
ContactPhone=rst5("ContactPhone")
FuelAdj =rst5("Fuel")
if FuelAdj = "" or isnull(FuelAdj) then
FuelAdj = 0
end if


if corpStreet="" then
corpStreet=replace(rst5("tstrt"),vbNewLine,"<br>")
corpCity=rst5("tcity")&","&  rst5("tstate")&","&  rst5("tzip")
corpState=""
corpZip=""


end if
dim rst555,utilityname
Set rst555 = Server.CreateObject("ADODB.recordset")
rst555.open "SELECT utility FROM tblutility WHERE utilityid="&utilityid, getConnect(pid,building,"billing")
if not rst555.eof then
	utilityname = rst555("utility")
end if
rst555.close


'middle rst5("subtotal"),rst5("FuelAdj"),rst5("servicefee"),rst5("Adminfee"),rst5("subtotal"),rst5("Taxnonull"),rst5("totalamtnonull"),rst5("rt")
%>
<table border="0">
<tr>
    <td align="center">- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - </td>
</tr>
<tr>
    <td height="5" align="center"><font size="3"><b>Please Detach This Section And Return With Your Payment</b></font></td>
  </tr>
  <tr>
    <td height="50"><div class="style11" style="border:4px double black">
      CHARGES COVERED BY THIS INVOICE ARE  DUE AND PAYABLE NOW.<BR />
        LATE CHARGES MAY BE APPLIED TO THE UNPAID PORTION OF YOUR ACCOUNT.<BR />
        IF YOU HAVE ANY QUESTIONS ABOUT HOW YOUR BILL WAS PREPARED, PLEASE CALL GENERGY AT 212-664-7600 <%if ContactName <> "" then%> FOR ANY OTHER
        QUESTIONS, CALL <%=ContactName%> AT <%=ContactPhone%><%end if%>
    </div></td>
  </tr>
  <tr>
    <td height="30"><table width="100%" height="30" border="0" ID="Table11">
      <tr>
        <td><font face="Arial" size="2"><b><%=utilityname%> Bill</b></font></td>
        <td><font face="Arial" size="2"><b>&nbsp;</b></font></td>
        <td><font face="Arial" size="2"><b>&nbsp;</b></font></td>
        <td><font face="Arial" size="2"><b>&nbsp;</b></font></td>
        <td><font face="Arial" size="2"><b>&nbsp;</b></font></td>
      </tr>
      <tr>
        <td><font face="Arial" size="2"><b>Account No.</b></font></td>
        <td><font face="Arial" size="2"><b>Sequence No.</b></font></td>
        <td><font face="Arial" size="2"><b>Lease No.</b></font></td>
        <td><font face="Arial" size="2"><b>Invoice No.</b></font></td>
        <td><font face="Arial" size="2"><b>Invoice Date</b></font></td>
      </tr>
      <tr>
	  



        <td height="20"><font face="Arial" size="2"><%=rst5("tenantnum")%></font></td>
        <td height="20"><font face="Arial" size="2"><%=seqnumber%></font></td>
        <td height="20"><font face="Arial" size="2"><%=leasenumber%></font></td>
        <td height="20"><font face="Arial" size="2"><%=sInvoiceNo%></font></td>
        <td height="20"><font face="Arial" size="2"><%if isdate(rst5("postdate")) then response.write formatdatetime(rst5("postdate"),2)%></font></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="100"><table width="100%" height="100" border="0" ID="Table12">
      <tr>
        <td valign="top" width="65%"><table width="100%" height="100" border="0" cellpadding="0" cellspacing="0" ID="Table13">
            <tr>
              <td height="20"><font face="Arial" size="2"><b>Name:</b></font></td>
              <td height="20"><font face="Arial" size="2"><%=rst5("tenantname")%></font></td>
            </tr>
            <tr>
              <td height="20" valign="top"><font face="Arial" size="2"><b>Corporate Address:</b></font></td>
              <td height="20"><font face="Arial" size="2"><%=corpStreet%></font></td>
            </tr>
            <tr>
              <td height="20">&nbsp;</td>
              <td height="20"><font face="Arial" size="2"><%=corpCity&corpState&corpZip%></font></td><!--41st Floor-->
            </tr>
            <tr>
              <td height="20">&nbsp;</td>
              <td height="20" class="style9"> </td>
            </tr>
            <tr>
              <td height="20">&nbsp;</td>
              <td height="20">&nbsp;</td>
            </tr>
        </table></td>
        <td valign="top" width="35%">
			<table height="100" border="0" style="border:2px solid black; background-color:#CCCCCC" ID="Table16">
          <tr>
            <td><font face="Arial" size="2"><b>Make Check Payable To:</b></font></td>
            <td align="left"><font face="Arial" size="2"><b>The Port Authority of NY &amp; NJ</b></font></td>
          </tr>
          <tr>
            <td colspan="2" align="center"><div align="left"><font color="#ff0000" face="Arial" size="2">For proper credit, you must enter your invoice number<br /><font color="#000000"><%=sInvoiceNo%></font> on your check.</font></div></td>
            </tr>
          <tr>
            <td valign="top"><font face="Arial" size="2"><b>Mail Payment To:</b></font></td>
            <td valign="top"><table width="100%" border="0" ID="Table17">
              <tr>
                <td><font face="Arial" size="2"><b>Port Authority of NY &amp; NJ</b></font></td>
              </tr>
              <tr>
                <td><font face="Arial" size="2"><b>P.O. Box 95000-1517</b></font></td>
              </tr>
              <tr>
                <td><font face="Arial" size="2"><b>Philadelphia, PA 19195-1517</b></font></td>
              </tr>
            </table>
          </td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="100"><table width="100%" height="100" border="0" ID="Table15">
      <tr>
        <td width="80%">
<table width="100%"  height="100" border="0" cellpadding="0" cellspacing="0" ID="Table14">
            <tr>
              <td height="20">&nbsp;</td>
              <td height="20" class="style9">&nbsp;</td>
            </tr>
            <tr>
              <td height="20"><font face="Arial" size="2"><b>Billing Address:</b></font></td>
              <td height="20"><font face="Arial" size="2"><%=rst5("tenantname")%></font></td>
            </tr>
            <tr>
              <td height="20">&nbsp;</td>
              <td height="20"><font face="Arial" size="2"><%=replace(rst5("tstrt"),vbNewLine,"<br>")%><br></font></td>
            </tr>
            <tr>
              <td height="20">&nbsp;</td>
              <td height="20"><font face="Arial" size="2"><%=rst5("tcity")&","&  rst5("tstate")&","&  rst5("tzip")%></font></td>
            </tr>
            <tr>
              <td height="20"><font face="Arial" size="2"><b><!--Attn to:--></b></font></td>
              <td height="20"><font face="Arial" size="2"><!--Bill Payyer --></font></td>
            </tr>
        </table>
                </td>
          </tr>
        </table></td>
        <td width="20%" valign="top"><table width="100%" height="100" border="0" ID="Table18">
        <!--  <tr>
            <td><font face="Arial" size="2">SubTotal</font></td>
           
		  <td align="right"><font face="Arial" size="2"><'%=FormatCurrency(rst5("subtotal"),2)%></font></td>
          </tr>-->
          <tr>
            <td><font face="Arial" size="2">Fuel Adj</font></td>
           <td align="right"><font face="Arial" size="2"><%=FormatCurrency(FuelAdj,2)%></font></td>
		   <!-- <td align="right"><font face="Arial" size="2"><'%=FormatCurrency(rst5("FuelAdj"),2)%></font></td>usefuelADJ-->
          </tr>
          <tr>
            <td><font face="Arial" size="2">Adjustment</font></td>
            <td align="right"><font face="Arial" size="2"><%=FormatCurrency(rst5("Adjustment"),2)%></font></td>
          </tr>
          <tr>
            <td><font face="Arial" size="2">Service Fee </font></td>
            <td align="right"><font face="Arial" size="2"><%=FormatCurrency(rst5("serviceFeenonull"),2)%></font></td>
          </tr>
          <tr style="border-bottom:2px solid black">
            <td><font face="Arial" size="2">Admin Fee</font></td>
            <td align="right"><font face="Arial" size="2"><%=FormatCurrency(rst5("Adminfeedollar"),2)%></font></td>
          </tr> 
          <tr>
            <td><font face="Arial" size="2">Sub Total </font></td>
            <td align="right"><font face="Arial" size="2"><%=FormatCurrency(rst5("subtotal"),2)%></font></td>
          </tr>
          <tr style="border-bottom:2px solid black">
            <td><font face="Arial" size="2">Sales Tax </font></td>
            <td align="right"><font face="Arial" size="2"><%=FormatCurrency(rst5("taxnonull"),2)%></font></td>
          </tr>
          <tr>
            <%
			'dim ttlamt
			' ttlamt = FormatCurrency(rst5("Totalamtnonull"),2)  + FormatCurrency(rst5("taxnonull"),2)
			'response.write ttlamt
			%>
			<td><font face="Arial" size="2"><b><font size="3">Total Due</font> - Pay This amount -&gt;</b></font></td>
            <td align="right"><font face="Arial" size="2"><b><font size="3"><%=FormatCurrency(rst5("Totalamtnonull"),2)%></font></b></font></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>
<%
end if
rst5.close
end sub
%>