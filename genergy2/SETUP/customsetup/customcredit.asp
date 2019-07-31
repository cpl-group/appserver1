<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then%>
	<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, customsrc, action, tid, lid, id, byear, bperiod, creditid, credit_adj
pid = request("pid")
bperiod = request("bperiod")
byear = request("byear")
bldg = request("bldg")
tid = request("tid")
lid = request("lid")
if trim(request("credit_adj"))="1" then credit_adj = true else credit_adj = false
customsrc = request("customsrc")
creditid = trim(request("creditid"))

dim cnn1, rst1, sql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

'dim DBmainmodIP
'DBmainmodIP = ""

dim amount, ctype, note
action = request("action")
amount = request("amount")
ctype = request("ctype")
note = request("note")
if not(isnumeric(amount)) then
	amount = 0
end if
if trim(action)<>"" then
	amount = abs(amount)
	if credit_adj then amount = amount * -1
	if trim(action)="Save" then
		sql = "INSERT INTO Misc_Inv_Credit (BillYear, BillPeriod, leaseutilityid, Amt, credit, note) VALUES ("&BYear&", "&BPeriod&", "&lid&", "&amount&",1,'"&note&"')"
	elseif trim(action)="Update" then
		sql = "UPDATE Misc_Inv_Credit SET amt="&amount&", note = '"&note&"' WHERE id="&creditid
	elseif trim(action)="Delete" then
		sql = "delete from Misc_inv_credit where id="&creditid	
	end if
end if
if sql<>"" then 
'  response.write sql
'  response.write cnn1
  cnn1.execute sql
'  response.end
  byear=""
  bperiod=""
end if
  

	
%>
<html>
<head>
<title>Custom Credit/Adjustment</title>
<link rel="Stylesheet" href="../setup.css" type="text/css">
</head>
<script>
function amountCheck(){
	var frm = document.forms[0];
	var amount = frm.amount.value;
	frm.amount.value = amount.replace(/\-/g,"");
}
</script>
<body bgcolor="#ffffff" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0>	<%
credit_adj = true
	if creditid <> "" and request("action") = "edit" then
'		response.write "!"&creditid&"!"
'		response.end
		rst1.open "SELECT * FROM Misc_INV_Credit WHERE id='"&creditid&"'", cnn1
		if not rst1.eof then 
			amount = rst1("amt")
			note = rst1("note")
			ctype = rst1("note")
			id = rst1("id")
		else
			amount = 0
			note = ""
			ctype = ""
			id = ""
		end if
		if cdbl(amount)>0 then credit_adj=false
		rst1.close
		
		dim bldgname, portfolioname, rid
		
		if trim(bldg)<>"" then
			rst1.open "SELECT bldgname, name, region FROM buildings b join portfolio p on b.portfolioid=p.id WHERE bldgnum='"&bldg&"'", cnn1
			if not rst1.EOF then
				bldgname = rst1("bldgname")
				portfolioname = rst1("name")
				rid = rst1("region")
			end if
			rst1.close
		end if
		
		dim billingname
		if trim(bldg)<>"" then
			rst1.open "SELECT billingname FROM tblleases WHERE billingid='"&tid&"'", cnn1
			if not rst1.EOF then
				billingname = rst1("billingname")
			end if
			rst1.close
		end if		%>
		<form name="form2" method="get" action="customcredit.asp">
			<table width="100%" border="0" cellpadding="3" cellspacing="0">		<%
				if 1=0 then 'if checkgroup("clientOperations")=0 then%>
					<tr><td bgcolor="#000000">
						<table border=0 cellpadding="0" cellspacing="0">
							<tr>
								<td>
									<span class="standardheader"><a href="index.asp" target="main" class="breadcrumb" style="text-decoration:none;">
									<img src="images/aro-left-000.gif" align="left" width="13" height="13" border="0">Utility Manager Setup
									</a></span>
								</td>
								<td width="12">
									<span class="standard" style="color:#ffffff;">&nbsp;|&nbsp;</span></td>
								<td>
									<span class="standardheader"><a href="frameset.asp" target="main" class="breadcrumb" style="text-decoration:none;">
									Update Meters</a></span>
								</td>
								<td width="12">
									<span class="standard" style="color:#ffffff;">&nbsp;|&nbsp;</span></td>
								<td>
									<span class="standardheader"><a href="portfolioview.asp" target="main" class="breadcrumb" style="text-decoration:none;">
									Set Up Portfolios</a></span>
								</td>
								<td width="12">
									<span class="standard" style="color:#ffffff;">&nbsp;|&nbsp;</span></td>
								<td>
									<span class="standardheader"><a href="regionView.asp" target="main" class="breadcrumb" style="text-decoration:none;">
									Set Up Rates</a></span>
								</td>
							</tr>
						</table>
					</td></tr>		<%
				end if%>
				<tr bgcolor="#3399cc">
					<td>
						<table border=0 cellpadding="0" cellspacing="0" width="100%">
							<tr>
								<td><span class="standardheader">
									Custom Credit/Adjustment | <span style="font-weight:normal;">
									<a href="portfolioedit.asp?pid=<%=pid%>" style="color:#ffffff;"><%'=portfolioname%></a> &gt; <%=bldgname%> &gt; <%=billingname%></span>
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr bgcolor="#eeeeee">
					<td style="border-bottom:1px solid #999999">
						<table border="0" cellpadding="3" cellspacing="0">
							<tr>
								<td align="right" valign="bottom"><span class="standard">Building Name</span></td>
								<td valign="bottom"><span class="standard"><%=bldgname%></span>&nbsp;&nbsp;&nbsp;</td>
							</tr>
							<tr bgcolor="#eeeeee" class="standard">
								<td align="right">Bill Year</td>
								<td><%=byear%></td>
							</tr>
							<tr bgcolor="#eeeeee" class="standard">
								<td align="right">Bill Period</td>
								<td><%=bperiod%></td>
							</tr>
							<tr bgcolor="#eeeeee">
								<td align="right"><span class="standard">Amount</span></td>
								<td><input type="text" name="amount" value="<%=abs(clng(amount))%>" size="8" onKeyPress="amountCheck()" onBlur="amountCheck()">&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="credit_adj" value="1" <%if credit_adj then response.write "CHECKED"%>>&nbsp;Credit&nbsp;&nbsp;&nbsp;<input type="radio" name="credit_adj" value="0" <%if not(credit_adj) then response.write "CHECKED"%>>&nbsp;Adjustment</td>
							</tr>
							<tr bgcolor="#eeeeee">
								<td align="right"><span class="standard">Description</span></td>
								<td><input type="text" size=30 name="note" value="<%=note%>"></td>
							</tr>
							<!-- <tr bgcolor="#eeeeee">
								<td align="right"><span class="standard">Type</span></td>
								<td>
									<select value="ctype">
										<option>types</option>
									</select>
								</td>
							</tr> -->
							<tr bgcolor="#eeeeee">
								<td></td>
								<td>
									<%if trim(id)<>"" then%>
										<input type="submit" name="action" value="Update" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
										<input type="submit" name="action" value="Delete" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
									<%else%>
										<input type="submit" name="action" value="Save" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
										<input type="submit" name="action" value="Cancel" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;">
									<%end if%>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<input type="hidden" name="pid" value="<%=pid%>">
			<input type="hidden" name="bldg" value="<%=bldg%>">
			<input type="hidden" name="customsrc" value="<%=customsrc%>">
			<input type="hidden" name="tid" value="<%=tid%>">
			<input type="hidden" name="lid" value="<%=lid%>">
			<input type="hidden" name="byear" value="<%=byear%>">
			<input type="hidden" name="bperiod" value="<%=bperiod%>">
			<input type="hidden" name="creditid" value="<%=creditid%>">
		</form>
	
<%elseif request("action")="choosecredit" then%>

	<table width="100%" border="0" cellpadding="3" cellspacing="1">
		<tr bgcolor="#dddddd">
			<td width="75%"><b>Credit Description</b></td>
			<td width="25%"><b>Amount</b></td>
		</tr>
		<%
		rst1.open "SELECT * FROM misc_inv_credit WHERE credit=1 and leaseutilityid="&lid&" and billyear="&byear&" and billperiod="&bperiod, cnn1
		if rst1.eof then %>
			<tr valign="top"><td colspan="2">There are no credits for this billperiod.</td></tr><%
		else
			do until rst1.eof	
				dim link
				link = "customcredit.asp?action=edit&pid="&pid&"&bldg="&bldg&"&tid="&tid&"&lid="&lid&"&byear="&rst1("billyear")&"&bperiod="&_
					rst1("billperiod")&"&creditid="&rst1("id")%>
				<tr bgcolor="#ffffff" valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" 
					onclick="document.location='<%=link%>'">
					<td><%=rst1("note")%></td>
					<td><%if rst1("credit")=1 then response.write clng(rst1("amt"))*-1 else response.write rst1("amt")%></td>
				</tr>
				<%
				rst1.movenext
			loop
		end if
		dim neLink
		neLink = "customcredit.asp?pid="&pid&"&bldg="&bldg&"&tid="&tid&"&lid="&lid&"&byear="&byear&"&bperiod="&bperiod&"&creditid=0&action=edit"
		%>
		<tr>
			<td colspan=2 onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="document.location='<%=neLink%>'">
				New Entry
			</td>
		</tr>
	</table><%
	
else%>

	<table width="100%" border="0" cellpadding="3" cellspacing="1">
		<tr bgcolor="#dddddd">
			<td width="18%"><b>Bill&nbsp;Year</b></td>
			<td width="23%"><b>Bill&nbsp;Period</b></td>
			<td width="17%"><b>Start&nbsp;Date</b></td>
			<td width="28%"><b>End&nbsp;Date</b></td>
			<td width="28%" nowrap><b>Total&nbsp;Amount</b></td>
		</tr>
		<%
		dim someSql
		someSql = "SELECT b.billyear, b.billperiod, b.datestart, b.dateend, sum(m.amt) as credit FROM billyrperiod b LEFT JOIN misc_inv_credit m ON m.billperiod=b.billperiod and m.credit=1 and m.billyear=b.billyear and m.leaseutilityid="&lid&" WHERE utility=(SELECT utility FROM tblleasesutilityprices WHERE leaseutilityid="&lid&") and b.bldgnum='"&bldg&"' group by b.billyear, b.billperiod,b.dateStart,b.dateend ORDER BY b.billyear desc, b.billperiod desc"
'		response.write someSql
		rst1.open someSql, cnn1
		
		if rst1.eof then response.write "<tr valign=""top"" colspan=""4""><td>There are no bill periods setup.</td></tr>"
		do until rst1.eof%>
			<tr bgcolor="#ffffff" valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="document.location='<%="customcredit.asp?pid="&pid&"&bldg="&bldg&"&tid="&tid&"&lid="&lid&"&byear="&rst1("billyear")&"&bperiod="&rst1("billperiod")&"&action=choosecredit"%>'">
				<td><%=rst1("billyear")%></td>
				<td><%=rst1("billperiod")%></td>
				<td><%=rst1("datestart")%></td>
				<td><%=rst1("dateend")%></td>
				<td><%=rst1("credit")%></td>
			</tr>
			<%
			rst1.movenext
		loop
		%>
	</table>
<%end if%>
</body>
</html>