<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!-- #include virtual="/genergy2_Intranet/itservices/ttracker/TTServices.inc" -->
<%


'***************************************************************
'12/28/2007 N.Ambo added default values of 0 to fields WattsperSqFtLow, WattsperSqFtHigh
'***************************************************************

dim pid, bldg, tid, edit,ticketcount, masterticketid,opentickets,criticalopentickets, totaltickets
pid = request("pid")
bldg = request("bldg")
tid = request("tid")
edit = request("edit")

dim cnn1, rst1, strsql, rst2, rst3
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.createobject("ADODB.recordset")
set rst3 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)
'response.write( getLocalConnect(bldg) )
'response.end()

dim bldgname, portfolioname, intermcharges
dim tenantnum, flr, sqft, taxexempt, billingname, leaseexpired, interm, startdate, tName, tStrt, tCity, tState, tZip, customsrc, lmepExempt, onlinebill, ibsexempt, bsexempt, accounttype
dim corpStreet, corpCity, corpState, corpZip, tCountry, corpCountry

dim leasenum, sequencenum

dim AcctCode
dim WattsperSqFtLow, WattsperSqFtHigh

'12/28/2007 N.Ambo added these two lines to create default values for these fields
WattsperSqFtLow = 0.0
WattsperSqFtHigh = 0.0

'01/25/2008 Tarun : Tenant Move In date related variables
dim rst4, TenantMoveInDate, LeaseExpirydate, TenantEmail
set rst4 = server.createobject("ADODB.recordset")

if trim(bldg)<>"" then
  rst1.open "SELECT bldgname, b.strt, b.city, b.state, b.zip, name FROM buildings b join portfolio p on b.portfolioid=p.id WHERE bldgnum='"&bldg&"'", cnn1
'	rst1.Open "SELECT bldgname FROM buildings WHERE bldgnum='"&bldg&"'", cnn1
	if not rst1.EOF then
		bldgname = rst1("bldgname")
		portfolioname = rst1("name")
		tstrt = rst1("strt")
		tcity = rst1("city")
		tstate = rst1("state")
		tzip = rst1("zip")
	end if
	rst1.close
end if

if trim(tid)<>"" then
	rst1.Open "SELECT * FROM tblleases where billingid='"&tid&"'", cnn1
'response.write "SELECT * FROM tblleases WHERE billingid='"&tid&"'"
'response.end
	if not rst1.EOF then
		tenantnum = rst1("tenantnum")
		flr = rst1("flr")
		sqft = rst1("sqft")
		taxexempt = rst1("taxexempt")
		billingname = rst1("billingname")
		leaseexpired = rst1("leaseexpired")
		interm = rst1("interm")
		tName = rst1("tName")
		tStrt = rst1("tStrt")
		tCity = rst1("tCity")
		tState = rst1("tState")
		tZip = rst1("tZip")
 
		tCountry = rst1("tCountry")
		corpStreet = rst1("corpStreet")
		corpCity = rst1("corpCity")
		corpState = rst1("corpState")
		corpZip = rst1("corpZip")
		corpCountry = rst1("corpCountry")
		customsrc = rst1("customsrc")
		ibsexempt = rst1("ibsexempt")
		bsexempt  = rst1("billsummaryexempt")
		if trim(rst1("intermcharges"))<>"" then 
			intermcharges = rst1("intermcharges")
		else
		    intermcharges = "0"
	    end if
		startdate = rst1("startdate")
		bldg = rst1("bldgnum")
	    lmepExempt = rst1("lmepExempt")
    	onlinebill = rst1("onlinebill")
		accounttype = rst1("accounttype")
		if isnumeric(accounttype) then accounttype = cint(accounttype) else accounttype = 2
		
		'marko added: lease/sequence number lookup (port authority specific)
		if pid = 108 then
			rst3.Open "SELECT * FROM custom_PABT where acctnumber='"& tid &"'", cnn1
			if not rst3.EOF then
				leasenum = rst3("leasenumber")
				sequencenum = rst3("seqnumber")	
			end if
			rst3.Close()
		end if

		if pid = 45 then
			rst3.Open "SELECT * FROM custom_SL where acctnumber='"& tid &"'", cnn1
			if not rst3.EOF then
				AcctCode = rst3("AcctCode")
			end if
			rst3.Close()
		end if	
		' Added by Tarun : 01/25/2008
		LeaseExpirydate = rst1("DateExpired")
		rst4.Open "SELECT BillingId, MoveInDate, TenantEmail FROM tblTenantExtDetails WHERE BillingId = '" & tid & "'", cnn1
		if not rst4.EOF then
			TenantMoveInDate = rst4("MoveInDate")
			if not isNull(rst4("TenantEmail")) then 
				TenantEmail = rst4("TenantEmail") 
			End If
		end if
		rst4.Close 
		
		
	end if
	rst1.close
else
	startdate = date()
	accounttype = 2
	intermcharges = "0"
	
end if

'if trim(tid) <> "" and bldg <> "" then 
'	dim ticket
'	set ticket = New tickets
'	ticket.Label="Account"
'	ticket.Note = "Master Ticket for Account Billing ID "&split(getBuildingIP(bldg),"\")(1)&"-"&tenantnum
'	ticket.ccuid  = "rbdept"
'	ticket.client = 1
'	if tid<>"0" then ticket.findtickets "tid", split(getBuildingIP(bldg),"\")(1)&"-"&tid
'end if 
%>
<html>
<head>
<title>Building View</title>
<script language="JavaScript">
function openCustomWin(clink, cname, cspec)
{	cWin = window.open(clink, cname, cspec)
  cWin.focus();
}

function leaseUtilityEdit(lid)
{	document.location.href = 'leaseutilityedit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid='+lid
}
function meterEdit(meterid,lid)
{	
  //if ((parent.name=="main") && (parent.frames.length >= 2)) { 
  //  document.location.href = "contentfrm.asp?action=meteredit&pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid="+lid+"&meterid="+meterid
  //} else {
    document.location.href = "contentfrm.asp?action=meteredit&pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid="+lid+"&meterid="+meterid
  //}
}
function meterAdd(meterid,lid)
{	
  //if ((parent.name=="main") && (parent.frames.length >= 2)) {
    document.location.href = "meteredit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid="+lid+"&meterid="+meterid
//    document.location.href = "contentfrm.asp?action=meteradd&pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid="+lid+"&meterid="+meterid
  //} else {
  //  document.location.href = "frameset.asp?action=meteradd&pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid="+lid+"&meterid="+meterid
  //}
}

function toggleInfoDisplay(util){
  utildiv = util + "info";
  utilbutton = util + "button"
  if (document.all[utildiv].style.display == "none") {
    document.all[utildiv].style.display = "inline";
    document.all[utilbutton].value = "Hide Info";
    //document.all[utilbutton].style.border = "1px inset #ffffff"
  } else {
    document.all[utildiv].style.display = "none";
    document.all[utilbutton].value = "Show Info";
    //document.all[utilbutton].style.border = "1px outset #ffffff"
  }
}

//visual feedback functions for img buttons
function buttonOver(obj,clr){
  if (arguments.length == 1) { clr = "#336699"; }
  obj.style.border = "1px solid " + clr;
}

function buttonDn(obj,clr){
  if (arguments.length == 1) { clr = "#000000"; }
  obj.style.border = "1px solid " + clr;
}

function buttonOut(obj,clr){
  if (arguments.length == 1) { clr = "#eeeeee"; }
  obj.style.border = "1px solid " + clr;
}

function toggleDisplayOnCheck(check,elemId)
{
	if(check.checked)
		document.getElementById(elemId).style.display = 'none'
	else
		document.getElementById(elemId).style.display = 'block'
}

</script>
<script src="/quickhelp/quickhelp.js" type="text/javascript" language="Javascript1.2"></script>
<style type="text/css">
.mgmtlink:hover { color:#3399cc; }
.custlink:hover { color:#339999; }
a.custlink { color:#006666; }
</style>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body bgcolor="#ffffff">
<form name="form2" method="post" action="tenantsave.asp">
  <table width="100%" border="0" cellpadding="3" cellspacing="0">
    <tr bgcolor="#6699cc"> 
      <td height="26"> <table border=0 cellpadding="0" cellspacing="0" width="100%">
          <tr> 
            <td> <span class="standardheader"> 
              <%if trim(tid)<>"" then%>
					<%if (edit) then %>
					Update
					<% else %>
					View
					<% end if %>
					Account | <span style="font-weight:normal;"><nobr><%=portfolioname%></nobr> 
					&gt; <nobr><%=bldgname%></nobr> 
					&gt; <nobr><%=billingname%></nobr></span> 
              <%else%>
					Add New Account to <%=bldgname%> 
              <%end if%>
              </span> </td>
          </tr>
        </table></td>
      <td align="right">
        <button id="qmark2" onclick="openCustomWin('help.asp?page=tenantedit','Help','width=400,height=550,scrollbars=1')" style="cursor:help;color:#339933;text-decoration:none;height:20px;background-color:#eeeeee;border:1px outset;color:009900;margin-left:4px;" class="standard">(<b>?</b>) 
        Quick Help</button></td>
    </tr>
    <tr bgcolor="#eeeeee" valign="top"> 
      <td colspan="2"> 
        <% if (edit) or trim(tid)="" then %>
        <table border="0" cellpadding="0" cellspacing="0" width="100%">
          <tr>
            <td valign="top"> <table border="0" cellpadding="3" cellspacing="0" width="100%">
                <tr>
                  <td align="right">Account Number</td>
                  <td><input type="text" name="tenantnum" maxlength="12" value="<%=tenantnum%>"></td>
                </tr>
                <% if pid = 108 then %>
                <tr>
                  <td align="right">Lease Number</td>
                  <td><input type="text" name="leasenum" maxlength="12" value="<%=leasenum%>" ID="Text6"></td>
                </tr>
                <tr>
                  <td align="right">Sequence Number</td>
                  <td><input type="text" name="seqnum" maxlength="12" value="<%=sequencenum%>" ID="Text7"></td>
                </tr>
                <% end if %>
                <% if pid = 45 then %>
                <tr>
                  <td align="right">Account Code</td>
                  <td><input type="text" name="AcctCode" maxlength="20" value="<%=leasenum%>" ID="AcctCode"></td>
                </tr>
                <% end if %>                
                <tr>
                  <td align="right">Lease Start Date</td>
                  <td><input type="text" name="startdate" value="<%=startdate%>"</td>
                </tr>
                <tr>
                <tr>
                  <td align="right">Tenant MoveIn Date</td>
                  <td><input type="text" name="TenantMoveIndate" value="<%=TenantMoveIndate%>"</td>
                </tr>
                <tr>
                  <td align="right">Billing Name</td>
                  <td><input type="text" name="billingname" value="<%=billingname%>" onchange="tname.value = this.value"></td>
                </tr>
                <tr>
                  <td align="right">Account Name</td>
                  <td><input type="text" name="tname" value="<%=tname%>"></td>
                </tr>
                <tr>
                  <td align="right" valign="top">Billing Address</td>
                  <td>
					<textarea cols="25" rows="2" name="tstrt" wrap="off"><%=tstrt%></textarea>
					<br/>City: <input type="text" name="tcity" value="<%=tcity%>"/>
					<br/>State: <input type="text" name="tstate" value="<%=tstate%>" size="10" maxlength="2" ID="Text1">
                    Postal Code (Zip): <input type="text" name="tzip" value="<%=tzip%>" size="10" ID="Text2">
                    <br/>Country: <input type="text" name="tcountry" value="<%=tCountry%>" size="10" ID="Text8">
                  </td>
                </tr>
                <tr>
                  <td align="right" valign="top">Corporate Address</td>
                  <td>
					<input type="checkbox" name="corpAddressSameAsBillingCheck" onclick="toggleDisplayOnCheck(this,'corpAddressBlock')" <% if not (corpStreet <> "") then %>checked<%end if%>/>Same as billing address
					<div id="corpAddressBlock" <% if not (corpStreet <> "") then %>style="display:none"<%end if%>>
					<br/>
					<textarea cols="25" rows="2" name="corpStreet" wrap="off" ID="Textarea1"><%=corpStreet%></textarea>
					<br/>City: <input type="text" name="corpCity" value="<%=corpCity%>" ID="Text3"/>
					<br/>State: <input type="text" name="corpState" value="<%=corpState%>" size="10" maxlength="2" ID="Text4">
                    Postal Code (Zip): <input type="text" name="corpZip" value="<%=corpZip%>" size="10" ID="Text5">
                    <br/>Country: <input type="text" name="corpCountry" value="<%=corpCountry%>" size="10" ID="Text9">
                    </div>
                  </td>
                </tr>
              </table></td>
            <td valign="top"> <table border="0" cellpadding="3" cellspacing="0" width="100%">
                <tr>
                  <td align="right">Account Type</td>
                  <td> <select name="accounttype">
                      <%
						strsql = "SELECT * FROM account_type ORDER by description"
						rst1.open strsql, getConnect(0,0,"dbCore")
						do until rst1.eof
							%>
                      <option value="<%=rst1("id")%>" <%if accounttype = cint(rst1("id")) then response.write "SELECTED"%>><%=rst1("description")%></option>
                      <%
							rst1.movenext
						loop
						rst1.close
						%>
                    </select> </td>
                </tr>
                <tr>
                  <td align="right">Floor</td>
                  <td><input type="text" name="flr" value="<%=flr%>"></td>
                </tr>
                <tr>
                  <td align="right">SQFT</td>
                  <td><input type="text" name="sqft" value="<%=sqft%>"></td>
                </tr>
                <tr>
                  <td align="right">Tax Exempt</td>
                  <td><input type="checkbox" value="1" name="taxexempt" <%if taxexempt="True" then Response.Write "CHECKED"%>></td>
                </tr>
                <tr>
                  <td align="right">Interim Charges</td>
                  <td><input type="checkbox" value="1" name="interm" <%if interm="True" then Response.Write "CHECKED"%>> 
                    &nbsp; <input type="text" name="intermcharges" value="<%=intermcharges%>"></td>
                </tr>
                <tr>
                  <td align="right">Tenant Offline</td>
                  <td><input type="checkbox" value="1" name="leaseexpired" <%if leaseexpired="True" then Response.Write "CHECKED"%>></td>
                </tr>
                <tr>
                  <td align="right">LMEP Exempt</td>
                  <td><input type="checkbox" value="1" name="lmepExempt" <%if lmepExempt="True" then Response.Write "CHECKED"%>></td>
                </tr>
                <tr>
                  <td align="right">Online Billing</td>
                  <td><input type="checkbox" value="1" name="onlinebill" <%if onlinebill="True" then Response.Write "CHECKED"%>></td>
                </tr>
                <tr>
                  <td align="right">Revenue Exempt</td>
                  <td><input type="checkbox" value="1" name="ibsexempt" <%if ibsexempt="True" then Response.Write "CHECKED"%>></td>
                </tr>
                <tr>
                  <td align="right">Report Exempt</td>
                  <td><input type="checkbox" value="1" name="bsexempt" <%if bsexempt="True" then Response.Write "CHECKED"%>></td>
                </tr>
                      <%
						strsql = "SELECT * FROM tblTenantVarianceLimits where billingid='" & tid & "'"
						rst1.open strsql, cnn1
						if not rst1.eof then
								WattsperSqFtLow = rst1("WattsPerSqFtLowLimit")
								WattsperSqFtHigh = rst1("WattsPerSqFtHighLimit")
							%>
						<%end if
						rst1.close
						%>  	

                <tr>
                  <td align="right">Watts per SqFt % Variance Limits</td>
					<td align="left">Low <input type="text" name="WattsPerSqFtLowLimit" value="<%=WattsperSqFtLow%>" size="5">
					 High <input type="text" name="WattsPerSqFtHighLimit" value="<%=WattsperSqFtHigh%>" size="5"></td>
                </tr>
									
                <tr>
                  <td align="right">Tenant Email</td>
					<td align="left"><input type="text" name="TenantEmail" value="<%=TenantEmail%>" size="20">
					</td>
                </tr>
												
              </table></td>
          </tr>
          <tr>
            <td colpsan="2"> 
			<%if not(isBuildingOff(bldg)) then%>
              <%if trim(tid)<>"" then%>
              <input type="submit" name="action" value="Update" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;"> 
              <input type="button" name="action" value="Cancel" onclick="document.location='tenantedit_pa.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&edit=0';" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;"> 
              &nbsp; 
              <%else%>
              <input type="submit" name="action" value="Save" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;"> 
              <input type="button" name="action" value="Cancel" onclick="history.back();" class="standard" style="border:1px outset #ddffdd;background-color:ccf3cc;"> 
              &nbsp; 
              <%end if%>
            <%end if%>
            </td>
          </tr>
        </table>
        <% else %>
        <table border=0 cellpadding="3" cellspacing="0" width="100%">
          <tr> 
            <td><b>Account # <%=tenantnum%> (<%=billingname%>)</b></td>
            <td align="right"><input type="button" id="accountbutton" value="Hide Info" onclick="toggleInfoDisplay('account');" class="standard" style="cursor:hand;text-decoration:underline;background-color:#eeeeee;border:0;color:336699;"></td>
          </tr>
        </table>
        <div id="accountinfo" style="display:inline;"> 
          <table border=0 cellpadding="3" cellspacing="0" style="border:1px solid #cccccc;" width="100%">
            <tr> 
              <td> 
                <!-- begin column 1-->
                <table border=0 cellpadding="3" cellspacing="0">
                  <tr valign="top" bgcolor="#eeeeee"> 
                    <td><span class="standard">Account #:</span></td>
                    <td><%=tenantnum%></td>
                  </tr>
                  <% if pid = 108 then %>
						<% if leasenum <> "" then %>
						<tr valign="top" bgcolor="#eeeeee"> 
						  <td><span class="standard">Lease #:</span></td>
						  <td><%=leasenum%></td>
						</tr>
						<% end if %>
						<% if sequencenum <> "" then %>
						<tr valign="top" bgcolor="#eeeeee"> 
						  <td><span class="standard">Sequence #:</span></td>
						  <td><%=sequencenum%></td>
						</tr>
						<% end if %>
                  <% end if %>
                  <% if pid = 45 then %>
					<% if AcctCode <> "" then %>
						<tr valign="top" bgcolor="#eeeeee"> 
						  <td><span class="standard">Account Code #:</span></td>
						  <td><%=AcctCode%></td>
						</tr>
					<% end if %>
                  <% end if %>
                  <tr valign="top" bgcolor="#eeeeee"> 
                    <td><span class="standard">Account Type:</span></td>
                    <td>
                      <%rst1.open "SELECT description FROM account_type WHERE id="&accounttype, getConnect(0,0,"dbCore")
						if not rst1.eof then response.write rst1(0)
						rst1.close%>
                    </td>
                  </tr>
                  <tr valign="top" bgcolor="#eeeeee"> 
                    <td><span class="standard">Lease Start Date:</span></td>
                    <td><%=startdate%></td>
                  </tr>
					<tr>
					  <td valign="top" bgcolor="#eeeeee"><span class="standard">Tenant MoveIn Date:</span></td>
					  <td><%=TenantMoveIndate%></td>
					</tr>
					<tr>                
					  <td valign="top" bgcolor="#eeeeee"><span class="standard">Lease Expiry Date</span></td>
					  <td><%=LeaseExpirydate%></td>
					</tr>                    
                  <tr valign="top" bgcolor="#eeeeee"> 
                    <td><span class="standard">Billing Name:</span></td>
                    <td><%=billingname%></td>
                  </tr>
                  <tr valign="top" bgcolor="#eeeeee"> 
                    <td><span class="standard">Account Name:</span></td>
                    <td><%=tname%></td>
                  </tr>
                  <tr valign="top" bgcolor="#eeeeee"> 
                    <td><span class="standard">Billing Address:</span></td>
                    <td> <%=tstrt%><br> <%=tcity%>
                      <% if tcity<>"" then Response.Write ",&nbsp;" end if%>
                      <%=tstate%>&nbsp;<%=tzip%>&nbsp;<%=tCountry%></td>
                  </tr>
                  <% if corpStreet <> "" then %>
                  <tr valign="top" bgcolor="#eeeeee"> 
                    <td><span class="standard">Corporate Address:</span></td>
                    <td> <%=corpStreet%><br> <%=corpCity%>
                      <% if corpCity<>"" then Response.Write ",&nbsp;" end if%>
                      <%=corpState%>&nbsp;<%=corpZip%>&nbsp;<%=corpCountry%></td>
                  </tr>
                  <% end if %>
                </table>
                <!-- end column 1-->
              </td>
              <td width="8" style="border-left:1px solid #cccccc;">&nbsp;</td>
              <td width="49%"> 
                <!-- begin column 2-->
                <table border=0 cellpadding="3" cellspacing="0">
                  <tr valign="top" bgcolor="#eeeeee"> 
                    <td><span class="standard">Floor:</span></td>
                    <td><%=flr%></td>
                  </tr>
                  <tr valign="top" bgcolor="#eeeeee"> 
                    <td><span class="standard">SQFT:</span></td>
                    <td><%=sqft%></td>
                  </tr>
                  <tr valign="top" bgcolor="#eeeeee"> 
                    <td><span class="standard">Tax Exempt:</span></td>
                    <td> 
                      <%if taxexempt="True" then Response.Write "Yes" else Response.Write "No"%>
                    </td>
                  </tr>
                  <tr valign="top" bgcolor="#eeeeee"> 
                    <td><span class="standard">Interim Charges:</span></td>
                    <td> 
                      <%if interm="False" then Response.Write "None" else Response.Write intermcharges end if%>
                    </td>
                  </tr>
                  <tr valign="top" bgcolor="#eeeeee"> 
                    <td><span class="standard">Tenant Offline:</span></td>
                    <td> 
                      <%if leaseexpired="True" then Response.Write "Yes" else Response.Write "No" end if%>
                    </td>
                  </tr>
                  <!-- onlinebill -->
                  <tr valign="top" bgcolor="#eeeeee"> 
                    <td><span class="standard">LMEP Exempt:</span></td>
                    <td> 
                      <%if lmepExempt="True" then Response.Write "Yes" else Response.Write "No" end if%>
                    </td>
                  </tr>
                  <tr valign="top" bgcolor="#eeeeee"> 
                    <td><span class="standard">Online Billing:</span></td>
                    <td> 
                      <%if onlinebill="True" then Response.Write "Yes" else Response.Write "No" end if%>
                    </td>
                  </tr>
                  <tr valign="top" bgcolor="#eeeeee"> 
                    <td><span class="standard">Revenue Exempt:</span></td>
                    <td> 
                      <%if ibsexempt="True" then Response.Write "Yes" else Response.Write "No" end if%>
                    </td>
                  </tr>
                  <tr valign="top" bgcolor="#eeeeee"> 
                    <td>Report <br>
                      exempt:</td>
                    <td valign="bottom"> 
                      <%if bsexempt="True" then Response.Write "Yes" else Response.Write "No" end if%>
                    </td>
                  </tr>
                  <tr valign="top" bgcolor="#eeeeee"> 
                    <td>Tenant Email:</td>
                    <td valign="bottom"> 
                      <%=TenantEmail%>
                    </td>
                  </tr>                  
                </table>
                <!-- end column 2-->
              </td>
            </tr>
          </table>
        </div>
			<%if trim(tid)<>"" and not(isBuildingOff(bldg)) then%>
			<input type="button" name="action" value="Edit Account Information" onclick="document.location='tenantedit_pa.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&edit=1';" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;margin-top:3px;"> 
			<%end if%>
        
      </td>
    </tr>
        </table>
       

      </td>
    </tr>
    <% end if %>
  </table>
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="tid" value="<%=tid%>">


<table width="100%" border=0 cellpadding="0" cellspacing="0">
<%if trim(tid)<>"" then
  rst1.open "select case when use_acctid=1 then 'Uses Rate Account' else case when acctid<>'0' then acctid else 'Default' end end as acctidname, tlup.*, f.[description], utilitydisplay from tblleasesutilityprices tlup left join tblutility tu on tu.utilityid=tlup.utility left join functiontypes f on f.id=tlup.procname WHERE billingid="&tid, cnn1, adopendynamic

   do until rst1.EOF%>
   <tr valign="top">
    <td bgcolor="#eeeeee" style="border-bottom:1px solid #cccccc;">
    <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr>
      <td><a name="<%=rst1("utilitydisplay")%>"></a><b><%=rst1("utilitydisplay")%> - (Lease ID SVR<%=split(getBuildingIP(bldg),"\")(1)%>-<%=rst1("leaseutilityid")%>)</b></td>
      <td align="right">
      <input type="button" id="<%=rst1("utilitydisplay")%>button" value="Show Info" onclick="toggleInfoDisplay('<%=rst1("utilitydisplay")%>');" class="standard" style="cursor:hand;text-decoration:underline;background-color:#eeeeee;border:0;color:336699;">&nbsp;
      </td>
    </tr>
    </table>
    <div id="<%=rst1("utilitydisplay")%>info" style="display:none;">
	<table border=0 cellpadding="2" cellspacing="0">
	<tr bgcolor="#eeeeee">
		<td>Billing Account:</td>
		<td><%=rst1("acctidname")%></td>
		<td width="12">&nbsp;</td>
		<td>Admin Fee:</td>
		<td><%=rst1("adminfee")%></td>
		<td width="12">&nbsp;</td>
		<td>Full On Peak:</td>
		<td><%if rst1("fullonpeak")="True" then Response.Write "Yes" else Response.Write "No" end if%></td>
		<td width="12">&nbsp;</td>
		<td>Shadow Bill:</td>
		<td><%if rst1("Shadow")="True" then Response.Write "Yes" else Response.Write "No" end if%></td>
	</tr>
	<tr bgcolor="#eeeeee">
		<td>Coincident:</td>
		<td><%if rst1("coincident")="True" then Response.Write "Yes" else Response.Write "No" end if%></td></td>
		<td width="12">&nbsp;</td>
		<td>Modify Rate:</td>
		<td><%=rst1("ratemodify")%></td>
		<td width="12">&nbsp;</td>
		<td>Coincident w/Building Peak:</td>
		<td><%if rst1("Coincident_peak")="True" then Response.Write "Yes" else Response.Write "No" end if%></td>
	</tr>
	<tr bgcolor="#eeeeee">
		<td>Add-on Fee:</td>
		<td><%=rst1("addonfee")%></td>
		<td width="12">&nbsp;</td>
		<td>Account Rate:</td>
		<td>
		<%
		if trim(rst1("ratetenant"))<>"" then
		rst2.open "SELECT type FROM ratetypes WHERE id='" & rst1("ratetenant") & "'", getMainConnect(pid) %>
		<%if not rst2.eof then %><%=rst2("type")%><% end if %>
		<% rst2.close 
		end if
		%>
		</td>
		<td width="12">&nbsp;</td>
		<td>Intermediate Peak:</td>
		<td><%if rst1("calcintpeak")="True" then Response.Write "Yes" else Response.Write "No" end if%></td>
    </tr>
    <tr><td colspan="5"></td></tr>
    </table><br>
    </div>
    </td>
  </tr>
  <tr>
    <td style="border-bottom:1px solid #cccccc">
    <table border=0 cellpadding="3" cellspacing="0" width="100%">
    <%
      
      rst2.Open "SELECT * FROM meters WHERE leaseutilityid='"&rst1("leaseutilityid")&"'", cnn1
      if not rst2.EOF then%>
        <tr bgcolor="#dddddd">
          <td><span class="standard"><b>Meter</b></span></td>
          <td><span class="standard"><b>Start Date</b></span></td>
          <td><span class="standard"><b>Date Off</b></span></td>
          <td><span class="standard"><b>Last Read</b></span></td>
          <td><span class="standard"><b>Location</b></span></td>
          <td><span class="standard"><b>Floor</b></span></td>
          <td><span class="standard"><b>Riser</b></span></td>
        </tr>
    
        <%do until rst2.EOF%>
        <tr bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" >
          <td><span class="standard"><%=rst2("meternum")%></span></td>
          <td><span class="standard"><%=rst2("datestart")%></span></td>
          <td><span class="standard"><%=rst2("dateoffline")%></span></td>
          <td><span class="standard"><%=rst2("datelastread")%></span></td>
          <td><span class="standard"><%=rst2("location")%></span></td>
          <td><span class="standard"><%=rst2("floor")%></span></td>
          <td><span class="standard"><%=rst2("riser")%></span></td>
        </tr>
        
        <%rst2.movenext
        loop%>
      <%
      else
      %>
      <tr bgcolor="#dddddd"><td><span class="standard"><b>Meters</b></span></td></tr>
      <tr bgcolor="#ffffff">
        <td><span class="standard">There are no meters set up for this lease utility.<br></span></td>
      </tr>
      <%
      end if
      rst2.close
      %>
    </table>
    </td>
  </tr>
      
      <%
      rst1.movenext
      loop
    rst1.close
  end if
%>
  <tr><td height="500"><br></td></tr>
	</table>


</form>
</body>
</html>
