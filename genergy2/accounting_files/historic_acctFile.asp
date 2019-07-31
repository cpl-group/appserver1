<%option explicit%>
<% 'TT 5/22/2008 UM page points to entry.asp (this is the original page) and G1console points to entryG1.asp.  This is so client can only view their portfolio ONLY and no longer has the JUMPTO option. TT %>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Utility Bill Entry</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">		
</head>
<%
Dim cnn1, rst1, rst2, str1, str, dateFrom, dateTo, lid, util, billy, billp, dumDate
dim acctid, pid, bldg, sqlStr, coreCmd, prm, i

bldg = request("bldg")
pid = request("pid")
util = request("util")
lid = request("lid")

dumDate = ""

if (request("billPeriod") <> "") then
    billy = split(request("billperiod"),"/")(1)
	billp = split(request("billperiod"),"/")(0)
end if

if (request("dateFrom") = "") then
    dateFrom = "mm/dd/yy"
else
    dateFrom = request("dateFrom")
end if

if (request("dateTo") = "") then
    dateTo = "mm/dd/yy"
else
    dateTo = request("dateTo")
end if

Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = server.createObject("ADODB.RecordSet")
Set rst2 = server.createObject("ADODB.RecordSet")
set coreCmd = server.createobject("ADODB.command")

cnn1.Open getConnect(pid,bldg,"Billing")

%>
<script src="/genergy2/calendar.js"></script>
<script src="/genergy2/sorttable.js"></script>
<script src="acctFile.js"></script>
<body bgcolor="#eeeeee" text="#000000">
<form name="acctTransForm" method="post" action="historic_acctFile.asp">      
<table width="100%" border="0">
		<tr>
			<td>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="91%" bgcolor="#6699cc"><span class="standardheader">Historic Accounting Transaction for <%=bldg%></span></td>
            <td width="9%" align="right" bgcolor="#6699cc"><% if trim(bldg)<>"" then %><select name="select" onChange="JumpTo(this.value, '<%=pid%>', '<%=bldg %>')">
                <option value="#" selected>Jump to...</option>
                <option value="/genergy2/billing/processor_select.asp">Bill Processor</option>
                <option value="../validation/re_index.asp">Review Edit</option>
                <option value="/genergy2/setup/buildingedit.asp">Building Setup</option>
                <option value="/genergy2/manualentry/entry_select.asp">Manual Entry</option>
                <option value="/genergy2/UMreports/meterProblemReport.asp">Meter Problem Report</option>
                <option value="/genergy2/accounting_files/historic_acctFile.asp">Accounting Transactions</option>
              </select><% end if %></td>
          
        </table>
				<table width="100%" border="0">
					<%Set rst1 = Server.CreateObject("ADODB.recordset")
					if not(trim(pid)="" and trim(bldg)<>"") then%>
						<tr>
							<td width="25%" height="27" align="right">
								Portfolio:
							</td>
							<td width="75%" height="27">								
								<%if allowGroups("Genergy Users") then%>
									<select name="pid" onChange="">
										<option value="">Select Portfolio...</option>
										<%rst1.open "SELECT distinct id, name FROM portfolio p WHERE id='" + pid + "' ORDER BY name", getConnect(0,0,"dbCore")
										do until rst1.eof
											%><option value="<%=trim(rst1("id"))%>"<%if trim(rst1("id"))=trim(pid) then %>SELECTED<%end if%>>
												<%=rst1("name")%>
											</option><%
											rst1.movenext
										loop
										rst1.close%>
									</select>
								<%elseif isnumeric(pid) then
									rst1.open "SELECT name FROM portfolio WHERE id="&pid&" ORDER BY name", cnn1
									if not rst1.eof then response.write rst1("name") end if
									rst1.close%>
									<input type="hidden" name="pid" value="<%=pid%>">
								<%end if%>								
							</td>
						</tr>
						<tr>
							<td width="25%" height="27" align="right">Building:</td>
							<td width="75%" height="27">
								<%if trim(pid)<>"" then %>
									<select name="bldg" onChange="loadbuilding('<%=pid%>', this.value)">
									<option selected>Select Building...</option>
									<%
										rst1.open "SELECT BldgNum, strt FROM buildings b WHERE portfolioid='"&pid&"' ORDER BY strt", cnn1
										do until rst1.eof
											%><option value="<%=trim(rst1("Bldgnum"))%>"<%if trim(rst1("Bldgnum"))=trim(bldg) then%>selected<%end if%>>
												<%=rst1("strt")%> (<%=rst1("Bldgnum")%>)
											</option><%
											rst1.movenext
										loop
										rst1.close%>
									</select>								
								<%else%>
									<input type="hidden" name="building" value="">No Building Selected
								<%end if%>
								</td>
						</tr>
						<tr>
						   <% rst1.open "SELECT DISTINCT tName, LeaseUtilityId FROM tblLeases l INNER JOIN tblLeasesUtilityPrices p ON p.billingId = l.billingId WHERE bldgNum='"+bldg+"' ORDER BY tName", cnn1  %>
						    <td align="right">Tenant: </td>
						    <td>&nbsp;<select name="lid" onChange=""> 
						        <option selected>Select Tenant...</option>
						        <option value=0>All</option>
						        <% do until rst1.eof %>
						        <option value="<%=trim(rst1("LeaseUtilityId"))%>" <%if (trim(rst1("LeaseUtilityId"))=trim(lid)) then %>selected <%end if %> ><%=trim(rst1("LeaseUtilityId")) + " - " + trim(rst1("tName"))%></option>
						        <% rst1.movenext
						           loop
						           rst1.close
						         %>
						    </select></td>
						</tr>
						<tr>
						   <% rst1.open "SELECT DISTINCT utilityId, utility FROM tblUtility ORDER BY utility", getConnect(0,0,"dbCore") %>
						    <td align="right">Utility: </td>
						    <td>&nbsp;<select name="util" onChange=""> 
						        <option selected>Select utility...</option>
						        <% do until rst1.eof %>
						        <option value="<%=trim(rst1("utilityId"))%>" <% if trim(rst1("utilityId")) = util then %>selected <%end if %> ><%=trim(rst1("utility"))%></option>
						        <% rst1.movenext
						           loop
						           rst1.close
						         %>
						    </select></td>
						</tr>
					<tr>
					    <td align="right">Post Date From:</td>
					    <td>&nbsp;<input type="text" name="dateFrom" value="<%=dateFrom%>" onfocus="this.select();lcs(this)" onclick="event.cancelBubble=true;this.select();lcs(this)">
					    </td>
					</tr>
					    
					<tr>
					 <td align="right">Date To: </td>
					 <td>&nbsp;<input type="text" name="dateTo" value="<%=dateTo%>" onfocus="this.select();lcs(this)" onclick="event.cancelBubble=true;this.select();lcs(this)"></td>      
					</tr>
					<tr>
					<% 
					    sqlStr = "SELECT distinct cast(billperiod as varchar)+'/'+billyear as periodyear, billyear, billperiod FROM billyrperiod WHERE bldgNum='"+bldg+"' ORDER BY billyear, billperiod" 
					    rst1.open sqlStr, cnn1
					    
					    if not rst1.eof then
					%>
					  <td align="right"> OR &nbsp;&nbsp;&nbsp; Bill Period:</td>
					  <td><select name="billPeriod">
					  <option value="">Select Bill Period...</option>
					  <% do until rst1.eof %>
					  <option value="<%=trim(rst1("periodyear"))%>" <%if trim(rst1("periodyear")) = request("billPeriod") then %>selected <%end if %> ><%=trim(rst1("periodyear"))%></option>
					  
					   <% rst1.movenext
			               loop
			            %>
			            </select>
			          </td>
					</tr>
					<% rst1.close 
					end if %>
					<tr>
					 <td></td>
					 <td><input type="submit" name="submit" value="View Transactions" onclick="return checkForm();" />
					     &nbsp; &nbsp; <input type="submit" name="submit" value="Create Acct File" onclick="return checkForm();" />
					 </td>
					</tr>
				</table>
		        <%end if %>
			</td>
		</tr>		
</table>
</form>

<br />

<% 


if (request("submit") = "View Transactions") then 
    Set prm = coreCmd.CreateParameter("bldg", adVarChar, adParamInput, 50)
    coreCmd.Parameters.Append prm
    Set prm = coreCmd.CreateParameter("dateFrom", adVarChar, adParamInput, 20)
    coreCmd.Parameters.Append prm
    Set prm = coreCmd.CreateParameter("dateTo", adVarChar, adParamInput, 20)
    coreCmd.Parameters.Append prm
    Set prm = coreCmd.CreateParameter("BY", adVarChar, adParamInput, 10)
    coreCmd.Parameters.Append prm
    Set prm = coreCmd.CreateParameter("BP", adTinyInt, adParamInput)
    coreCmd.Parameters.Append prm
    Set prm = coreCmd.CreateParameter("lid", adInteger, adParamInput)
    coreCmd.Parameters.Append prm
    Set prm = coreCmd.CreateParameter("UTIL", adInteger, adParamInput)
    coreCmd.Parameters.Append prm

    coreCmd.Parameters("bldg") = trim(bldg)
    if (dateFrom = "" OR dateFrom = "mm/dd/yy") then
        coreCmd.Parameters("dateFrom") = ""
    else
        coreCmd.Parameters("dateFrom") = CDate(dateFrom)
    end if 
    
    if (dateTo = "" OR dateTo = "mm/dd/yy") then
        coreCmd.Parameters("dateTo") = ""
    else
        coreCmd.Parameters("dateTo") = CDate(dateTo)
    end if
    
    coreCmd.Parameters("BY") = trim(billy)
    coreCmd.Parameters("BP") = CInt(billp)
    coreCmd.Parameters("lid") = CInt(lid)
    coreCmd.Parameters("UTIL") = CInt(util)
    
    coreCmd.ActiveConnection = cnn1
    coreCmd.CommandText = "sp_getAcctTrans"
    coreCmd.CommandType = adCmdStoredProc
    
    rst2.open coreCmd
       
%>
<div id="entryframe" >
<!--IFRAME name="entry" width="100%" height="550" src="" scrolling="no" marginwidth="0" marginheight="0" frameborder="0"></IFRAME -->
<table id="sortTable" class="sortable" style="font-size: 11px; font-family: Arial, Helvetica, sans-serif;" cellspacing="1" cellpadding="3" border="1" width="99%">
                	<thead align="center">
                    	<% for i = 0 to rst2.fields.Count - 1 %>
                        <th><a href="#"><%=rst2.fields(i).Name%></a></th>
                        <%next%>
                    </thead>
                    <tbody align="center">
                    	<%do while not rst2.eof%>
                    	<tr onMouseOver="this.style.backgroundColor = 'lightgreen'" onMouseOut="this.style.backgroundColor = '#eeeeee'" onclick="popUp('accountingFile_output.asp?bldg=<%=rst2("bldgNum")%>&util=<%=rst2("utility")%>&tenant=<%=rst2("leaseNo")%>&billyear=<%=billy%>&billperiod=<%=billp%>&datefrom=<%=rst2("transdate")%>', 'createFiles')" > 
                    	<%for i = 0 to rst2.fields.Count - 1%>
                    		<td style="border-bottom: 1px solid #CCCCCC"><%=UCase(rst2(i))%></td>
                    	<%next%>
                    	</tr><%
                    rst2.movenext
                    loop%>
                	</tbody>
            	</table> 
</div>
<%
else
    if (request("submit") = "Create Acct File") then %>
     <IFRAME name="entry" width="100%" height="550" src="accountingFile_output.asp?bldg=<%=bldg%>&util=<%=util%>&lid=<%=lid%>&billyear=<%=billy%>&billperiod=<%=billp%>&dateFrom=<%=dateFrom%>&dateTo=<%=dateTo%>" scrolling="no" marginwidth="0" marginheight="0" frameborder="0"></IFRAME>
<%    
    end if
    
end if

'TK: 04/28/2006
on error resume next
If rst1.State = 1 Then
	rst1.Close 
End If
If rst2.State = 1 Then
	rst2.Close 
End If
set rst1 = nothing
set rst2 = nothing
set cnn1=nothing%>
</body>
</html>