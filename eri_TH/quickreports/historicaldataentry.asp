<%@Language="VBScript"%>

<% 
'15/19/2008 N.Ambo added this screen to allow user to enter historical data and save to the utiltiy bill table
'User will have to select criteria of utility and bill period, the building number is already passed unto the page via the querystring
%>
<HTML>
	<HEAD>
		<title>Historical Data Entry</title> 
		<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
		<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
		<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
		<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">
			<script id="clientEventHandlersJS" language="javascript">
<!--

function Select2_onblur() {

}

//-->
</script>

	</HEAD>
	<body>
<%
dim pid, building, utilityid, meterid, byear, bperiod, showvalues
dim buttonval, totalkwh, totalkw,totalkwhcost,totalkwcost,totalbilled

dim rst1,rst2,rst3, cnn1, sql, cmd, cmd2, cmd3, cmd4, prm, prm2, sql2

set rst1 = server.createobject("ADODB.Recordset")
set rst2 = server.createobject("ADODB.Recordset")
set rst3 = server.createobject("ADODB.Recordset")

set cnn1 = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set cmd2 = server.createobject("ADODB.Command")
set cmd3 = server.createobject("ADODB.Command")
set cmd4 = server.createobject("ADODB.Command")

totalkwh = request.Form("totalkwh")
totalkw = request.Form("totalkw")
totalkwhcost = request.Form("totalkwhcost")
totalkwcost = request.Form("totalkwcost")
totalbilled = request.Form("totalbilled")


showvalues = request("showvalues")
pid = request("pid")
building = request("bldgNum")
utilityid = request("utilityid")
'meterid = request("meterid")

if instr(request("bperiod"),"/")>0 then
	byear = split(request("bperiod"),"/")(1)
	bperiod = split(request("bperiod"),"/")(0)
else
	byear = request("byear")
	bperiod = request("bperiod")
end if


cnn1.open getLocalConnect(building)

cmd.CommandType = adCmdStoredProc
cmd.CommandText = "getUtilsPerBldg"
Set prm = cmd.CreateParameter("bldgnum", adVarChar, adParamInput, 20, building)
cmd.Parameters.Append prm
cmd.Name = "test"

Set cmd.ActiveConnection = cnn1
cnn1.test   rst1 


cmd2.CommandType = adCmdStoredProc
cmd2.CommandText = "getMetersPerUtility"
Set prm = cmd2.CreateParameter("bldgnum", adVarChar, adParamInput, 20, building)
cmd2.Parameters.Append prm
Set prm2 = cmd2.CreateParameter("utility", adInteger, adParamInput, , utilityid)
cmd2.Parameters.Append prm2
cmd2.Name = "test2"

if utilityid <> "" and building <> "" then
	Set cmd2.ActiveConnection = cnn1
    cnn1.test2 rst2     
    
    sql = "SELECT distinct cast(billperiod as varchar)+'/'+billyear as periodyear, billyear, billperiod FROM billyrperiod WHERE "
	'if not(historic) then sql = sql & "billyear>=year(getdate())-1 and "
	sql = sql & "bldgnum='"&building&"' and utility="&utilityid&" order by billyear, billperiod"
	'response.Write sql
	rst3.open sql, cnn1 
     
 end if
 

 %>
		<form method="post" name="form1" action="index.asp">
			<table width="100%" border="0" cellpadding="3" cellspacing="0">
				<tr>
					<td width="52%" bgcolor="#6699cc"><span class="standardheader">Historical Data Entry: 
			<%=bldgname%></span></td>
					<td width="48%" bgcolor="#6699cc"><div align="right">
							<input name="button" type="button" value="Building Cost Analysis" onclick="LoadReport()" ></div>
					</td>
				</tr>
			</table>
			<table border=0 cellpadding="3" cellspacing="0" ID="Table2">
          <tr> 
        
			<% if trim(building)<>"" then %>
			<td> <select name="utilityid" onChange="loadutility()">
				<option value="0">Select Utility</option>					
					<%do until rst1.eof   %>
					<option value="<%=rst1("utilityid")%>"<%if trim(rst1("utilityid"))=trim(utilityid) then response.write " SELECTED"%>><%=rst1("utility")%></option>
				<% rst1.movenext
				loop%>
				</select> </td>	
			 <%end if 
			 rst1.close%>      
       			
			<td> <select name="bperiod" onchange="getYPID()">
                <option value="0">Select Bill Period</option>
               <% if utilityid <> "" and building <> "" then %>
				<%do until rst3.eof%>
				<option value="<%=rst3("periodyear")%>"<%if trim(rst3("periodyear"))=trim(bperiod&"/"&byear) or (bperiod="0" and month(dateadd("m",-1,now))&"/"&year(dateadd("m",-1,now))=rst3("periodyear")) then response.write " SELECTED"%>><%=rst3("periodyear")%></option>
				 <% rst3.movenext
				loop
				rst3.close 					
				end if %>	
              </select> </td>			
			
			
          
			<td>		
				<input type="hidden" name="bldgNum" value="<%=Building%>" > 	
				<input type="hidden" name="pid" value="<%=pid%>" > 			
				<input type="button" onclick="loadvalues()" name="Select" value="Select" >
				<input type="button" value="AddPeriod" name="AddPeriod" onClick="billPeriodAdd('');" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" > 
				<input type="button" value="EditPeriod" name="EditPeriod" onClick="billPeriodEdit()" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" > 
            </td>
            
          </tr>
        </table>
        
		</form>
		<script language="javascript" type="text/javascript">
// <!CDATA[

function loadvalues() {
	var frm = document.forms['form1'];
	if (frm.utilityid.value == "0" || frm.bperiod.value == "0") {
		alert("You must select both the utility and bill period!")
	}
	else {
		var newhref = "historicaldataentry.asp?pid="+frm.pid.value+"&bldgNum="+frm.bldgNum.value+"&utilityid="+frm.utilityid.value+"&bperiod="+frm.bperiod.value+"&showvalues=1";
		document.location.href=newhref;
	}
}
function loadutility()
{	var frm = document.forms['form1'];
	var newhref = "historicaldataentry.asp?pid="+frm.pid.value+"&bldgNum="+frm.bldgNum.value+"&utilityid="+frm.utilityid.value;
	document.location.href=newhref;
}
function billPeriodAdd()
{
	var frm = document.forms['form1'];	
	var newhref = "AddbillPeriod.asp?pid="+frm.pid.value+"&bldgNum="+frm.bldgNum.value;
	document.location.href=newhref;
}
function billPeriodEdit()
{
	var frm = document.forms['form1'];
	
	if (frm.bperiod.value == "0") {
		alert("You must select a bill period!")
		}
	else {
		
		var newhref = "AddbillPeriod.asp?pid="+frm.pid.value+"&bldgNum="+frm.bldgNum.value+"&bperiod="+frm.bperiod.value+"&utilityid="+frm.utilityid.value;
		document.location.href=newhref;
	}
}
function AddData()
{	
//this function is currently not being used
	var frm = document.forms['form2'];
	var newhref = "SaveBillPeriod.asp?";
	//if (IsNumeric(frm.totalkwh.value)== false){
		//alert("value is not numeric!");
	//}
	document.location.href=newhref;
}
function LoadReport()
{	
	var frm = document.forms['form1'];
	var newhref = "index.asp?pid="+frm.pid.value+"&bldg="+frm.bldgNum.value+"&utility="+frm.utilityid.value+"&de=1";
	document.location.href=newhref;
}
function IsNumeric(strString)
   //  check for valid numeric strings	
{
   var strValidChars = "0123456789.";
   var strChar;
   var blnResult = true;

   if (strString.length == 0) return false;

   //  test strString consists of valid characters listed above
   for (i = 0; i < strString.length && blnResult == true; i++)
      {
      strChar = strString.charAt(i);
      if (strValidChars.indexOf(strChar) == -1)
         {
         blnResult = false;
         }
      }
   return blnResult;
}




// ]]>
		</script>
	
<%
if (showvalues= 1 and utilityid <> "0" and billyear <> "0" and billperiod <> "0") then
%>
	<STRONG><FONT face="Arial" size="2">Period: <%=bperiod+" "+byear%></FONT></STRONG>
<%	
cmd3.CommandType = adCmdStoredProc
cmd3.CommandText = "getBillingData"
Set prm = cmd.CreateParameter("bldgnum", adVarChar, adParamInput, 10, building)
cmd3.Parameters.Append prm
Set prm = cmd.CreateParameter("billyear", adVarChar, adParamInput, 4, byear)
cmd3.Parameters.Append prm
Set prm = cmd.CreateParameter("billperiod", adVarChar, adParamInput, 4, bperiod)
cmd3.Parameters.Append prm
Set prm = cmd.CreateParameter("utility", adInteger, adParamInput, , utilityid)
cmd3.Parameters.Append prm
cmd3.Name = "test3"

Set cmd3.ActiveConnection = cnn1
cnn1.test3 rst1 

if rst1.EOF then
	buttonval = "Save"
	totalkwh = 0
	totalkw = 0
	totalkwhcost = 0
	totalkwcost = 0
	totalbilled = 0
else
	totalkwh = rst1("totalkwh")
	totalkw = rst1("totalkw")
	totalkwhcost = rst1("costkwh")
	totalkwcost = rst1("costkw")
	totalbilled = rst1("totalbillamt")
	buttonval = "Update"
end if
rst1.Close

%>
<form method="post" action="SaveData.asp" name="form2" >
<%if utilityid = 4 then%>
<TABLE id="Table3" cellSpacing="1" cellPadding="1" width="304" bgColor="gainsboro"
			border="1">
	<TR>
		<TD width="127"><FONT face="Arial" size="2"><STRONG>Total Therms:</STRONG></FONT></TD>
		<TD><INPUT id="Text6" type="text"   name="totalkwh" value="<%=totalkwh%>" ></TD>
	</TR>
	<TR>
		<TD width="127"><FONT face="Arial" size="2"><STRONG>Total Billed:</STRONG></FONT></TD>
		<TD><INPUT id="Text10" type="text" name="totalbilled" value="<%=totalbilled%>"</TD>
	</TR>
	<INPUT  type="hidden" name="totalkw" value="<%=totalkw%>">
	<INPUT  type="hidden" name="totalkwhcost" value="<%=totalkwhcost%>">
	<INPUT  type="hidden" name="totalkwcost" value="<%=totalkwcost%>">
</TABLE>

<%else%>

<TABLE id="Table1" height="148" cellSpacing="1" cellPadding="1" width="304" bgColor="gainsboro"
			border="1">
	<TR>
		<TD width="127"><FONT face="Arial" size="2"><STRONG>Total KWH:</STRONG></FONT></TD>
		<TD><INPUT id="Text1" type="text"   name="totalkwh" value="<%=totalkwh%>" ></TD>
	</TR>
	<TR>
		<TD width="127"><FONT face="Arial" size="2"><STRONG>Total KW:</STRONG></FONT></TD>
		<TD><INPUT id="Text2" type="text" name="totalkw" value="<%=totalkw%>"></TD>
	</TR>
	<TR>
		<TD width="127"><FONT face="Arial" size="2"><STRONG>Total KWH Cost:</STRONG></FONT></TD>
		<TD><INPUT id="Text3" type="text" name="totalkwhcost" value="<%=totalkwhcost%>"></TD>
	</TR>
	<TR>
		<TD width="127"><FONT face="Arial" size="2"><STRONG>Total KW cost:</STRONG></FONT></TD>
		<TD><INPUT id="Text4" type="text" name="totalkwcost" value="<%=totalkwcost%>"></TD>
	</TR>
	<TR>
		<TD width="127"><FONT face="Arial" size="2"><STRONG>Total Billed:</STRONG></FONT></TD>
		<TD><INPUT id="Text5" type="text" name="totalbilled" value="<%=totalbilled%>"</TD>
	</TR>
</TABLE>

<%end if%>
<input type="hidden" name="pid" value="<%=pid%>" ID="Hidden3"> 
<input type="hidden" name="bldgNum" value="<%=building%>" > 
<input type="hidden" name="utilityid" value="<%=utilityId%>" > 
<input type="hidden" name="byear" value="<%=byear%>" ID="Hidden1"> 
<input type="hidden" name="bperiod" value="<%=bperiod%>" ID="Hidden2"> 
<INPUT type="submit" value="<%=buttonval%>" NAME="AddData" >
<INPUT type="button" value="Cancel" NAME="Cancel" onclick="loadvalues()">
</form>
<%
	
else
	response.Write("Required criteria is missing. Please select an option for utility and bill period.")
end if

cnn1.Close
%>
	</body>
</HTML>
