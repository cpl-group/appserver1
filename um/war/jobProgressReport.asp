<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
	

   
		
	Dim rstJobResults, objCnn	
	Dim strCompanyId, strRunDate, strMonth, strYear

	set rstJobResults = server.createobject("ADODB.Recordset")
	set objCnn = server.createobject("ADODB.Connection")

	objCnn.open getConnect(0,0,"Intranet")
	
	strCompanyId = request.Form("optCompanyId")
	strMonth = request.Form("monthNum")
	strYear = request.Form("optYear")


	if trim(strMonth) = "" then
		if Month(Now) < 10 then
			strMonth = "0" & trim(Cstr(Month(Now)))
		else
			strMonth = trim(Cstr(Month(Now)))
		end If
	end if
	
	if trim(strYear) = "" then 
		strYear = trim(Cstr(Year(Now)))
	end if
	'format rundate
	

	' Set the default value for company as Genergy
	if trim(strCompanyId) = "" then
		strCompanyId = "GE"
	end if
	

	

%>
<html>
<head>
<title>Job Progress Report</title>
	<script language="JavaScript" type="text/javascript">
		function loadResults()
		{	
			
			var frm = document.forms['form1'];
			if((frm.optCompanyId.value!='')&&(frm.monthNum.value!='')&&(frm.optYear.value!=''))
			{	
				var newhref = "ProcessResults.asp?optCompanyId="+frm.optCompanyId.value+"&monthNum="+frm.monthNum.value+"&optYear="+frm.optYear.value;
			}

			var xmlHTTP = new ActiveXObject("Microsoft.XMLHTTP");
				
			if (xmlHTTP != null)
			{	
				
				xmlHTTP.open("GET",newhref,false);
				xmlHTTP.send();
			}
			
			var resultset = document.getElementById("ResultsTable");
			resultset.cursor = 'wait';
			this.setTimeout('changeMouse()', 1000); 
			resultset.innerHTML = xmlHTTP.responseText;
			resultset.cursor = 'auto';
		}

		function changeMouse()
		{
			document.getElementById("ResultsTable").cursor='auto';
		}

		function setformaction(act)
		{
			document.form1.action = act
		}

		function openwin(url,mwidth,mheight){
		window.name="opener";
		popwin = window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth);
		popwin.focus();
		}
	</script>

	<script language="JavaScript" type="text/javascript">
		if (screen.width > 1024) {
		document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
		} else {
		document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
		}
	</script>
</head>
<% If not allowGroups("Genergy_Corp,AR_Admin,gAccounting,IT Services") then %>
<body bgcolor="#eeeeee">	
	<table width="100%" border="0" cellpadding="3" cellspacing="0" ID="Table1">
		<tr><td>Security Restriction: You do not have rights to view this report.</td></tr>	</table>
</body>
</html>  		
<% else %>
<body bgcolor="#eeeeee">
	<form name="form1" target="_top" ID="Form1">
	<table width="100%" border="0" cellpadding="3" cellspacing="0">
		<tr>
			<td colspan=8 bgcolor="#228866" align=center><span class="standardheader">Job Progress Report</span></td>
		</tr>
		<tr>
			<td style="width:15%">&nbsp;</td>
			<td style="width:12%;font-weight:bold" align=right >Select Company :</td>
			<td width="10%" style="border-top:1px solid #ffffff;" align=left>
				
					<select name="optCompanyId">
						<%rstJobResults.open "SELECT distinct company FROM report ORDER BY company", objCnn 
						do until rstJobResults.eof%>
						<option value="<%=trim(rstJobResults("company"))%>"<%if trim(rstJobResults("company"))=trim(strCompanyId) then response.write " SELECTED"%>><%=rstJobResults("company")%></option>
						<%	rstJobResults.movenext
						loop
						rstJobResults.close%>
					</select>
			</td>
			<td style="width:10%;font-weight:bold" align=right>Select Month :</td>
			<td width="10%" style="border-top:1px solid #ffffff;"  align=left>
					<select name="monthNum" ID="monthNum">
						<option value="01" <% if strMonth = "01" then Response.Write "Selected"  end if %>  >Jan</option>
						<option value="02" <% if strMonth = "02"  then Response.Write "Selected" end if %>  >Feb</option>
						<option value="03" <% if strMonth = "03"  then Response.Write "Selected" end if %> >Mar</option>
						<option value="04" <% if strMonth = "04"  then Response.Write "Selected" end if %> >Apr</option>
						<option value="05" <% if strMonth = "05" then Response.Write "Selected" end if %> >May</option>
						<option value="06" <% if strMonth = "06" then Response.Write "Selected" end if %> >Jun</option>
						<option value="07" <% if strMonth = "07" then Response.Write "Selected" end if %> >Jul</option>
						<option value="08" <% if strMonth = "08" then Response.Write "Selected" end if %> >Aug</option>
						<option value="09" <% if strMonth = "09" then Response.Write "Selected" end if %> >Sept</option>
						<option value="10" <% if strMonth = "10" then Response.Write "Selected" end if %> >Oct</option>
						<option value="11" <% if strMonth = "11" then Response.Write "Selected" end if %> >Nov</option>
						<option value="12" <% if strMonth = "12" then Response.Write "Selected" end if %> >Dec</option>
					</select>
			</td>
			<td style="width:10%;font-weight:bold" align=right>Select Year :</td>
			<td width="10%" style="border-top:1px solid #ffffff;"  align=left>
					<select name="optYear" ID="optYear">
						<option value="2005" <% if strYear = "2005" then response.write "selected" end if %>>2005</option>
						<option value="2006" <% if strYear = "2006" then response.write "selected" end if %>>2006</option>
						<option value="2007" <% if strYear = "2007" then response.write "selected" end if %>>2007</option>
					</select>
			</td>
			<td width="23%" style="border-top:1px solid #ffffff;" align=left   >
				
				<span id="GetResults" style="font-weight:bold; border=1" onclick="loadResults();" 
					onmouseover="this.style.backgroundColor='#ccff66';" 
					onmouseout="this.style.backgroundColor='#eeeeee';">
				<img src="PRINT_16.GIF" alt="Generate Report" >&nbsp;Generate Report</span>
			</td>
		</tr>
	</table>
	<br>
	<%  
		set rstJobResults = Nothing
		set objCnn = Nothing
	%>
	</form>
	<div id="ResultsTable">
	</div>
</body>
</html>
<script language="JavaScript">
	loadResults();
</script>
<%	End If %>

