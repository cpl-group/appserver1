<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>RF</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/javascript">
function fillup(typebox){
	document.location="poindex.asp?typebox=" + typebox
}

function searchjob(typebox, searchitem) {
	//alert(typebox)
	//alert(searchitem)
	if (((typebox=="jobnum") | (typebox=="ponum")) && (searchitem=="")) {
		alert("Please enter a number to look for");
		return false;
	}
	
	var temp
	if (typebox=="ponum"){
		temp="poview.asp?po=" + searchitem+"&printview="+ document.forms.form1.printview.value
	} else {
		if (typebox=="date"){
			temp="posearch.asp?select=" + typebox + "&findvar=var&fromdate=" + document.forms.form1.fromdate.value + "&todate=" + document.forms.form1.todate.value+"&printview=" + document.forms.form1.printview.value
		} else{
			temp="posearch.asp?select=" + typebox + "&findvar="+searchitem+"&printview=" + document.forms.form1.printview.value
		}
	}
	if (document.forms.form1.printview.value == "yes"){
		window.open(temp,'searchresults','toolbar=no,scrollbars=yes')
	} else {
	document.frames.oplog.location=temp
	}

}
function report(spec, job) {
	document.frames.oplog.location=spec
}

function timesheetjob(typebox, job){
	//alert(typebox)
	//alert(document.forms[0].findvar.value)
	var temp
	if(typebox =="Job Number"){
		if(job == ""){
			alert("Please Enter Job Number")
		}else{
			temp="timesheetsearch.asp?job="+job
			document.frames.oplog.location=temp
		}
	} else{
		temp="null.htm"
		document.frames.oplog.location=temp
	}

}

</script>
<link rel="Stylesheet" href="../../gEnergy2_Intranet/styles.css" type="text/css">
</head>

<%
ReDim Ay(6)
ReDim By(6)
Ay(0) = "jobnum"
Ay(1) = "vendor"
Ay(2) = "ponum"
Ay(3) = "description"
Ay(4) = "requistioner"
Ay(5) = "date"
By(0) = "Job Number"
By(1) = "Vendor"
By(2) = "RF Number"
By(3) = "Description"
By(4) = "Requisitioner"
By(5) = "Date"

msg = Request.querystring("msg")
typebox = Request("typebox")
if isempty(msg) then
	msg=""
end if
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")

%>
<body bgcolor="#ffffff" text="#000000">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
	<tr>
		<td bgcolor="#6699cc" colspan="2">
			<table border=0 cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td><span class="standardheader">Requisition Forms</span></td>
					
          <td align="right">&nbsp; </td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td colspan="2" bgcolor="#eeeeee"><%=msg%></td>
	</tr>
	<form name="form1" method="post">
		<tr>
			<td bgcolor="#eeeeee" style="border-bottom:1px solid #cccccc;">
				<table border=0 cellpadding="3" cellspacing="0">
					<tr> 
						<td>Search for requisition form by:</td>
						<td>
							<select name="typebox" size="1" onChange=fillup(this.value)>
								<%
								for i=0 to 5
									if Ay(i) = typebox then%>
										<option value="<%=typebox%>" selected><%=By(i)%></option><%
									else %>
										<option value="<%=Ay(i)%>"><%=By(i)%></option><%
									end if
								Next
								%>
							</select>
						</td><%
						if (typebox = "requistioner") then%>
							<td>
								<select name="findvar">
									<%
									'sqlstr = "select [first name]+' '+ [last name] as name from employees where active=1"
									dim name
									sqlstr = "select [first name], [last name] from employees where active=1 order by [last name], [first name]"
									rst1.Open sqlstr, cnn1, 0, 1, 1
									if not rst1.eof then
										do until rst1.eof 
											%><option value="<%=rst1("first name") + " " + rst1("last name")%>"><%=rst1("last name") + ", " + rst1("first name")%></option><%
											rst1.movenext
										loop
									end if
									%>
								</select>
							</td><%
						elseif (typebox = "vendor") then
							%>
							<td>
								<select name="findvar">
									<%
									sqlstr = "select distinct name as vendor from (select distinct name from gy_master_apm_vendor union select distinct name from ge_master_apm_vendor) a order by name"
									rst1.Open sqlstr, cnn1, 0, 1, 1
									if not rst1.eof then
										do until rst1.eof%>
											<option value="<%=rst1("vendor")%>"><%=left(rst1("vendor"),40)%></option><%
											rst1.movenext
										loop
									end if
									%>
								</select>
							</td><%
						elseif (typebox = "date") then%>
							<td>
								From Date&nbsp;<input type="text" name="fromdate" value="<%=cdate(date()-30)%>" size="13%">&nbsp;&nbsp;&nbsp;&nbsp;
								To Date&nbsp;<input type="text" name="todate" value="<%=cdate(date())%>" size="13%">
								<input type="hidden" name="findvar" value="">
							</td>
							<%
						else%>
							<td><input type="text" name="findvar" size="30" maxlength="50"></td>
							<%
						end if%>
						<td><input type="button" name="Submit3" value="Find" onClick="printview.value='no';searchjob(typebox.value,findvar.value);"></td>

					</tr>
				</table>
			</td>
			<td align="right" bgcolor="#eeeeee"  style="border-bottom:1px solid #cccccc;">
				<script>
					function printView(){
						document.forms.form1.printview.value='yes';
						searchjob(document.forms.form1.typebox.value,document.forms.form1.findvar.value);
					}
				</script>
				<input type="hidden" name="printview" value="no">
        <input type="hidden" name="np" value="newpo.asp"> <input type="button" name="job" value="New RF"  onClick="report(np.value)"> 
        <input type="button" name="print" value="Print Search Results" onClick="javascript:printView();"> 
      </td>
		</tr>
	</form>
</table>
	
<IFRAME name="oplog" width="100%" height="90%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0></IFRAME> 

</body>
</html>