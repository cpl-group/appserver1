<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<html>
<head>


<title>Reporting Search</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
try{top.applabel("Maintenance Reports");}catch(exception){}
function fillup(bldg,portid){
	document.location="reportingindex.asp?bldg=" + bldg+"&portid="+portid
}
function pickreport(bldg,floor1){
	document.location="reportingindex.asp?bldg=" + bldg+"&floor="+floor1
}

function report(job) {
document.frames.reports.location=job
}

</script>
<link rel="Stylesheet" href="/genergy2/styles.css">
</head>
<%

msg = Request.querystring("msg")
bldg = Request("bldg")
portid = Request("portid")


			if isempty(msg) then
				msg="Please select a floor and type of report (lamping or ballast), then click <i>Show Report</i>"
			end if
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2= Server.CreateObject("ADODB.recordset")


cnn1.Open getconnect(0,0,"engineering")


		
%>
<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post">
  <table width="100%" cellpadding="3" cellspacing="0" border="0" style="border:1px solid #ffffff">
    <tr> 
      <td align="left" bgcolor="#FFFFFF" nowrap><span class=standardheader><font color="#000000">Select 
        Building :</font></span> <select name="select">
          <option value="#">123 Main Street</option>
        </select> </td>
      <td align="right" bgcolor="#FFFFFF"><input type="button" name="print" value="Print Current View" onClick="javascript:document.frames.reports.focus();document.frames.reports.print()" class="standard"></td>
    </tr>
  </table>

<table width="100%" border="0" cellpadding="3" cellspacing="0" style="border:1px solid #ffffff">
	<tr>
		<td bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif"><span class="standard"><%=msg%></span></font><br></td>
	</tr>
	<tr>
	
		<td bgcolor="#eeeeee">
		<% if trim(portid)<>"" then %>
		<input type="hidden" name="portid" value="<%=portid%>">
        <select name="typebox" size="1" onChange=fillup(this.value,portid.value)>
			<option value="">Select Building</option>
          <%
		  sqlstr = "select distinct bldgname,id from facilityinfo where portfolioid='"&portid&"' "
   			rst1.Open sqlstr, cnn1, 0, 1, 1
			do until rst1.eof
		  %>
		  <option value="<%=rst1("id")%>" <%if bldg=trim(rst1("id"))then response.write "selected"%>><%=rst1("bldgname")%></option>
		  <%		
					rst1.movenext
					loop			
				%>
        </select>
	<%else%>
	<input type="hidden" name="typebox" value="<%=bldg%>">
		<%end if%>
        <select name="findvar" >
		<option value="All">All Floors</option>
          <%
			sqlstr = "select distinct ID,FLOOR from FLOOR where bldg='"&request("bldg")&"' order by floor"
   			rst2.Open sqlstr, cnn1, 0, 1, 1
				do until rst2.eof	
		%>
          <option value="<%=rst2("ID")%>"><%=rst2("floor")%></option>
          <%
					rst2.movenext
					loop
				%>
        </select>
		
		<select name="whichreport">
		<option value="lampreport">Lamping Report</option>
		<option value="ballastreport">Ballast Report</option>
		</select>
        <input type="hidden" name="lr" value="lamping report.asp?"> <input type="hidden" name="br" value="ballastreport.asp.asp"> 
        <!--input type="button" name="Lamping" value="Lamping Report" onclick="report('lampreport.asp?bldgnum='+typebox.value+'&floor='+findvar.value)" class="standard"-->
        <!--input type="button" name="ballast" value="Ballast Report" onclick="report('ballastreport.asp?bldgnum='+typebox.value+'&floor='+findvar.value)" class="standard"-->
        <input type="button" name="showreport" value="Show Report" onClick="report(whichreport.value+'.asp?bldgnum='+typebox.value+'&floor='+findvar.value)" class="standard">	
      </td>
	</tr>
	</table>
	<table width="100%" cellpadding="3" cellspacing="0" border="0" style="border:1px solid #ffffff">	
    <tr> 
      <td bgcolor="#cccccc">&nbsp; </td>
      <td bgcolor="#cccccc" align="right">&nbsp;</td>
    </tr>
  </table>
  <font face="Arial, Helvetica, sans-serif"> </font> 
</form>
<IFRAME name="reports" width="100%" height="78%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0></IFRAME> 
</body>
</html>