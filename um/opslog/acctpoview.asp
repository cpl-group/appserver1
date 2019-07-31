<%@Language="VBScript"%>
<%option explicit%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<script>
function RunImport()
{
document.location="PoOrder.asp";
}

function processpo(poid,action,ponum,podate) {
	if (action=="ACCEPT") {
		var poaction="accept"
	}else if (action=="REJECT"){
		var poaction="reject"
	} else if (action=="APPROVE"){
		var poaction="approve"
	}else{
		var poaction="question"
	}   
	var temp = "processpo1.asp?poid=" + poid + "&poaction=" + poaction + "&ponum=" + ponum + "&podate="+ podate
	window.open(temp,"", "scrollbars=yes,width=600, height=300, status=no, menubar=no" );
}



</script>
<link rel="Stylesheet" href="../../gEnergy2_Intranet/styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table border=0 cellpadding="6" cellspacing="0" width="100%" bgcolor="#eeeeee" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
  <tr>
    <td><a href="corppoview.asp" style="color:#333366;">Approve/Reject Submitted RFs</a> &nbsp;|&nbsp; <b>View Approved RFs</b> &nbsp;|&nbsp; 
	 	<a href="poviewdaterange.asp" style="color:#333366;">View All RFs</a>
    </td><td align="right">
	<%if  allowgroups("gAccounting,IT Services") then %>
	 <input type="submit" name="PO" value="Run Purhcase Order" onClick="return RunImport();"></td>
 	<%End if%>
  </tr>
</table>
<%
dim flag
flag=request("POflag")
if flag ="1" then
%>
<script>alert('Po File Created')</script>
<%
end if

Dim cnn1, rst1,rst2, sqlstr
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")

dim vendorSelect
rst1.open "select * from companycodes where active = 1 and code <> 'AC' order by name", getConnect(0,0,"intranet")
do until rst1.eof
	vendorSelect = vendorSelect & "SELECT [name], vendor, '"&rst1("code")&"' as comp FROM ["&rst1("code")&"_MASTER_APM_VENDOR] UNION all "
	rst1.movenext
loop
rst1.close
vendorSelect = "(SELECT distinct * FROM (" & vendorSelect
vendorSelect = left(vendorSelect,len(vendorSelect)-10) & ") v)"
sqlstr = "select ltrim(str(po.Jobnum)) + '.' + ltrim(str(po.POnum)) as ponumber,po.*, employees.[first name]+' '+employees.[last name] as req, case when po.vid<>'0' then vs.name else po.vendor end as vendorname from po INNER JOIN master_job m ON po.jobnum=m.id join employees on po.requistioner=substring(employees.username,7,20) LEFT JOIN "&vendorSelect&" vs ON vs.vendor=vid and vs.comp=m.company WHERE accepted = 1 and closed=0 order by podate desc"
'response.write sqlstr
rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.EOF then 
%>	
	
<table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr> 
    <td>No approved RFs waiting.</td>
  </tr>
</table>
<%
Else
x=0
%>


<table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
  <tr bgcolor="#dddddd"> 
    <td width="12%" height="2">RF #</td>
    <td width="15%" height="2">RF Date</td>
    <td width="23%" height="2">Vendor</td>
    <td width="19%" height="2">Requisitioner</td>
    <td width="13%" height="2">&nbsp;</td>
  </tr>
  <%
	While not rst1.EOF
		%>
		<form name="form1" method="post" action="">
			<tr bgcolor="#ffffff"> 
				<td width="12%"><a href='<%="poview.asp?po=" & rst1("ponumber") %>'><%=rst1("ponumber") %></a></td>
				<td width="15%"><%=rst1("podate") %></td>
				<td width="23%"><%=rst1("vendorname")%></td>
				<td width="19%"><%=rst1("req")%> 
					<input type="hidden" name="job" value="<%=rst1("requistioner")%>">
				</td>
				<td width="13%" align="center" nowrap> 
					<input type="hidden" name="id1" value="<%=rst1("id")%>">

					<%
					if ((not rst1("approved")) and allowgroups("Genergy Accounting")) then		%>
												
						<input type="button" name="Button" value="APPROVE" onclick="processpo(id1.value, this.value,id1.value,'<%=rst1("podate")%>')">
						<input type="button" name="Button" value="REJECT" onclick="processpo(id1.value, this.value,id1.value,'<%=rst1("podate")%>')">		
						<a onMouseOver="this.style.cursor='hand'" onclick="processpo('<%=rst1("id")%>','Question','<%=rst1("ponumber")%>','<%=rst1("podate")%>')">
							<img src="question-ccc.gif" border="0">
						</a>		<%
					end if		%>
				</td>
			</tr>
		</form>
		<%
		rst1.movenext
	Wend
  end if
  %>
</table>
</body>
</html>
