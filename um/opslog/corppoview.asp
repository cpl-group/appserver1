<%@Language="VBScript"%>
<%option explicit%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="../../gEnergy2_Intranet/styles.css" type="text/css">
</head>
<script>
window.name = "corppoview"

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

<body bgcolor="#FFFFFF" text="#000000">
<table border=0 cellpadding="6" cellspacing="0" width="100%" bgcolor="#eeeeee" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
	<tr>
		<td><b>Approve/Reject Submitted RFs</b> &nbsp;|&nbsp; <a href="acctpoview.asp" style="color:#333366;">View Approved RFs</a> &nbsp;|&nbsp; 
		<a href="poviewdaterange.asp" style="color:#333366;">View All RFs</a>
		</td>
	</tr>
</table>
<%
Dim cnnMain, rs, sqlstr
Set cnnMain = Server.CreateObject("ADODB.connection")
Set rs = Server.CreateObject("ADODB.recordset")
cnnMain.Open getConnect(0,0,"intranet")

sqlstr = "select ltrim(str(po.Jobnum)) + '.' + ltrim(str(po.POnum)) as ponumber,po.* ,employees.[first name]+' '+employees.[last name] as req from po join employees on po.requistioner=substring(employees.username,7,20) where submitted = 1 and accepted = 0 order by podate desc"
'response.write sqlstr
'response.end
rs.Open sqlstr, cnnMain, 0, 1, 1

if rs.EOF then 
	%>  
	  
	<table border=0 cellpadding="3" cellspacing="0" width="100%">
		<tr> 
			<td>No RFs waiting for review.</td>
		</tr>
	</table>
	<%
Else
	x=0
	%>
	
	<table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
		<tr bgcolor="#dddddd"> 
			<td width="15%">RF #</td>
			<td width="15%">RF Date</td>
			<td width="30%">Requisitioner</td>
			<td width="25%">&nbsp;</td>
		</tr>
	  <%
		While not rs.EOF
			
			%><form name="form1" method="post" action="">
			<tr bgcolor="#ffffff"> 
				<td>
					<%if rs("question")="True" then %>
						<font color="#0033FF">*
					<%end if%>
					<a href="poview.asp?po=<%=rs("ponumber")%>&jid=<%=rs("jobnum")%>"><%=rs("ponumber") %></a>
					<input type="hidden" name="poid" value="<%=rs("id")%>">
				</td>
				<td><%=rs("podate")%></td>
				<td>
					<%=rs("req")%>
					<input type="hidden" name="job" value="<%=rs("requistioner")%>">
				</td>
				<td>  
					<input type="hidden" name="ponum" value="<%=rs("ponumber")%>">
					<input type="hidden" name="d" value="<%=rs("podate") %>">		<%
					'if allowgroups("GenergyAccounting") and cint(rs("approved")) = 0 then		%>
						<!--<input type="button" name="Button" value="APPROVE" onclick="processpo(poid.value, this.value,ponum.value,d.value)">-->		<%
					'end if
							
					'3/6/2009 N.Ambo cahnged group fro approving POs to be only 'Enterprise Exec'
					'if cint(rs("accepted")) = 0 and allowgroups("Genergy_Supervisors,Genergy_Corp") then' then	'3/6/2009 N.Ambo blobked of and replaced	
					if cint(rs("accepted")) = 0 and allowgroups("Enterprise Exec") then' then	%>							
						<input type="button" name="Button" value="ACCEPT" onclick="processpo(poid.value, this.value,ponum.value,d.value)">		<%
					end if
					
					if cint(rs("accepted")) = 0 and allowgroups("GenergyAccounting,Genergy_Supervisors,Genergy_Corp") then' then		%>
						<input type="button" name="Button" value="REJECT" onclick="processpo(poid.value, this.value,ponum.value,d.value)">		
						<a onclick="processpo('<%=rs("id")%>','Question','<%=rs("ponumber")%>','<%=rs("podate")%>')">
							<img src="question-ccc.gif" border="0">
						</a>			<%
					end if		%>

				</td>
			</tr>
			</form>
			<%
			rs.movenext
		Wend	%>
	</table>	<%
end if	%>
	
</body>
</html>