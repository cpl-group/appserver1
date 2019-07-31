<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim cnn, rstUser, rst, strsql, Uname,TotalHrs,TotalOverTime,fromTime, toTime, andWhere, orderby, viewall
					
if not secureRequest("UsrName")="" then
	Uname=secureRequest("UsrName")
end if

set cnn = server.createobject("ADODB.connection")
set rstUser = server.createobject("ADODB.recordset")
set rst = server.createobject("ADODB.recordset")

cnn.open getConnect(0,0,"intranet")

dim printFlag
printFlag	= secureRequest("printview")
fromTime 	= secureRequest("fromTime")
toTime 		= secureRequest("toTime")
orderby		= secureRequest("orderby")
viewall		= secureRequest("viewall")

if secureRequest("printview") = "" then
	printFlag = "false"
end if

if secureRequest("orderby")="" then	
	orderby=0
end if

if secureRequest("viewall")="" then	
	viewall=0
end if

if printFlag = "true" then
	fromTime = request.QueryString("fromTime")
	toTime= request.QueryString("toTime")
end if


if Trim(fromTime) <> "" and viewall=1 then 
	if Trim(toTime) <> "" then
		andWhere = " and date between '" & fromTime & "' and '" & toTime & "' "
	else
		andWhere = " and date >= '" & fromTime & "' " 
	end if
else
	andWhere = " "
end if

strsql = "Select [date], JobNo, description, hours, overT from times where matricola='" & Uname & "' " & andWhere 

if orderby = 0  then	
	strsql = strsql & " Order By [date] desc "
else
	strsql = strsql & "  Order By JobNo "
end if		


rst.open strsql,cnn
Dim sUser

if trim(fromTime) = "" then 
	fromTime=Month(Now) & "/01/" & Year(Now) 
End if
%>
<html>
	<head>
		<title>Search Job
			<%if printFlag = "true" then %> 
			<%end if%>
		</title>
		<script language="JavaScript1.2">
			function openwin(url,mwidth,mheight){
				window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
			}
		</script>
		<link rel="Stylesheet" href="../../styles.css" type="text/css">   
	</head>
	<body bgcolor="#eeeeee" <%if printFlag = "true" then%> onLoad="window.print()" <%end if%>>
		<form name="frmSearchJob" method="post" action="SearchAllJob.asp">
			<% if printFlag = "false" then %>
				<table width="100%" border="0" cellpadding="3" cellspacing="0" align="center">
					<tr bgcolor="#3399cc">
						<td><font color='white'>Search Job by User </font></td>
					</tr>
				</table>
				 &nbsp;
				<div id="Datasourceinfo" style="BORDER-RIGHT: #cccccc 1px solid; PADDING-RIGHT: 3px; BORDER-TOP: #cccccc 1px solid; DISPLAY: inline; PADDING-LEFT: 3px; PADDING-BOTTOM: 3px; BORDER-LEFT: #cccccc 1px solid; WIDTH: 98%; PADDING-TOP: 3px; BORDER-BOTTOM: #cccccc 1px solid"> 
					<table width="100%" border="0" cellpadding="3" cellspacing="0">
						<tr> 
						  <td align="Left">Select User:</td>
						  <td align="left">
							<select style="WIDTH: 202px; HEIGHT: 22px" name="ddlUser" onChange="document.frmSearchJob.UsrName.value=this.value">
								<option value="All">Select Users</option>
								<%Set rstUser = Server.CreateObject("ADODB.recordset")
									rstUser.Open "Select Distinct(RTrim(LTrim(matricola))) as [User] FROM times Order by [User] ", cnn
								do until rstUser.eof
									sUser=Split(rstUser("User"),"\")
									%>
									<option value="<%=rstUser("User")%>" <%if trim(Uname)=trim(rstUser("User")) then response.write " SELECTED"%>>
										<%=sUser(1)%>
									</option>
									<%
									rstUser.movenext
								loop
								rstUser.close%>
							</select>
						  </td>
						  <td>
							<input type="button" name="PrintViewUser" value="Print Frame" onClick="openwin('SearchAllJob.asp?UsrName='+document.frmSearchJob.ddlUser.value+'&fromTime='+document.frmSearchJob.fromTime.value+'&toTime='+document.frmSearchJob.toTime.value+'&printview=true', 600, 600)">
						  </td>	
						</tr>
						<tr>
							<td>
								Adjust Timeframe to View: 
							</td>
							<td>
								From Date - 
								<input type="text" name="fromTime" value="<%=fromTime%>">
								To Date - 
								<input type="text" name="toTime" value="<%=toTime%>">
								&nbsp;&nbsp;
								<input type="submit" name="Submit" value="View" onClick="document.frmSearchJob.viewall.value = 1">
							</td>
							<td>
								<%if orderby = 0 then %>
									<input type="submit" name="ByJob" value="Sort By Job" onClick="document.frmSearchJob.orderby.value = 1">
								<%else%>
									<input type="submit" name="ByDate" value="Sort By Date" onClick="document.frmSearchJob.orderby.value = 0">
								<%end if%>	
							</td>	
						</tr>
					</table>
				</div>
				<input type="hidden" name="UsrName" value="<%=Uname%>">
				<input type="hidden" name="orderby" value="<%=orderby%>">
				<input type="hidden" name="viewall" value="<%=viewall%>">
				<input type = "hidden" name = "printview" value = "false">	
			<% end if %>
			&nbsp;
			<table border=0 cellpadding="0" cellspacing="0" width="100%" bgcolor="#eeeeee" style="border-right:1px solid #cccccc;">
				<tr>
					<td align="center"> 
						<table border=0 cellpadding="3" cellspacing="1" width="98%">
							<tr bgcolor="#dddddd"> 
								<td width="10%">Date</td>
								<td width="10%">Job Id</td>
								<td width="60%">Description</td>
								<td width="10%">Hours</td>
								<td width="10%">Overtime</td>
							</tr>
						</table>
							
						 <%if printFlag = "false" then %>
							<div style="width:98%; overflow:auto; height:540px;border-right:1px solid #cccccc;">
						<%end if%>  
 
						<table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#dddddd">
							<%while not rst.EOF%>
							<tr bgcolor="#ffffff" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="document.location='jobtime.asp?jid=<%=rst("JobNo")%>'">
						 		<td width="10%"><%=rst("date")%> </td>
  								<td width="10%"><%=rst("JobNo")%> </td>
								<td width="60%"><%=rst("description")%></td>
								<td width="10%" align="right"><%=formatnumber(rst("hours"),2)%></td>
								<td width="10%" align="right"><%=formatnumber(rst("overt"),2)%></td>
							</tr>
							<% 
							rst.movenext
							wend 
							%>
						</table>
						
						 <%if printFlag = "false" then %>
							</div>
						<%end if
						
						 rst.Close
						 set cnn = nothing 
						 %>
					 </td>
				</tr>
			</table>
		</form>
	</body>	
</html>	