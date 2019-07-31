<%option explicit%>
<!-- #include virtual="/genergy2/secure.inc" -->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
'http params
'adodb vars

dim cnn, rs, sql, rsMore, sqlMore,order
set cnn = server.createobject("ADODB.Connection")
set rs = server.createobject("ADODB.Recordset")
set rsMore = server.createobject("ADODB.Recordset")
' open connection
cnn.open getConnect(0,0,"intranet")
cnn.CursorLocation = adUseClient

dim matchingJobs
matchingJobs = request("matchingjobs")
order = request("orderjobs")

if isempty(matchingJobs) or matchingJobs = "" then
	%>Page did not recieve proper parameters:  matchingJobs was null.<%
	response.End()
end if

sql = "select * from Master_job m left join (select jobno, isnull(sum(invoice_amt),0) as amtbilled from (select distinct jobno, invoice_date, invoice_amt from invoice_submission ) s group by s.jobno) inv on inv.jobno = m.id where id in " & matchingJobs & " order by " & order
'response.write sql
'response.end
dim jtd_work_billed_total, jtd_cost_total, profit_total, wip_total
jtd_work_billed_total = 0
jtd_cost_total = 0
profit_total = 0
wip_total = 0

'response.Write("<!--"&sql&"-->")
rs.open sql,cnn,1
if rs.recordcount > 1 then 

	%>
	<html>
	<head>
	<title>Job Search</title>
	<link rel="Stylesheet" href="../../styles.css" type="text/css">
	<style type="text/css">
		<!--
		BODY {
		SCROLLBAR-FACE-COLOR: #dddddd;
		SCROLLBAR-HIGHLIGHT-COLOR: #ffffff;
		SCROLLBAR-SHADOW-COLOR: #eeeeee;
		SCROLLBAR-3DLIGHT-COLOR: #999999;
		SCROLLBAR-ARROW-COLOR: #000000;
		SCROLLBAR-TRACK-COLOR: #336699;
		SCROLLBAR-DARKSHADOW-COLOR: #333333;
		}
		.printtable td { border-bottom:1px solid #cccccc; }
		-->
	</style>
	</head>
	
	<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#ffffff" onload="print();">
      
	
<table border=0 cellpadding="2" cellspacing="0" class="printtable" style="border:1px solid #cccccc;">
  <tr> 
    <td><b>Status</b></td>
    <td>&nbsp;</td>
    <td><b>Job Number</b></td>
    <td>&nbsp;</td>
    <td><b>Opened</b></td>
    <td><b>Customer</b></td>
    <td>&nbsp;</td>
    <td><b>Job Address</b></td>
    <td>&nbsp;</td>
    <td><b>Description</b></td>
    <td>&nbsp;</td>
    <td><b>Project Manager</b></td>
    <td>&nbsp;</td>
    <td align="center"><b>Amount Billed</b></td>
    <td align="center"><b>Contract Amount</b></td>
    <td>&nbsp;</td>
    <td align="center"><b>Percent Complete</b></td>
    <td>&nbsp;</td>
    <td align="center"><b>Work Billed to Date</b></td>
    <td>&nbsp;</td>
    <td align="center"><b>Cost to Date</b></td>
    <td>&nbsp;</td>
    <td align="center"><b>Profit</b></td>
    <td>&nbsp;</td>
    <td align="center"><b>WIP</b></td>
  </tr>
  <% 
		dim jtd_work_billed, jtd_cost, profit, wip, amt_1_total, amt_1, amtbilled, totamtbilled

				
		while not rs.EOF
			amt_1 = "N/A"
			jtd_work_billed = "N/A"
			jtd_cost = "N/A"
			profit = "N/A"
			wip = "N/A"
				sqlMore = "select c.jtd_work_billed,c.jtd_cost,c.jtd_work_billed-jtd_cost as profit,(" & rs("amt_1") & " * " & rs("percent_complete") & " /100)- jtd_work_billed as wip from " & rs("company")&"_master_job c where c.job='" & rs("job") & "'"
				'response.write sqlMore
				'response.end
				rsMore.open sqlMore, cnn
				if not rsMore.eof then
					if isnumeric(rs("amt_1")) then amt_1=formatcurrency(rs("amt_1"),0) else amt_1=rs("amt_1")  end if 
					jtd_work_billed = roundH(rsMore("jtd_work_billed"))
					jtd_cost = roundH(rsMore("jtd_cost"))
					profit = roundH(rsMore("profit"))
					wip = roundH(rsMore("wip"))
				else 
					if isnumeric(rs("amt_1")) then amt_1=formatcurrency(rs("amt_1"),0) end if 
				end if
				rsMore.close
				if isnumeric(rs("amtbilled")) then 
						amtbilled=formatcurrency(rs("amtbilled"),0) 
						totamtbilled = totamtbilled + rs("amtbilled")
				else 
						amtbilled="NA" 
				end if
			
			 dim rst666,projmanager
	  			set rst666 = server.createobject("ADODB.recordset")
	  		rst666.Open "select Firstname,lastname from managers where mid='"&rs("project_manager")& "'" , cnn
				if not rst666.EOF then
			projmanager = rst666("lastname") & ", " & rst666("Firstname") 
				end if
				rst666.close
		
			%>
  <tr bgcolor="#ffffff" valign="top"> 
    <td nowrap><%=rs("status")%> 
      <%if lcase(trim(rs("rfp"))) = "true" then response.write " - RFP" end if%>
      &nbsp;</td>
    <td>&nbsp;</td>
    <td nowrap><%=rs("job")%>&nbsp;</td>
    <td>&nbsp;</td>
    <td><%=rs("actual_start_date")%></td>
    <td><%=rs("customer_name")%>&nbsp;</td>
    <td>&nbsp;</td>
    <td><%=rs("address_1")%>&nbsp;</td>
    <td>&nbsp;</td>
    <td width="20%"><%=rs("description")%>&nbsp;</td>
    <td>&nbsp;</td>
    <td width="10%" nowrap><%=projmanager%>&nbsp;</td> 
	<td>&nbsp;</td>
    <td align="right"><%=amtbilled%></td>
    <td align="right"><%=amt_1%>&nbsp;</td>
    <td>&nbsp;</td>
    <td align="right"><%=rs("percent_complete")%>&nbsp;</td>
    <td>&nbsp;</td>
    <td align="right"><%=jtd_work_billed%>&nbsp;</td>
    <td>&nbsp;</td>
    <td align="right"><%=jtd_cost%>&nbsp;</td>
    <td>&nbsp;</td>
    <td align="right" width="9%" nowrap><%=profit%>&nbsp;</td>
    <td>&nbsp;</td>
    <td align="right" width="8%" nowrap><%=wip%>&nbsp;</td>
  </tr>
  <% 
			if amt_1 <>"N/A" then amt_1_total = amt_1_total + rs("amt_1")
			if jtd_work_billed<>"N/A" then jtd_work_billed_total = jtd_work_billed_total + jtd_work_billed
			if jtd_cost <>"N/A" then jtd_cost_total = jtd_cost_total + jtd_cost
			if profit <>"N/A" then profit_total = profit_total + profit
			if wip <>"N/A" then wip_total = wip_total + wip
			rs.movenext
		wend
	%>
  <tr bgcolor="#eeeeee" valign="top"> 
    <td colspan="13">Total</td>
    <td align="right"><%=formatcurrency(totamtbilled,0)%></td>
    <td align="right"><%=formatcurrency(amt_1_total,0)%>&nbsp;</td>
    <td>&nbsp;</td>
    <td align="right">&nbsp;</td>
    <td>&nbsp;</td>
    <td align="right"><%=jtd_work_billed_total%>&nbsp;</td>
    <td>&nbsp;</td>
    <td align="right"><%=jtd_cost_total%>&nbsp;</td>
    <td>&nbsp;</td>
    <td align="right" width="9%" nowrap><%=profit_total%>&nbsp;</td>
    <td>&nbsp;</td>
    <td align="right" width="8%" nowrap><%=wip_total%>&nbsp;</td>
  </tr>
</table>	 
	
	</body>
	</html>
	<%
	rs.close
end if

function roundH( num)
	roundH = ((round(num * 100))/ 100)
end function%>