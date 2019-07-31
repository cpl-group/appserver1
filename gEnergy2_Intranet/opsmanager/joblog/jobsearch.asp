<%option explicit%>
<!-- #include virtual="/genergy2/secure.inc" -->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
'http params
dim crdate,j,c,avg,wip,tcost,OPENPO,link,amt_paid,aTable,sql, prefix,hours, company, search,order, tcolor,status,jtype
dim print,restore,search0,order0,showprint,custid, unstarted, inprogress, closed, statusset, rfp,rfpsql,filter_fixed, pm

c= request("cc")
search 	= request("search")
if search = "Insert Search Text" then
	search = ""
end if
search0=search
order = request("order")
order0=order

unstarted	=request("unstarted")
inprogress	=request("inprogress")
closed		=request("closed")
rfp 		=request("rfp")
filter_fixed=request("filter_fixed")
pm = request("pm")

if rfp <> "1" then 
	rfp = "0"
end if

statusset = "'" & unstarted& "','" & inprogress &"','"& closed &"'"
				
jtype=request("jtype")
company=request("company")
custid=request("custid")
restore=1

if request("print")="yes" then
	print=true
else
	print=false
end if

if order="status" then order=order + " desc"
if order="job" then order="id"  'instead of left based on id
'adodb vars
dim cnn, cmd, rs, pstatus,ptype,commaloc
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")
' open connection
cnn.CommandTimeout = 60*5
cnn.open getConnect(0,0,"intranet")
cnn.CursorLocation = adUseClient

Dim param, tempstr, flagarray, invalid, xcriteria, pidx, typesegs, segment, segidx
invalid = False
xcriteria = 0
pidx = 0
sql = "SELECT top 800 isnull(tot_hours,0) as hours, * FROM Master_job m LEFT JOIN (SELECT isnull(SUM(Hours_bill),0) AS tot_hours, jobno FROM times GROUP BY jobno) hb ON hb.jobno=m.id left join (select jobno, isnull(sum(invoice_amt),0) as amtbilled from (select distinct jobno, invoice_date, invoice_amt from invoice_submission ) s group by s.jobno) inv on inv.jobno = m.id WHERE"

select case filter_fixed

case 1 
	sql = sql & " Billing_Method_1 not like '%Charge%' and Billing_Method_2 not like '%Charge%' and Billing_Method_1 not like '%Recurring%' and Billing_Method_2 not like '%Recurring%' and"
case 2 
	sql = sql & " percent_complete < 100 and"
case 3 
	sql = sql & " type not like '%ALL-OFFICE%' and type not like '%Internal%' and"
case 4 
	sql = sql & " Billing_Method_1 not like '%Charge%' and Billing_Method_1 not like '%Recurring%' and Billing_Method_2 not like '%Charge%' and Billing_Method_2 not like '%Recuring%' and percent_complete < 100 and type not like '%ALL-OFFICE%' and  type not like '%Internal%' and "
case 5
	sql = sql & " Billing_Method_1 not like '%Charge%' and Billing_Method_1 not like '%Recurring%' and Billing_Method_2 not like '%Charge%' and Billing_Method_2 not like '%Recurring%' and percent_complete > 99 and type not like '%ALL-OFFICE%' and type not like '%Internal%' and "
case 6
	sql = sql & " Billing_Method_1 not like '%Charge%' and Billing_Method_1 not like '%Recurring%' and Billing_Method_2 not like '%Charge%' and Billing_Method_2 not like '%Recurring%' and type not like '%ALL-OFFICE%' and type not like '%Internal%' and"
case 7
	sql = sql & " (Billing_Method_1 like '%Charge%' or Billing_Method_2 like '%Charge%' or Billing_Method_1 like '%Recurring%' or Billing_Method_2 like '%Recurring%') and " 
end select

If InStr(search, ",") Then
	flagarray = Split(search, ",")
	search = flagarray(UBound(flagarray))
	While pidx < UBound(flagarray) And Not invalid
		param = trim(flagarray(pidx))
		
		If LCase(param) = "unstarted" Or LCase(param) = "in progress" Or LCase(param) = "closed" Then
		
			If xcriteria Mod 2 = 0 Then
				If xcriteria = 0 Then
					sql = sql & " status='" & LCase(param) & "'"
				Else
					sql = sql & " and status='" & LCase(param) & "'"
				End If
				xcriteria = xcriteria + 1
				param=""
			Else
				invalid = True
			End If
			
		
		Else  ' try status
			rs.open "select distinct left(type,2) from master_job_types union select distinct substring(type,4,3) from master_job_types",cnn
			segment=""
			While not rs.eof
				segment=segment & "," & lcase(rs(0))
				rs.movenext
			Wend
			rs.close  ' end load of valid status segments
			typesegs = Split(segment, ",")
			While Len(param) > 1 And Not invalid  'should run twice max
				If Len(param) > 3 Then
					segment = Left(param, 2)
					param = Right(param, 3)
				ElseIf Len(param) > 2 Then
					segment = Left(param, 3)
					param = ""
				Else
					segment = Left(param, 2)
					param = ""
				End If
				
				For segidx = 0 To UBound(typesegs)
					If typesegs(segidx) = lcase(segment) Then
						segment = ""
					End If
				Next
				
				If segment <> "" Then
					invalid = True
					param=segment
				End If
			Wend   ' end run twice loop
			If Not invalid Then  ' modify sql
				If xcriteria < 2 Then
					param=trim(flagarray(pidx))
					
					If Len(param) = 6 Then
						If xcriteria = 0 Then
							sql = sql & " left(type,6)='" & flagarray(pidx) & "'"
						Else
							sql = sql & " and left(type,6)='" & flagarray(pidx) & "'"
						End If
					End If
					
					If Len(param) = 3 Then
						If xcriteria = 0 Then
							sql = sql & " right(left(type,6),3)='" & flagarray(pidx) & "'"
						Else
							sql = sql & " and right(left(type,6),3)='" & flagarray(pidx) & "'"
						End If
					End If
					
					If Len(param) = 2 Then
						If xcriteria = 0 Then
							sql = sql & " left(type,2)='" & flagarray(pidx) & "'"
						Else
							sql = sql & " and left(type,2)='" & flagarray(pidx) & "'"
						End If
					End If
					
					xcriteria = xcriteria + 2
				Else
					invalid = True
				End If  
			End If  ' sql is modified, or invalid
		End If    ' end of status matching
		pidx = pidx + 1  ' try another text flag
	Wend
			
Elseif statusset<>"'','',''" or jtype<>"" or company<>"AC" or custid<>"" then ' try pulldown flags
	if statusset<>"'','',''" then
		sql=sql&" status in ("&statusset&")"
		xcriteria=xcriteria+1
	end if
	
	if jtype<>"" then
		if xcriteria>0 then
			sql=sql&" and type_id='"&jtype&"'"
		else
			sql=sql&" type_id='"&jtype&"'"
		end if
		xcriteria=xcriteria+2
	end if
	
	'3/17/2008 N.Ambo added to search by project managers
	if pm<>"" then
		if xcriteria>0 then
			sql=sql&" and project_manager='"&pm&"'"
		else
			sql=sql&" project_manager='"&pm&"'"
		end if
		xcriteria=xcriteria+3
	end if
	
	if company<>"AC" then
		if xcriteria>0 then
			sql=sql&" and company='"&company&"'"
		else
			sql=sql&" company='"&company&"'"
		end if
	end if
	
	if custid<>"" then
		if xcriteria>0 or company<>"AC" then
			sql=sql&" and customer='"&custid&"'"
		else
			sql=sql&" customer='"&custid&"'"
		end if
	end if
	
	if xcriteria>0 or company<>"AC" or custid<>"" then
		restore=0
	end if
End If	 ' end pulldown search
		  
		
If Not invalid Then
	if rfp <> "1" then 
		rfpsql = " and rfp = 0 " 
	elseif statusset = "'','',''" then 
		rfpsql = " and rfp = 1 " 
	else 
		'rfpsql = "or rfp = 1"
	end if
		
	If xcriteria = 0 Or xcriteria = 2 Then
		If xcriteria = 2 or company<>"AC" or custid<>"" Then
			sql = sql & " and"
		End If
		if statusset = "'','',''" and rfp = "1" then
			sql = sql & " ((job like '%" & search & "%' or m.description like '%" & search & "%' or Address_1 like '%" & search & "%' or Address_2 like '%" & _
				search & "%' or customer_name like '%" & search & "%' or bldgnum like '%" & _
				search & "%'  or bldgnum in (select bldgnum from ( select  * from ["&Application("CoreIP")&"].dbCore.dbo.Buildings where bldgname like '%" & search & "%' or strt like '%" & search & "%' ) manBl ))"&rfpsql&") order by " & order
				if order <> "actual_start_date" then 
					sql = sql & ",actual_start_date" 
				else
					sql = sql & " desc"
				end if
		else
			sql = sql & " ((job like '%" & search & "%' or description like '%" & search & "%' or Address_1 like '%" & search & "%' or Address_2 like '%" & _
				search & "%' or customer_name like '%" & search & "%' or bldgnum like '%" & _
				search & "%'  or bldgnum in (select bldgnum from ( select * from ["&Application("CoreIP")&"].dbCore.dbo.Buildings where bldgname like '%" & search & "%' or strt like '%" & search & "%' ) manBl))) "&rfpsql&" order by " & order
				if order <> "actual_start_date" then 
					sql = sql & ",actual_start_date" 
				else
					sql = sql & " desc"
				end if
		end if
	End If

	
	If xcriteria = 1 Or xcriteria = 3 Then
		sql = sql & rfpsql & " and ((job like '%" & search & "%' or description like '%" & search & "%' or Address_1 like '%" & search & _
			"%' or Address_2 like '%" & search & "%'  or customer_name like '%" & _
			search & "%' or bldgnum like '%" & search & "%' or bldgnum in (select bldgnum from (select * from ["&Application("CoreIP")&"].dbCore.dbo.Buildings where bldgname like '%" & _
			search & "%' or strt like '%" & search & "%' ) manBl))  "&rfpsql&") order by " & order
				if order <> "actual_start_date" then 
					sql = sql & ",actual_start_date" 
				else
					sql = sql & " desc"
				end if
	End If
End If
		
'response.Write sql
'response.end
if restore=1 then
	status=""
	jtype=""
	company=""
end if
				
				
if invalid then
	if param="" then
		param="too many"
	end if
else
	rs.open sql,cnn
	if rs.EOF then
		invalid=True
	end if
end if
		
		
if invalid then 
	response.write "<html><head><link rel=""Stylesheet"" href=""../../styles.css"" type=""text/css"">	</head><body bgcolor=ffffff>"
	
	if param="" then
		%>
		<div style="padding:10px;">No records found<br><br>
		<a href="javascript:history.back()">
		<img src="/um/opslog/images/btn-back.gif" onmouseover="this.style.border='1px solid #6699cc'" 
			onmouseout="this.style.border='1px solid #ffffff'" style="border:1px solid #ffffff" border="0">
		</a>
		</div>
		<%
	else %>
		<div style="padding:10px;"><%=param%> is not a valid status/job type</div> <%
	end if %>
	
	</body>
	</html> <%
	set cnn = nothing 
	response.end
end if
		
if rs.recordcount > 1 then 

	%>
	<html>
	<head>
	<title>Job Search</title>
	<script language="JavaScript" type="text/javascript">
	//<!--
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
		}//-->
	</script>
	<link rel="Stylesheet" href="../../styles.css" type="text/css">
	<%if not print then%>		
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
		
		td.red {color: red}
		-->
		</style>
		</head>
				
		<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#ffffff">
	<%end if%>
	
	<form name="toJobPrint"  target="_blank" action="jobprint.asp" method="post">
		<input type="hidden" id="matchingJobs" name="matchingJobs" value="(">
		<input type="hidden" id="orderjobs" name="orderjobs" value="<%=order%>">
	</form>
		
	
<table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#dddddd">
  <tr bgcolor="#6699cc"> 
    <td colspan="10" align="center" nowrap><span class="standardheader"><b>&nbsp;&nbsp;</b></span></td>
    <td colspan="40" nowrap><span class="standardheader">Job #</span></td>
    <td colspan="30" align="center"><span class="standardheader">Opened</span>&nbsp;</td>
    <td colspan="50" ><span class="standardheader">Contract Amount</span></td>
    <td width="5%" align="center"><span class="standardheader">Percent Conplete</span></td>
    <td width="5%"><span class="standardheader">Amount Billed</span></td>
    <td width="23%"><span class="standardheader">Customer</span></td>
    <td width="21%"><span class="standardheader">Job Address</span></td>
    <td width="23%"><span class="standardheader">Description</span></td>
    <td width="13%"><span class="standardheader">Project Manager</span></td>
    <td align="center"><span class="standardheader">Hours Posted</span></td>
    <td align="center" width="2%" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <% if not print then %>
</table>
			
			
  <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#dddddd">
    <% end if
		
		Dim InProgressCount, UnstartedCount, ClosedCount, RFPCount
		
		InProgressCount = 0
		UnstartedCount 	= 0
		ClosedCount		= 0
		RFPCount		= 0
		
		while not rs.EOF
			%>
    <script>document.getElementById("matchingJobs").value = document.getElementById("matchingJobs").value + "<%=rs("id")%>"</script>
    <%
			Select Case lcase(trim(rs("status")))
				case "in progress"
					tcolor = "#66ff66"
					if lcase(trim(rs("rfp"))) <> "true" then 								
						InProgressCount = InProgressCount + 1
					end if
				case "unstarted"
					tcolor = "#FFcc00"
					if lcase(trim(rs("rfp"))) <> "true" then 
						UnstartedCount = UnstartedCount + 1
					end if
				case "closed"
					tcolor = "#cc0033"
					if lcase(trim(rs("rfp"))) <> "true" then 
						ClosedCount    = ClosedCount + 1
					end if
				case else
			end select
			if lcase(trim(rs("rfp"))) = "true" then 
				rfpcount = rfpcount + 1
			end if
			%>
    <tr bgcolor="#ffffff" valign="top" <% if not print then %>onMouseOver="this.style.backgroundColor = 'lightgreen'" 
				style="cursor:hand" onMouseOut="this.style.backgroundColor = 'white'" onClick="javascript:document.location='viewjob.asp?jid=<%=rs("id")%>'"<% end if %>> 
      <% dim output
				if not print and lcase(trim(rs("rfp"))) = "true" then
					output = "<b>R</b>"
					if trim(rs("status")) = "Closed" then
						output = "<font color='white'>" & output & "</font>"
					end if
				end if
				if print then
					output = rs("status")
					if lcase(trim(rs("rfp"))) = "true" then
						output = output & " - RFP"
					end if
				end if
				%>
      <td bgcolor="<%=tcolor%>" colspan="10" align="center" nowrap><%=output%></td>
      <%output=""%>
      <td colspan="40" nowrap><%=rs("job")%></td>
      <td colspan="30" align="right"> 
        <%if isdate(rs("actual_start_date")) then
						response.write(formatdatetime(rs("actual_start_date"),vbShortDate))
					else
						response.write(rs("actual_start_date"))
					end if%>
      </td>
      <td colspan="50" align="right"> 
        <%if isnumeric(rs("amt_1")) then %>
        <%=formatcurrency(rs("amt_1"))%> 
        <%end if%>
        &nbsp;</td>
      <td width="5%"align="right"><%=rs("percent_complete")%>%</td>
      <td width="5%"align="right"><% if isnull(rs("amtbilled")) then %>NA<%else Response.write formatcurrency(rs("amtbilled"),0) end if %></td>
      <td width="23%"><%=rs("customer_name")%></td>
      <td width="21%"><%=rs("address_1")%></td>
      <td width="23%"><%=rs("description")%></td>
      <%
	  dim rst666,projmanager
	  set rst666 = server.createobject("ADODB.recordset")
	  rst666.Open "select Firstname,lastname from managers where mid='"&rs("project_manager")& "'" , cnn
		if not rst666.EOF then
		projmanager = rst666("lastname") & ", " & rst666("Firstname") 'used to use Ando rs("pm_last") &" "& rs("pm_first")
		end if
		rst666.close
		'response.write projmanager
		'response.end
		if IsNull(rs("pm_last")) then
			projmanager = ""
		End If 
		%>
	   <td width="13%"><%=projmanager%></td>
	 <!-- <td width="13%"><%'=trim(rs("pm_last"))%>, <%'=trim(rs("pm_first"))%></td>-->
      <td align="right"> 
        <%
					'dim rstTotTime
					'set rstTotTime = server.createobject("adodb.recordset")
					'rstTotTime.open "select isnull(SUM(Hours_bill),0) AS tot_hours from times where jobno = " & rs("id"), cnn
					'response.write "select isnull(SUM(Hours_bill),0) AS tot_hours from times where jobno = " & rs("id")
					'if not rstTotTime.eof then
						response.write(rs("hours"))
					'end if
					'rstTotTime.close
					'set rstTotTime = nothing
					%>
      </td>
    </tr>
    <%
			rs.movenext
			if not rs.eof then
				%>
    <script>document.getElementById("matchingJobs").value = document.getElementById("matchingJobs").value + ","</script>
    <%
			end if				
		wend
		%>
    <script>document.getElementById("matchingJobs").value = document.getElementById("matchingJobs").value + ")"</script>
  </table>	 
	


	<table border=0 cellpadding="0" cellspacing="0" width="100%">
		<tr>
			<td width="20%">
				<table border="0" cellspacing="2" cellpadding="3" style="border:1px solid #dddddd;margin:3px;">
					<tr> 
						<td nowrap>Total&nbsp;Jobs&nbsp;Found:&nbsp;<%=unstartedcount + inprogresscount + closedcount + RFPCount%></td>
						<td nowrap align="center" style="position:inline;width:18px;height:12px;background:#ffcc00;border:1px solid #999999;"><%=UnstartedCount%></td>
						<td>&nbsp;Unstarted</td>
						<td nowrap align="center" style="position:inline;width:18px;height:12px;background:#66ff66;border:1px solid #999999;"><%=InProgressCount%></td>
						<td>&nbsp;In&nbsp;Progress</td>
						<td nowrap align="center" style="position:inline;width:18px;height:12px;background:#cc0033;border:1px solid #999999;"><%=ClosedCount%></td>
						<td>&nbsp;Closed</td>
						<td nowrap align="center" style="position:inline;width:18px;height:12px;border:1px solid #999999;"><%=RFPCount%></td>
						<td>&nbsp;RFP</td>
					</tr>
				</table>
			</td>
			<td>*Results limited to 800 jobs.</td>
			 <% if not print then %>
				<td align="right">
					<input type="image" id="printresults" src="/um/opslog/images/btn-print_results.gif" 
						onClick="document.forms.toJobPrint.submit()"
						value="Print" onmouseover="buttonOver(this);" onmouseout="buttonOut(this,'#ffffff');" border="0" style="border:1px solid #ffffff;">
				</td>
			<%end if%>
		</tr>
	</table>
	<!--dont think this form is used any more, but it's there just in case.  -->
	<form name="printwin" method="POST" action="">
		<input type="hidden" name="search" value="<%=search0%>">
		<input type="hidden" name="order" value="<%=order0%>">
		<input type="hidden" name="status" value="<%=status%>">
		<input type="hidden" name="jtype" value="<%=jtype%>">
		<input type="hidden" name="company" value="<%=company%>">
		<input type="hidden" name="custid" value="<%=custid%>">
		<input type="hidden" name="print" value="yes">
	</form>

	</body>
	</html>
	<%
	rs.close
else
	response.redirect "viewjob.asp?jid=" & rs("id")
end if 
%>