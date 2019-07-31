<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim cnn, cmd, rs, sql, fielddef,i, orderby, view, syssql, appmode,rs2,queryfieldname, filtertype, filterval1, filterval2,pdf, showtbox,cmdsql,tracktype, orderbydir, bldgnum, bldgid
view = request("view")
sql = request("sql")
appmode = request("appmode")
orderby = request("orderby") 
orderbydir = request("orderbydir")
bldgid = request("bldgid")

if trim(request("pdf"))="yes" then pdf = true else pdf = false

if sql<>"" and view <> "" then 
	sql = ""
end if
if instr(sql,"delete") or instr(sql,"update") or instr(sql,"drop table") then 
	sql = ""
end if 

queryfieldname 	= request("queryfieldname")
filtertype 		= request("filtertype")
filterval1 		= request("filterval1")
filterval2 		= request("filterval2")

set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")
set rs2 = server.createobject("ADODB.Recordset")
cnn.open application("cnnstr_supermod")
rs.open "SELECT * FROM super_main WHERE id='"&bldgid&"'", cnn
if not rs.eof then bldgnum=rs("bldgnum")
rs.close

syssql = sql

if sql = "" and view <> "" then 
	sql = "select * from  " & view
	showtbox = false
else
	showtbox = true
end if 

if sql="" and view = "" then 
showtbox = false
end if 
%>
<html>
<head>
<title>Utility Manager Reporter</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>
function showfilterval(operator){

if (operator=='between'){
	document.all.filterbox2.style.display='block'
}else{
	document.all.filterbox2.style.display='none'
	document.all.filterval2.value=''
}

}
function clearform(){
	try{
		document.qryform.view.options[0].selected=true
		document.qryform.queryfieldname.options[0].selected=true
		document.qryform.filtertype.options[0].selected=true
		document.qryform.filterval1.value=""
		document.qryform.filterval2.value=""
		document.qryform.sql.value=""
	}catch(exception){//alert(exception.description);
	}
}
</script>
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<body bgcolor="#eeeeee">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<tr>
	<td width="49%" bgcolor="#6699cc" nowrap><span class="standardheader">&nbsp;&nbsp;Utility Manager Reporter<%if pdf then%>: <%=view%>. Printed <%=date()%><%end if%></span></td>
	<%if not pdf then%>
		<td width="51%" align="right" bgcolor="#6699cc" >
			<input type="button" value="Print PDF" onclick="window.open('http://pdfmaker.genergyonline.com/pdfmaker/pdfReport_v2.asp?landscape=true&devIP=<%=request.servervariables("server_name")%>&sn=<%=request.servervariables("script_name")%>&qs=<%=server.urlencode(request.servervariables("query_string"))%>','','')" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
		</td>
	<%end if%>
</tr>
</table>
<table cellpadding="5" cellspacing="0" border="0"><tr><td>
<% if not pdf then %>
	<form name="qryform" action="reportUM.asp" method="get">
	<input type="hidden" name="bldgid" value="<%=bldgid%>">
	<%
	cmdsql = "select * from FreeReporter_Views order by type, description"
	rs.open cmdsql, cnn
	tracktype = ""
	if not rs.eof then%>
		<table width="100%" border="0" cellspacing="3" cellpadding="0" id="simplesearch" <% if showtbox then %>style="display:none;"<%end if%>>
		<tr> 
			<td width="5" nowrap>
				<select name="view" onChange="qryform.submit()">
				<option value="">Select a Report</option><%
				while not rs.eof
					if tracktype = "" then
						%><OPTGROUP Label="<%=rs("type")%> Reports"><%
					elseif trim(tracktype) <> trim(rs("type")) then
						%></OPTGROUP><OPTGROUP Label="<%=rs("type")%> Reports"><% 
					end if
					tracktype = trim(rs("type"))
					%><option value="<%=rs("viewname")%>" <%if trim(rs("viewname"))=trim(view) then%> selected<%end if%>><%=rs("description")%></option><%
					rs.movenext
				wend%>
		        </select>
			</td>
			<%if view <> "" then%>
			<td id="showfilter" align="left">&nbsp;&nbsp;<a href=# onclick="document.all.filters.style.display='block';showfilter.style.display='none'">show filter</a>&nbsp;&nbsp;</td>
			<td id="filters" style="display:none">
				<table>
				<tr>
					<td width="1" align="center">&nbsp;&nbsp;where&nbsp;&nbsp;</td>
					<td nowrap><%response.write bldgnum
						cmdsql = "select * from " & view & " WHERE bldg_id='"&bldgid&"'"
						rs2.open cmdsql, getLocalConnect(bldgnum)%>
						<select name="queryfieldname">
							<option value="">Select a field to filter by</option><%
							for i = 0 to rs2.Fields.Count-1
								if lcase(rs2.fields(i).Name)<>"bldg_id" then
									%><option value="<%=rs2.fields(i).Name%>" <%if trim(rs2.fields(i).Name)=trim(queryfieldname) then%> selected<%end if%>><%=rs2.fields(i).Name%></option><%
								end if
							next
							rs2.close%>
						</select>
					</td>
					<td nowrap>
						<select name="filtertype" onchange="showfilterval(this.value)">
							<option value="">Select a filtertype</option>
							<option value="like"  <%if filtertype="like" then%>selected <%end if%> >is like</option>
							<option value=">" <%if filtertype=">" then%>selected <%end if%> >is greater than</option>
							<option value="<" <%if filtertype="<" then%>selected <%end if%> >is less than</option>
							<option value="<>"<%if filtertype="<>" then%>selected <%end if%> >does not equal</option>
							<option value="between"  <%if filtertype="between" then%>selected <%end if%> >is between</option>
						</select>
					</td>
					<td nowrap><input name="filterval1" type="text" value="<%=filterval1%>" size="10"></td>
					<td <% if filterval2 = "" then %>style="display:none"<%end if%> id="filterbox2" nowrap><input name="filterval2" type="text" value="<%=filterval2%>" size="10"></td>
				</tr>
				</table>
			<%end if %>
			</td>
		</tr>
		</table>
		<br>
		<input type="submit" name="appmode" value="show"><%
	end if 
	rs.close
end if

sql = sql & " WHERE bldg_id='"&bldgid&"'"
'response.write sql
'response.end
if filterval1<>"" then
	select case filtertype
	case "between" 
		filterval1="'" & filterval1 & "' and '" & filterval2 & "'"
	case "like" 
		filterval1="'%" & filterval1 & "%'"
	case ">","<"
	case else
		filterval1="'"&filterval1&"'"
	end select
end if
if queryfieldname<>"" and filtertype<>"" and filterval1<>"" then sql = sql&" and ["&queryfieldname&"] "&filtertype&" "&filterval1
if orderby <> "" and instr(sql,"from") then 
	sql = split(sql, "order by")(0)
	sql = sql & " order by [" & orderby & "] " & orderbydir
end if 
if sql <> "" and appmode = "show" then 
	rs.open sql, getLocalConnect(bldgnum)
	if not rs.eof then 
		if instr(sql, "where") then 
			sql = split(sql, "where")(0)
		end if
		if instr(sql, "order by") then 
			sql = split(sql, "order by")(0)
		end if%>
		<table width="100%" cellpadding="3" cellspacing="0">
		<%if not pdf then%>
			<tr>
			<%for i = 0 to rs.fields.Count - 1%>
				<td nowrap style="border-bottom:1px solid #000000;" align="center">Sort This Column:<br>[<%if instr(sql,"from") then%><a href="reportUM.asp?sql=<%=syssql%>&view=<%=view%>&orderby=<%=rs.fields(i).Name%>&orderbydir=desc&appmode=show&queryfieldname=<%=queryfieldname%>&filtertype=<%=filtertype%>&filterval1=<%=filterval1%>&filterval2=<%=filterval2%>&bldgid=<%=bldgid%>"><%end if%>descending<%if instr(sql,"from") then%></a><%end if%>] [<%if instr(sql,"from") then%><a href="reportUM.asp?sql=<%=syssql%>&view=<%=view%>&orderby=<%=rs.fields(i).Name%>&appmode=show&queryfieldname=<%=queryfieldname%>&filtertype=<%=filtertype%>&filterval1=<%=filterval1%>&filterval2=<%=filterval2%>&bldgid=<%=bldgid%>"><%end if%>ascending<%if instr(sql,"from") then%></a><%end if%>]</td>
			<%next%>
			</tr>
		<%end if%>
		<tr><%for i = 0 to rs.fields.Count - 1%><td nowrap style="border-bottom:1px solid #000000;" align="center"><%if trim(orderby)=trim(rs.fields(i).Name) then%>[<%end if%><%=rs.fields(i).Name%><%if trim(orderby)=trim(rs.fields(i).Name) then%>]<%end if%></td><%next%></tr>
		<%do while not rs.eof%>
			<tr><%for i = 0 to rs.fields.Count - 1%><td <%if rs.Fields.Item(i).Type=2 or rs.Fields.Item(i).Type=2 then %>align="right"<%end if%> nowrap style="border-bottom:1px solid #cccccc"><%=rs(i)%></td><%next%></tr><%
			rs.movenext
		loop%>
		</table>
		<%
	end if
	rs.close
	set rs = nothing
	set cnn = nothing
end if%>
</td></tr></table>
</body>
</html>
