<%@ language = VBScript %>
<%option explicit%>
<!-- #include virtual="/genergy2/secure.inc" -->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%
dim cnn, cmd, rs, sql, fielddef,i, orderby, view, syssql, appmode,rs2,queryfieldname, filtertype, filterval1, filterval2,pdf, showtbox,cmdsql,tracktype, orderbydir, bgcolor


view = request("view")
sql = request("sql")
appmode = request("appmode")
orderby = request("orderby") 
orderbydir = request("orderbydir")

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
cnn.open getConnect(0,0,"intranet")

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



if sql <> "" and appmode = "show" then 

	if view = "[GY.PayableView]" and queryfieldname = "day" then 
	filtertype ="> PayableView"
	end if
	
	if queryfieldname <> "" and filterval1 <> "" and filtertype <> "" then 
		select case filtertype
		case ">","<","<>"
			sql = sql & " where [" & queryfieldname & "] " & filtertype & " '" & filterval1 & "'"
		case "like"
			sql = sql & " where [" & queryfieldname & "] like '%" & filterval1 & "%'"
		case "between"
			if filterval2 <> "" then 
				sql = sql & " where [" & queryfieldname & "] between '" & filterval1 & "' and '" & filterval2 & "'"
			end if
		case "> PayableView"
		'sql = sql & " where (DATEDIFF(dd,Invoice_Date, GETDATE()) >='"& filterval1 &"')"
		
		end select

	end if
	if orderby <> "" and instr(sql,"from") then 
		sql = split(sql, "order by")(0)
		sql = sql & " order by [" & orderby & "] " & orderbydir
      end if 
       if queryfieldname = "day" and  filtertype ="> PayableView" and view = "[GY.PayableView]"  then 
	  	sql = sql & " where (DATEDIFF(dd,Invoice_Date, GETDATE()) >='"& filterval1 &"') order by customer,Job"
	'   response.write sql' &"<BR>"
	  ' response.end
	   end if
end if 
'Open query
if sql <> "" then 
		'response.write  sql
		'response.end
		sql = "select top 1000 * from ("&sql&") s" 
		rs.open sql, cnn
end if 

%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
 <script>
 /*
 //try {
  var obj= new Active.Controls.Grid;
  /*obj.setId("datagrid");
  obj.setRowCount(57);
  obj.setColumnCount(7);
  obj.setDataText(function(i, j){return obj_data[i][j]});
  obj.setColumnText(function(i){return obj_columns[i]});
  document.write(obj);
 document.write(obj.getSortProperty("direction")); 
/*}
catch (error){
  document.write(error.description);
}
finally
{

}*/
 
 
 
 
 
 </script>
<% if not pdf then %>
	<style> body, html {margin:0px; padding: 0px; overflow: hidden;font: menu;border: none;} </style>
	<!-- ActiveWidgets stylesheet and scripts -->
	<link href="/includes/aw/styles/xp/grid.css" rel="stylesheet" type="text/css" ></link>
	<script src="/includes/aw/lib/grid.js"></script>

	<!-- ActiveWidgets ASP functions -->
	<!-- #INCLUDE virtual="/includes/aw/activewidgets.asp" -->
<style>
			#datagrid{height: 400px; border: 2px inset; background: white}
			.active-grid-row {border-bottom: 1px solid threedlightshadow;}
			.active-grid-column {border-right: 1px solid threedlightshadow;}
			
</style>
<% end if%>
</head>

<script>
function AndoOrder(){
try {
  var obj= new Active.Controls.Grid;
  /*obj.setId("datagrid");
  obj.setRowCount(57);
  obj.setColumnCount(7);
  obj.setDataText(function(i, j){return obj_data[i][j]});
  obj.setColumnText(function(i){return obj_columns[i]});
  document.write(obj);*/
}
catch (error){
  document.write(error.description);
}
finally
{
document.write(obj.getSortProperty("direction"));
}


}		
function showfilterval(operator){

if (operator=='between'){
	document.all.filterbox2.style.display='block'
}else{
	document.all.filterbox2.style.display='none'
	document.all.filterval2.value=''
}

}
function clearform(view){
	try{
	document.qryform.queryfieldname.options[0].selected=true
	document.qryform.filtertype.options[0].selected=true
	document.qryform.filterval1.value=""
	document.qryform.filterval2.value=""
	document.qryform.sql.value=""
	document.all.filters.style.display='none'
	document.all.showfilter.style.display='none'
	}catch(exception){//alert(exception.description);
	}
	try{
	if(view != "true"){
		document.qryform.view.options[0].selected=true
	}
	}catch(exception){}

}
function savesql(sql){
	 sql = encodeURI(sql)
	 window.open('/genergy2_intranet/freereporter/savesql.asp?sql='+sql,'SaveSQl','height=350, width=400')
}
</script>
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<% if pdf then bgcolor = "#FFFFFF" else bgcolor = "#eeeeee" end if %>
<body bgcolor="<%=bgcolor%>">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="49%" bgcolor="#6699cc" nowrap><span class="standardheader">Free Reporter<%if pdf then%>:query=<%=sql%>. Printed <%=date()%><%end if%></span></td>
     <% 'response.write sql
	' response.end%>
	  <%if not pdf then%>
    <td width="51%" align="right" bgcolor="#6699cc" >
    <%' response.write sql &"<BR>"
	'response.write "orderby="&orderby
	'response.end
	'orderby="amount"'''''''''''''''
	%>
	
	<!--<input type="button" value="Print PDF" onClick="window.open('http://pdfmaker.genergyonline.com/pdfmaker/AndoReport_v2.asp?landscape=true&devIP=<'%=request.servervariables("server_name")%>&sn=<'%=request.servervariables("script_name")%>&qs=<%=server.urlencode("sql="&sql&"&view="&view&"&orderby="&orderby&"&appmode=show&queryfieldname="&queryfieldname&"&filtertype="&filtertype&"&filterval1="&filterval1&"&filterval2="&filterval2)%>','','')" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;">-->
    <input type="button" value="Print PDF" onclick="window.open('http://pdfmaker.genergyonline.com/pdfmaker/pdfReport_v2.asp?landscape=true&devIP=<%=request.servervariables("server_name")%>&sn=<%=request.servervariables("script_name")%>&qs=<%=server.urlencode("sql="&sql&"&view="&view&"&orderby="&orderby&"&appmode=show&queryfieldname="&queryfieldname&"&filtertype="&filtertype&"&filterval1="&filterval1&"&filterval2="&filterval2)%>','','')" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
    </td>
       <%end if%>
 </tr>
</table>
<% if not pdf then %>
<form name="qryform" action="report.asp" method="get">
<%
Dim grpStr,splitter,grpstrWord,count,sign,groupsql
grpStr =getkeyvalue ("groupstring")
splitter = Split(left(grpStr,len(grpStr)-1),"|",-1,1) 
'splitter = Split(grpStr,"|",-1,1)
count =0
For each grpstrWord in splitter

if count > 0 then groupsql= groupsql +" or groupname like '%" & ltrim(rtrim(grpstrWord)) & "%'" else groupsql= " groupname like '%" & ltrim(rtrim(grpstrWord)) & "%'" 
count = count +1
next
groupsql = Replace(groupsql, "or groupname like '%%'", "")
'response.write groupsql' getkeyvalue ("groupstring")
'response.end
cmdsql = "select * from FreeReporter_Views where " & groupsql &" order by type, description"'"select * from FreeReporter_Views order by type, description"
'cmdsql = "select * from FreeReporter_Views order by type, description"
'response.write cmdsql
rs2.open cmdsql, cnn
tracktype = ""
if not rs2.eof then 
%>
  <table width="100%" border="0" cellspacing="3" cellpadding="0" id="simplesearch" <% if showtbox then %>style="display:none;"<%end if%>>
    <tr>
      <td nowrap colspan=6><a href="#" onClick="simplesearch.style.display='none';freesql.style.display='block';clearform('false');">Free 
        Form SQL</a></td>
    </tr>
    <tr> 
      <td width="5" nowrap><select name="view" onChange="clearform('true');">
          <option value="">Select a Report</option>
          <% while not rs2.eof %>
          		  <% 
		  if tracktype = "" then
		  %>
		  <OPTGROUP Label="<%=rs2("type")%> Reports">
		  <%
		  elseif trim(tracktype) <> trim(rs2("type")) then 
		  %>
		  </OPTGROUP><OPTGROUP Label="<%=rs2("type")%> Reports">
		  <% 
		  end if 
		  tracktype = trim(rs2("type"))
		  %>
<option value="<%=rs2("viewname")%>" <%if trim(rs2("viewname"))=trim(view) then%> selected<%end if%>><%=rs2("description")%></option>
          <%
		rs2.movenext
		wend
		rs2.close
%>
        </select></td>
      <% if view <> "" then %>
      <td id="showfilter" align="left" <% if trim(filtertype) <> "" then %>style="display:none"<%end if%>>&nbsp;&nbsp;<a href=# onClick="document.all.filters.style.display='block';showfilter.style.display='none';">show filter</a>&nbsp;&nbsp;</td>
      <td id="filters" <% if trim(filtertype) = "" then %>style="display:none"<%end if%>>
	  <table>
	  <tr>
	  <td width="1" align="center">&nbsp;&nbsp;where&nbsp;&nbsp;</td>
		<td nowrap> 
		<%
		'cmdsql = "select top 1 * from " & view 
		%>        <select name="queryfieldname">
          <option value="">Select a field to filter by</option>
          <%for i = 0 to rs.Fields.Count-1%>
          <option value="<%=rs.fields(i).Name%>" <%if trim(rs.fields(i).Name)=trim(queryfieldname) then%> selected<%end if%>><%=rs.fields(i).Name%></option>
          <%
next
%>
        </select> </td>
      <td  nowrap> <select name="filtertype" onChange="showfilterval(this.value)">
          <option value="">Select a filtertype</option>
          <option value="like"  <%if filtertype="like" then%>selected <%end if%> >is 
          like</option>
          <option value=">" <%if filtertype=">" then%>selected <%end if%> >is greater than</option>
          <option value="<" <%if filtertype="<" then%>selected <%end if%> >is less than</option>
          <option value="<>"<%if filtertype="<>" then%>selected <%end if%> >does not equal 
          </option>
          <option value="between"  <%if filtertype="between" then%>selected <%end if%> >is 
          between</option>
        </select> </td>
      <td nowrap>
<input name="filterval1" type="text" value="<%=filterval1%>" size="10"> 
      </td>
      <td  <% if filterval2 = "" then %>style="display:none"<%end if%> id="filterbox2" nowrap><input name="filterval2" type="text" value="<%=filterval2%>" size="10"></td>
      <%end if %>
	    <%' if report = "dsdsd" then%>
		<!--<td>order by</td>
		<td>
		  <select name="groupbytype">
		 <option value="">Select group</option>
		  <option value="A">Account Manager</option> 
		 <option value="C">Customer</option>
         <option value="J">Job</option>
		 </select>
		</td>-->
	   <%'end if%>
	   </tr>
	  </table>
	
	  </td>
    </tr>
  </table>
<br>
  <% 
else
	rs2.close
end if 
%>
<table width="100%" border="0" cellspacing="3" cellpadding="0" id="freesql" <% if not showtbox then %>style="display:none;"<%end if%>>
<tr>
<td><a href="#" onClick="simplesearch.style.display='block';freesql.style.display='none';savesqlasview.style.display='none';clearform();">Simple Search</a><br>
Type Your SQL Statement here:
</td>
</tr>
<tr>
<td><textarea name="sql" cols="100" rows="2"><%=sql%></textarea>
 		<input name="appmode" type="hidden" value="show">
        <% if sql <> "" then %><input name="savesqlasview" type="button" id="savesqlasview" value="Save as Report" onClick="savesql(sql.value)"><% end if %></td>
</tr>
</table>
  &nbsp;&nbsp;<input type="reset" name="Submit2" value="Clear">&nbsp;<input type="submit" name="Submit" value="Show Report">
</form><br>
<% if view <> "" then %>
<i>&nbsp;&nbsp; returns only top 1000 records found. Refine your query using the filters or request an expanded output from I.T.</i>
<%end if%>
<%end if %>
<%

if sql <> "" and appmode = "show" then 
'	rs.open sql, cnn
	
	if not rs.eof then 
				if instr(sql, "where") then 
					sql = split(sql, "where")(0)
				end if
				if instr(sql, "order by") then 
					sql = split(sql, "order by")(0)
				end if
				
				%>
				<%if not pdf then
							
						Response.write(activewidgets_grid("obj", rs, "datagrid"))
				else %>
				<table width="100%" cellpadding="3" cellspacing="0">
				  <tr>
				<%			
				for i = 0 to rs.fields.Count - 1
				'response.write sql
				'response.end
				%><td nowrap style="border-bottom:1px solid #000000;" align="center"><%if trim(orderby)=trim(rs.fields(i).Name) then%>[<%end if%><%=rs.fields(i).Name%><%if trim(orderby)=trim(rs.fields(i).Name) then%>]<%end if%></td>
				<%
				next
				%></tr>
				<%	
				do while not rs.eof
				%><tr><%
				for i = 0 to rs.fields.Count - 1
				%>	
				<td <%if isnumeric(rs(i)) then %>align="right"<%end if%> nowrap style="border-bottom:1px solid #cccccc"><%=rs(i)%></td>
				<%
				next	
				%></tr><%
			rs.movenext
			loop
				%></table>
<%				
	end if
	end if
	rs.close
	set rs = nothing
	set cnn = nothing
	
end if
%>
</body>
</html>
