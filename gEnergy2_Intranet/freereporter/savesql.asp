<%@ language = VBScript %>
<%option explicit%>
<!-- #include virtual="/genergy2/secure.inc" -->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%
	Dim sql, cnn, rs, sqlstr,userid, appmode,grpsave
	
	
	
	  %>
	     <script>
			<!--testing-->
			function groupSave(groupval){
			//alert(groupval);
			 }
			</script>
	<%
    grpsave= replace(request("group"),"|","")
	if grpsave="0" then grpsave="Genergy Users"
	sql = request("sql")
	'response.write grpsave
	'response.end
	set cnn = server.createobject("ADODB.Connection")
	set rs = server.createobject("ADODB.Recordset")
	cnn.open getConnect(0,0,"intranet")
	userid = getKeyValue("user")
	appmode = request("mode")
	if appmode = "" then appmode = "Show"
	select case appmode
	case "save"
			Dim rptname, viewname, categorytext
			sql =request("sql")
			rptname = replace(request("rptname"),"'","")
			viewname = replace(request("viewname"),"'","")
			categorytext = replace(request("categorytext"),"'","")
			
			if instr(sql,"delete") or instr(sql,"update") or instr(sql,"drop table") then 
			 sql = ""
			end if 
			'response.write grpsave
			'response.end
			sqlstr = "create view ["&viewname&"] as " & sql
			'response.end
			cnn.execute sqlstr
			sqlstr = "insert into FreeReporter_Views (description, viewname, type,groupname) values ('"&rptname&"','"&viewname&"','"&categorytext&"','"&grpsave&",Report_Admin')" ' add  |grpsave| to sql insert
			'response.write sqlstr
			'response.end
			cnn.execute sqlstr
			%>
			<script>
			opener.document.location = "/genergy2_intranet/freereporter/report.asp?view=<%=viewname%>&sql=&appmode=show&Submit=Show+Report"
			window.close()
			</script>
			<%
	case "Show"
%>
	<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
	<body bgcolor="#eeeeee" leftmargin="3">
	<table width="100%" border="0" cellspacing="0" cellpadding="0">
	  <tr> 
		<td width="49%" bgcolor="#6699cc" nowrap><span class="standardheader">&nbsp;Save 
		  custom SQL query as Report</span></td>
	  </tr>
	</table>
	<br>
	<form  name="saveSQL" action="savesql.asp" method="get">
	<table width="100%" border="0" cellspacing="3" cellpadding="0" id="freesql">
	<tr>
		  <td>This Report will be saved using your sql query as seen below: <br>
			<font size="1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;modifications can be made in the free reporter</font></td>
	</tr>
	<tr>
	<td>
		<textarea name="sql" cols="75" rows="5" readonly style="background-color:#CCCCCC;"><%=sql%></textarea>
		<input name="appmode" type="hidden" value="show">
	</td>
	</tr>
	</table>
	<br>
	  <table width="100%" border="0" cellspacing="0" cellpadding="0">
		<tr> 
		  <td>&nbsp;Report Name:
	      <input name="rptName" type="text" onChange="viewname.value = 'FR_'+this.value.replace(/ /gi, '_')+'_Created_by_<%=userid%>_<%=replace(date(),"/","")%>_<%=replace(replace(time(),":","")," ","")%>';" maxlength="50"></td>
		  <td>&nbsp;</td>
		</tr>
		<tr>
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		</tr>
		<tr bgcolor="#FFFFCC"> 
		  <td colspan=2 style="border:1px black solid">&nbsp;&nbsp;&nbsp;&nbsp;For 
			your reference, a new view will be created as:   		    <br> 
		    		    &nbsp;&nbsp;&nbsp;&nbsp; 
			<input name="viewname" type="text" style="background-color:#CCCCCC;" size="65" readonly><br>
			&nbsp;&nbsp;&nbsp;&nbsp;<font size="1"><strong>Any References to this 
			report to IT Services should be made using the view<br>
		  &nbsp;&nbsp;&nbsp;&nbsp;name seen above </strong></font> </td>
		</tr>
		<tr> 
		  <td>&nbsp;</td>
		  <td>&nbsp;</td>
		</tr>
		<tr><td colspan='2'>&nbsp;Limit Report To:
		<!--New Ando-->
		<select name="group" Onchange='groupSave(this.value);'>
            <option value="0">Select Group for Report </option>
            <%	
					Dim sqlstr2
					sqlstr2 = "select distinct groupname,grouplabel from AD_Groups order by grouplabel"
					rs.open sqlstr2, getconnect(0,0,"dbcore")
					if not rs.eof then
						while not rs.eof
						%>
            <option value="<%="|"&rs("groupname")&"|"%>"><%=rs("grouplabel")%></option>
            <%
						rs.movenext
						wend
					end if 
					rs.close
				%>
          </select>
		<!--New Ando-->
		</td></tr>
		  <td>&nbsp;Category:
		    <select name="category" onChange="categorytext.value=this.value;categorytext.style.backgroundColor='pink'">
              <option value="Enter New Category Name Here">New Category</option>
              <%	
					sqlstr = "select distinct type from dbo.FreeReporter_Views order by type desc"
					rs.open sqlstr, cnn
					if not rs.eof then
						while not rs.eof
						%>
              <option value="<%=rs("type")%>"><%=rs("type")%></option>
              <%
						rs.movenext
						wend
					end if 
					rs.close
				%>
            </select></td><td align='left'>or
              <input name="categorytext" type="text" style="background-color:pink;" onClick="this.style.backgroundColor='white';" onChange="category.selectedIndex=0" value="Enter New Category Name Here" size="30" maxlength="30"></td>
		    <td>            
		    </td><td>&nbsp;
		  </td>
		</tr>
		<tr> 
		  <td>&nbsp;</td>
		<!--  <td>&nbsp;</td>-->
		</tr>
	  </table>
	  <div align="center">
		<input type="hidden" name="mode" value="save">
       <!-- <input name="categorytext" type="text" style="background-color:pink;" onClick="this.style.backgroundColor='white';" onChange="category.selectedIndex=0" value="Enter New Category Name Here" size="30" maxlength="30">-->
        <input name="Save" type="submit" id="Save4" value="Save Report">
	  </div>
	</form>
<% end select%>
	