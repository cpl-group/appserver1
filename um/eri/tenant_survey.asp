<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
bldg_no= Request("bldg")
tenant_no= Request("tenant_no")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.open getConnect(0,0,"Engineering")


Set rst1 = Server.CreateObject("ADODB.Recordset")
Set rst2 = Server.CreateObject("ADODB.Recordset")

if isempty(tenant_no) then 
sql = "SELECT * FROM tenant_info WHERE bldg_no='" & bldg_no & "'"
else 
sql = "SELECT * FROM tenant_info WHERE tenant_no='" & tenant_no & "'"

end if 
sql2 = "SELECT tenant_no, tenantname FROM tenant_info WHERE (tenant_no!='" & tenant_no & "') and bldg_no='" & bldg_no &"'"
rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
rst2.Open sql2, cnn1, adOpenStatic, adLockReadOnly
' Write a browser-side script to update another frame (named
' detail) within the same frameset that displays this page.

If not rst1.EOF then 
%>

<html>
<head>
<script language="JavaScript">
function reload(tenant_no, bldg){
	var temp="tenant_survey.asp?tenant_no="+tenant_no+"&bldg="+bldg;
    window.location=temp;
}

function copyTenantSurvey(tenant, bldg){
	var url = "tenant_survey_copier.asp?tenant_no="+tenant+"&bldg="+bldg
	window.open(url,'surveycopier', 'width=400, height=200, scrollbar=no')
}
</script>	
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<script language="JavaScript" type="text/javascript">
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
</script>
</head>

<body bgcolor="eeeeee" text="#000000">
<form name="form1" method="post" action="tenant_survey_update.asp">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr>
  <td width="12%"><b>Select tenant:</b></td>
  <td width="88%"><select name="tenantname" onChange="reload(this.value, bldg_no.value)" >
  <option value="<%=rst1("tenant_no")%>"><%=rst1("tenantname")%></option>
  <%
  Do Until rst2.EOF
  %>
  <option value="<%=rst2("tenant_no")%>"><%=rst2("tenantname")%> , <%=rst2("tenant_no")%></option>
  <%
  rst2.MoveNext
  Loop
  %>
  </select>
  </td>
</tr>
</table>

<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr valign="top">
  <td width="12%">Tenant Number</td>
  <td width="18%">
  <%=rst1("tenant_no")%> 
  <%tmpMoveFrame =  "parent.frames.tenant.location =""survey_detail.asp?tenant_no="&rst1("tenant_no")&"&bldg="&bldg_no&"&surveyid=0"""& vbCrLf 
  
  Response.Write "<script>" & vbCrLf
  
  Response.Write tmpMoveFrame
  Response.Write "</script>" & vbCrLf %>
  
  <input type="hidden" name="tenant_id" value="<%=rst1("tenant_no")%>">
  <input type="hidden" name="bldg_no" value="<%=rst1("bldg_no")%>">
  <input type="hidden" name="tname" value="<%=rst1("tenantname")%>">
  </td>
  <td width="30">&nbsp;</td>
  <td width="12%">Surveyed KWH</td>
  <td width="18%">
  <% if IsNull(rst1("last_sur_kwh")) then %>
  <input type="text" name="last_sur_kwh" value="0">
  <% Else %>
  <input type="text" name="last_sur_kwh" value="<%=rst1("last_sur_kwh")%>">
  <% End if %>
  </td>
  <td width="30">&nbsp;</td>
  <td rowspan="2">
  <table border=0 cellpadding="0" cellspacing="0">
  <tr valign="top">
    <td>Note</td>
    <td width="12">&nbsp;</td>
    <td><textarea name="notes" value="<%=rst1("notes")%>" cols="40"><%=rst1("notes")%> </textarea></td>
  </tr>
  </table>
  </td>
</tr>
<tr valign="top">
  <td>Tenant Name</td>
  <td><%=rst1("tenantname")%></td>
  <td width="30">&nbsp;</td>
  <td>Surveyed KW</td>
  <td>
  <% If IsNull(rst1("last_sur_kw")) then %>
  <input type="text" name="last_sur_kw" value="0">
  <% else %>
  <input type="text" name="last_sur_kw" value="<%=rst1("last_sur_kw")%>">
  <% end if %>
  </td>
  <td width="30">&nbsp;</td>
</tr>
<tr valign="top">
  <td>SqFt</td>
  <td>
  <% if isnull(rst1("sqft")) then %>
  <input type="text" name="sqft" value="0" size="16" maxlength="16">
  <% else %>
  <input type="text" name="sqft" value="<%=rst1("sqft")%>" size="16" maxlength="16">
  <% end if %>
  </td>
  <td width="30">&nbsp;</td>
  <td>Base Hours</td>
  <td>
  <% If isnull(rst1("base_hours")) then %>
  <input type="text" name="base_hours" value="0">
  <% else %>
  <input type="text" name="base_hours" value="<%=rst1("base_hours")%>">
  <% end if %>
  <td width="30"></td>
   <td align="center"><%if not(isBuildingOff(bldg_no)) then%>
   					  <a href="javascript:copyTenantSurvey(document.form1.tenant_id.value, document.form1.bldg_no.value)">Copy Tenant Survey</a><br>
   					  <%if allowgroups("IT Services,Energy Services") then%><a href="javascript:if( confirm('Are you sure you want to delete the entire survey?')){document.location.href='tenant_survey_update.asp?tenant_id=<%=server.urlencode(rst1("tenant_no"))%>&bldg_no=<%=server.urlencode(rst1("bldg_no"))%>&Submit=Delete'}">Delete Tenant Survey</a><%end if%>
					  <%end if%>
   </td>
</tr>
<tr valign="top">
  <td></td>
  <td colspan="6"><%if not(isBuildingOff(bldg_no)) then%><input type="submit" name="Submit" value="Update"><%end if%></td>
</tr>
</table>  

</form>
<% 
End if
rst2.close
rst1.close
set cnn1=nothing
%>
</body>
</html>
