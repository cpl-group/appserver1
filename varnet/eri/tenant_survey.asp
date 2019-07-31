<%@Language="VBScript"%>
<!-- #include file="./adovbs.inc" -->
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'			Response.Redirect "http://www.genergyonline.com"
		end if	

bldg_no= Request("bldg")
tenant_no= Request("tenant_no")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"


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
</script>	
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="FFFFFF" text="#000000">
<form name="form1" method="post" action="tenant_survey_update.asp">
  <table width="100%" height="100%" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="18%" height="3" valign="top"><font face="Arial, Helvetica, sans-serif"><b>Tenant</b></font></td>
      <td width="27%" height="3" valign="top"><font face="Arial, Helvetica, sans-serif"></font></td>
    </tr>
    <tr> 
      <td width="18%" height="2" valign="top"> <font face="Arial, Helvetica, sans-serif" color="#000000"> 
        <%=rst1("tenant_no")%></font> 
        <%tmpMoveFrame =  "parent.frames.tenant.location =" & Chr(34) & _
				  "survey_detail.asp?tenant_no=" & rst1("tenant_no") & "&surveyid=0" & chr(34) & vbCrLf 

		Response.Write "<script>" & vbCrLf
		
		Response.Write tmpMoveFrame
		Response.Write "</script>" & vbCrLf %>
		
        <input type="hidden" name="tenant_id" value="<%=rst1("tenant_no")%>">
        <input type="hidden" name="bldg_no" value="<%=rst1("bldg_no")%>">
        <input type="hidden" name="tname" value="<%=rst1("tenantname")%>">
      </td>
      <td width="27%" height="2" valign="top"> 
        <select name="tenantname" onChange="reload(this.value, bldg_no.value)" >
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
    <tr> 
      <td width="18%" height="2" valign="top"><b><font face="Arial, Helvetica, sans-serif">Surveyed 
        KWH</font></b></td>
      <td width="50%" height="2" valign="top"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="50%"><b><font face="Arial, Helvetica, sans-serif">Surveyed 
              KW </font></b> </td>
            <td width="50%"><b><font face="Arial, Helvetica, sans-serif">SQFT</font></b></td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td width="18%" height="2" valign="top"> <font face="Arial, Helvetica, sans-serif"> 
        <% if IsNull(rst1("last_sur_kwh")) then %>
        <input type="text" name="last_sur_kwh" value="0">
        <% Else %>
        <input type="text" name="last_sur_kwh" value="<%=rst1("last_sur_kwh")%>">
        <% End if %>
        </font> </td>
      <td width="27%" height="2" valign="top"> <font face="Arial, Helvetica, sans-serif"> 
        </font> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="50%"><font face="Arial, Helvetica, sans-serif"> 
              <% If IsNull(rst1("last_sur_kw")) then %>
              <input type="text" name="last_sur_kw" value="0">
              <% else %>
              <input type="text" name="last_sur_kw" value="<%=rst1("last_sur_kw")%>">
              <% end if %>
              </font></td>
            <td width="50%"> 
              <% if isnull(rst1("sqft")) then %>
              <input type="text" name="sqft" value="0" size="16" maxlength="16">
              <% else %>
              <input type="text" name="sqft" value="<%=rst1("sqft")%>" size="16" maxlength="16">
              <% end if %>
            </td>
          </tr>
        </table>
        <font face="Arial, Helvetica, sans-serif"> </font></td>
    </tr>
    <tr> 
      <td width="18%" height="2" valign="top"><b><font face="Arial, Helvetica, sans-serif">Base 
        Hrs</font></b></td>
      <td width="27%" height="2" valign="top"><b><font face="Arial, Helvetica, sans-serif">Note</font></b></td>
    </tr>
    <tr> 
      <td width="18%" height="2" valign="top"><font face="Arial, Helvetica, sans-serif"> 
        <% If isnull(rst1("base_hours")) then %>
        <input type="text" name="base_hours" value="0">
        <% else %>
        <input type="text" name="base_hours" value="<%=rst1("base_hours")%>">
        <% end if %>
        </font></td>
      <td width="27%" height="2"><font face="Arial, Helvetica, sans-serif"> 
        <textarea name="notes" value="<%=rst1("notes")%> cols="40 cols="30""><%=rst1("notes")%></textarea>
        <input type="submit" name="Submit" value="Update">
        </font></td>
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
