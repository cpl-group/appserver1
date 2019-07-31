<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!-- #include file="./adovbs.inc" -->
<%
		if isempty(Session("name")) then
%>
<script>
opener.location="../index.asp"
window.close()
</script>
<%
			'			Response.Redirect "http://www.genergyonline.com"
		end if	

%>
<html>
<head>

<title>Survey Details</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function setValue(tenant_no, surveyid){
 
	var temp="survey_detail.asp?tenant_no="+tenant_no+"&surveyid="+surveyid
	parent.frames.tenant.location = temp
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
	<%
	if request("choice") = "update" then

tenant_no= Request("tenant_no")
survey_id= Request ("survey_id")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.open getConnect(0,0,"Engineering")

Set rst2 = Server.CreateObject("ADODB.Recordset")
Set rst1 = Server.CreateObject("ADODB.Recordset")
	
		sql = "Select * from tblTenantSurvey WHERE (tenant_no = '" & tenant_no & "') and id=" & survey_id 

		rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly

	If not rst1.EOF then 	
	%>
		<form name="submit" method="post" action="surveyupdate.asp">
	  <font face="Arial, Helvetica, sans-serif"> </font> 
	  
  <table width="100%" >
    <tr> 
      <td width="20%" ><font face="Arial, Helvetica, sans-serif"><b>Update Record</b></font></td>
      <td width="20%" >&nbsp;</td>
    </tr>
    <tr> 
      <td width="20%" ><i><font face="Arial, Helvetica, sans-serif">Survey Date</font></i></td>
      <td width="20%" ><i><font face="Arial, Helvetica, sans-serif">Location</font></i></td>
    </tr>
    <tr> 
      <input type="hidden" name="tenant_id" value="<%=rst1("tenant_no")%>">
      <td width="20%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="surveydate" value="<%=rst1("surveydate")%>">
        </font></td>
      <td width="20%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="hidden" name="tenant_no" value="<%=tenant_no%>">
        <input type="hidden" name="survey_id" value="<%=survey_id%>">
        <input type="text" name="location" value="<%=rst1("location")%>">
        </font></td>
    </tr>
    <tr> 
      <td width="20%"><i><font face="Arial, Helvetica, sans-serif">Floor</font></i></td>
      <td width="20%"><i><font face="Arial, Helvetica, sans-serif">Order No</font></i></td>
    </tr>
    <tr> 
      <td width="20%" height="34"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="floor" value="<%=rst1("floor")%>">
        </font></td>
      <td width="20%" height="34"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="orderno" value="<%=rst1("orderno")%>">
        <input type="hidden" name="tenant_no_new" value="0">
        <input type="submit" name="Submit" value="Update">
        </font></td>
    </tr>
  </table>
		</form>
	<% end if
	rst1.close
	set cnn1=nothing
	else 
	
	tenant_no= Request("tenant_no")
	%>

<form name="form1" method="post" action="surveyupdate.asp">
  <table width="100%" >
    <tr> 
      <td width="20%" ><font face="Arial, Helvetica, sans-serif"><b>Add Record</b></font></td>
      <td width="20%" >&nbsp;</td>
    </tr>
    <tr> 
      <td width="20%" ><i><font face="Arial, Helvetica, sans-serif">Survey Date</font></i></td>
      <td width="20%" ><i><font face="Arial, Helvetica, sans-serif">Location</font></i></td>
    </tr>
    <tr> 
      <input type="hidden" name="tenant_id" value="<%=tenant_no%>">
      <td width="20%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="surveydate" value="">
        </font></td>
      <td width="20%"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="location" value="">
        </font></td>
    </tr>
    <tr> 
      <td width="20%"><i><font face="Arial, Helvetica, sans-serif">Floor</font></i></td>
      <td width="20%"><i><font face="Arial, Helvetica, sans-serif">Order No</font></i></td>
    </tr>
    <tr> 
      <td width="20%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="floor" value="">
        </font></td>
      <td width="20%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="orderno" value="">
		<input type="hidden" name="tenant_no" value="<%=tenant_no %>">
        <input type="submit" name="Submit" value="Save">
        </font></td>
    </tr>
  </table>
</form>
<% 
End if

%>
</body>
</html>

