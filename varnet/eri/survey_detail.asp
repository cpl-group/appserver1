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

tenant_no= Request("tenant_no")
survey_id= Request ("surveyid")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"

Set rst2 = Server.CreateObject("ADODB.Recordset")
Set rst1 = Server.CreateObject("ADODB.Recordset")
	
	if survey_id= 0 then 	
		sql = "SELECT * FROM tblTenantSurvey WHERE (tenant_no='" & tenant_no & "')"
	else
		sql = "Select * from tblTenantSurvey WHERE (tenant_no = '" & tenant_no & "' and id=" & survey_id & ")"
	end if
	rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
%>
<html>
<head>

<title>Survey Details</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function  processchange(tenant_no, id, processtype){
	var temp="survey_details_update.asp?tenant_no="+tenant_no+"&survey_id="+id+"&choice="+processtype
	//alert(temp)
    window.open(temp,"", "scrollbars=yes, width=500,height=180" );
}
function setValue(tenant_no, surveyid){
 
	var temp="survey_detail.asp?tenant_no="+tenant_no+"&surveyid="+surveyid
	parent.frames.tenant.location = temp
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%
If not rst1.EOF then 

		' Write a browser-side script to update another frame (named
		' detail) within the same frameset that displays this page.

		tmpMoveFrame =  "parent.frames.details.location = " & Chr(34) & _
				  "surveyitems.asp?survey_id=" & rst1("id") & "&orderno=" & rst1("orderno") &"&id=" & rst1("id") & "&tenant_no=" & rst1("tenant_no") & "&location=" & rst1("location") & chr(34) & vbCrLf 

		Response.Write "<script>" & vbCrLf
		Response.Write tmpMoveFrame
		Response.Write "</script>" & vbCrLf 

%>
	<form name="form1" method="post" action="">
	  <table width="100%" >
		<tr>  
		
		  <td width="20%" ><b><font face="Arial, Helvetica, sans-serif">Survey Date</font></b></td>
		  <td width="20%" ><b><font face="Arial, Helvetica, sans-serif">Location</font></b></td>
		  <td width="20%" ><b><font face="Arial, Helvetica, sans-serif">Floor</font></b></td>
		  <td width="20%" ><b><font face="Arial, Helvetica, sans-serif">Order No</font></b></td>
		<td>&nbsp;</td>
		</tr>
		<tr> 
			<input type="hidden" name="tenant_id" value="<%=rst1("tenant_no")%>">		
		  
      <td width="20%"> <font face="Arial, Helvetica, sans-serif"><%=rst1("surveydate")%> </font></td>
		  <td width="20%"> <font face="Arial, Helvetica, sans-serif"> 
			<input type="hidden" name="tenant_no" value="<%=tenant_no%>">
			<input type="hidden" name="survey_id" value="<%=survey_id%>">		
			<select name="location" onChange="setValue( tenant_no.value,this.value)">
			<% 
			
			if survey_id > 0 then 
			
				sql = "Select id, location,ORDERNO from tblTenantSurvey WHERE (tenant_no = '" & tenant_no & "' and id=" & survey_id & ") order by orderno"
				rst2.Open sql, cnn1, adOpenStatic, adLockReadOnly
				
				if not rst2.eof then	    
			%>
			<option value="<%=rst2("id")%>" selected>   <%=rst2("ORDERNO") & "   " %><%=left(rst2("location"),30)%> </option>
			<option value="<%=survey_id%>"> ====================== </option>					
			<%
				rst2.close
				end if
				
			else %>
			<option value="<%=rst1("id")%>"> <%=rst1("orderno") & "   " %><%=left(rst1("location"),30)%></option>
			<option value="<%=survey_id%>"> ====================== </option>					
			<%
			end if
					
			sql = " Select id, location,orderno  from tblTenantSurvey WHERE (tenant_no = '" & tenant_no & "') order by orderno"
			rst2.Open sql, cnn1, adOpenStatic, adLockReadOnly
			
			Do until rst2.EOF 
			%>
			<option value="<%=rst2("id")%>">  <%=rst2("orderno") & "  " %><%=left(rst2("location"),30)%></option>
			<%
			rst2.movenext
			loop
			rst2.close
			%>
			</select>
			</font></td>
		  <td width="20%"> <font face="Arial, Helvetica, sans-serif"> 
			<%=rst1("floor")%>
			</font></td>
		  <td width="20%"> <font face="Arial, Helvetica, sans-serif"> 
			<%=rst1("orderno")%>
			</font></td>
		  <td width="20%"><font face="Arial, Helvetica, sans-serif"> 
			<input type="hidden" name="add" value="add">			
			<input type="hidden" name="update" value="update">						
        	<input type="button" name="Button" value="Update" onclick="processchange(tenant_no.value, survey_id.value, update.value)">
			<input type="button" name="Button" value="Add" onClick="processchange(tenant_no.value, survey_id.value, add.value)">
			</font></td>
			
		</tr>
	  </table>
	</form>

<% Else 

' Write a browser-side script to update another frame (named
		' detail) within the same frameset that displays this page.

		tmpMoveFrame =  "parent.frames.details.location = " & Chr(34) & _
				  "blank.htm"& chr(34) & vbCrLf 

		Response.Write "<script>" & vbCrLf
		Response.Write tmpMoveFrame
		Response.Write "</script>" & vbCrLf 
%> 
<form name="form1" method="post" action="surveyupdate.asp">
  <table width="100%" >
    <tr> 
      <td width="20%"> <font face="Arial, Helvetica, sans-serif"> 
	  <input type="hidden" name="add" value="add">			
	  <input type="hidden" name="update" value="update">
      <input type="hidden" name="tenant_no" value="<%=tenant_no%>">
	  <input type="hidden" name="survey_id" value="<%=survey_id%>">
	    <input type="button" name="Button" value="Add Location" onClick="processchange(tenant_no.value, survey_id.value, add.value)">
        </font></td>
    </tr>
  </table>
</form>
<% 
End if
rst1.close
set cnn1=nothing
%>
</body>
</html>
