<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
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
survey_id= Request("surveyid")
bldg_no =  Request("bldg")

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.open getConnect(0,0,"Engineering")

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
 
	var temp="survey_detail.asp?bldg=<%=bldg_no%>&tenant_no="+tenant_no+"&surveyid="+surveyid
	parent.frames.tenant.location = temp
}
</script>
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">   
</head>

<body bgcolor="#eeeeee" text="#000000">
<%
If not rst1.EOF then 

		' Write a browser-side script to update another frame (named
		' detail) within the same frameset that displays this page.

		tmpMoveFrame =  "parent.frames.details.location = ""surveyitems.asp?survey_id=" & rst1("id") & "&bldg="&bldg_no&"&orderno=" & rst1("orderno") &"&id=" & rst1("id") & "&tenant_no=" & rst1("tenant_no") & "&location=" & rst1("location") & "&xscroll=" & request("xscroll") & "&yscroll=" & request("yscroll") & """"

		Response.Write "<script>" & vbCrLf
		Response.Write tmpMoveFrame
		Response.Write "</script>" & vbCrLf 

%>
	  <table border=0 cellpadding="3" cellspacing="1" bgcolor="#dddddd" width="100%">
		<tr bgcolor="#dddddd">  
		  <td width="20%">Survey Date</td>
		  <td width="20%">Location</td>
		  <td width="20%">Floor</td>
		  <td width="20%">Order No</td>
		  <td width="20%">&nbsp;</td>
		</tr>
<form name="form1" method="post" action="surveyupdate.asp">
<%if not(isBuildingOff(bldg_no)) then%>
    <tr bgcolor="#eeeeee"> 
      <input type="hidden" name="tenant_id" value="<%=tenant_no%>">
        <td width="20%"><input type="text" name="surveydate" value=""></td>
        <td width="20%"><input type="text" name="location" value=""></td>
        <td width="20%"><input type="text" name="floor" value=""></td>
        <td width="20%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="orderno" value=""></td>
        <td>
        <input type="hidden" name="tenant_no" value="<%=tenant_no %>">
        <input type="submit" name="Submit" value="Add" style="padding-left:6px;padding-right:6px;">
        </td>
    </tr>
<%end if%>
</form>
	<form name="form1" method="post" action="">
		<tr bgcolor="#eeeeee"> 
			<input type="hidden" name="tenant_id" value="<%=rst1("tenant_no")%>">		
		  
      <td> <%=rst1("surveydate")%> </td>
		  <td>  
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
			'survey_id = rst1("id")
			end if
					
			sql = " Select id, location,orderno  from tblTenantSurvey WHERE (tenant_no = '" & tenant_no & "') order by orderno"
			rst2.Open sql, cnn1, adOpenStatic, adLockReadOnly
			
			Do until rst2.EOF 
			if survey_id=0 then survey_id=rst2("id")
			%>
			<option value="<%=rst2("id")%>">  <%=rst2("orderno") & "  " %><%=left(rst2("location"),30)%></option>
			<%
			rst2.movenext
			loop
			rst2.close
			%>
			</select>
			</td>
		  <td>  
			<%=rst1("floor")%>
			</td>
		  <td>  
			<%=rst1("orderno")%>
			</td>
		  <td> 
			<input type="hidden" name="add" value="add">			
			<input type="hidden" name="update" value="update">						
			<%if not(isBuildingOff(bldg_no)) then%>
        	<input type="button" name="Button" value="Update" onclick="processchange(tenant_no.value, survey_id.value, update.value)">
        	<%if allowgroups("IT Services,Energy Services") then%><input type="button" name="Delete" value="Delete" onclick="document.location.href='surveyupdate.asp?tenant_no=<%=tenant_no%>&survey_id=<%=survey_id%>&submit=Delete'"><%end if%>
			<!--[[input type="button" name="Button" value="Add" onClick="processchange(tenant_no.value, survey_id.value, add.value)"]]-->
			<%end if%>
			</td>
			
		</tr>
	</form>
	  </table>

<% Else 

' Write a browser-side script to update another frame (named
		' detail) within the same frameset that displays this page.

		tmpMoveFrame =  "parent.frames.details.location = ""blank.htm"""

		Response.Write "<script>" & vbCrLf
		Response.Write tmpMoveFrame
		Response.Write "</script>" & vbCrLf 
%> 
<form name="form1" method="post" action="surveyupdate.asp">
  <table width="100%" >
    <tr> 
      <td width="20%">  
	  <input type="hidden" name="add" value="add">			
	  <input type="hidden" name="update" value="update">
      <input type="hidden" name="tenant_no" value="<%=tenant_no%>">
	  <input type="hidden" name="survey_id" value="<%=survey_id%>">
	    <%if not(isBuildingOff(bldg_no)) then%><input type="button" name="Button" value="Add Location" onClick="processchange(tenant_no.value, survey_id.value, add.value)"><%end if%>
        </td>
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
