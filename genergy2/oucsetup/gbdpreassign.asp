<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<html>
<head>
<title>gEnergyOne</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}

.standard { font-family:Arial,Helvetica,sans-serif;font-size:8pt; }
.bottomline { border-bottom:1px solid #eeeeee; }
.floorlink { font-family:Arial,Helvetica,sans-serif;font-size:8pt; color:#0099ff; }
a.floorlink:hover { color:lightgreen; }
.shrunkenheader { font-family:Arial,Helvetica,sans-serif;font-size:7pt;font-weight:bold; }

-->
</style>
<script language="javascript">
function formsubmit()
{
 document.form1.formsubmitted.value = 1
 //alert(document.form1.formsubmitted.value)
}
function loadmeters()
{
	url = "gbdpreassign.asp?premise_point=" + document.form1.premise_point.value
	document.location = url
}
</script>

<%
if cint(request("formsubmitted"))=1 then 
	dim cnn, rst, cmd, prm,premisepoint, premise, point, meterid, datadef, meterPname
	premisepoint = split(request("premise_point"),"_")
	premise = premisepoint(0)
	meterPname = premisepoint(1)
	point 	= premisepoint(2)
	meterid = request("meterid")
	datadef = request("datadef")
'	response.write premise &":"& point&":"& meterid&":"& datadef
'	response.end	
	Set rst = Server.CreateObject("ADODB.recordset")
	set cnn = server.createobject("ADODB.Connection")
	set cmd = server.createobject("ADODB.Command")
	set rst = server.createobject("ADODB.Recordset")
	cnn.Open application("Cnnstr_genergy2")
	cnn.CursorLocation = adUseClient
	cmd.CommandText = "sp_assigngb"
	cmd.CommandType = adCmdStoredProc
	
	Set prm = cmd.CreateParameter("premise", adVarChar, adParamInput, 50)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("meterPname", adVarChar, adParamInput, 50)
	cmd.Parameters.Append prm 
	Set prm = cmd.CreateParameter("point", adVarChar, adParamInput, 50)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("meterid", adVarChar, adParamInput, 50)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("datadef", adTinyInt, adParamInput)
	cmd.Parameters.Append prm
	cmd.Name = "reassign"
	
	Set cmd.ActiveConnection = cnn

	
'	response.write "'"&premise&"','"&meterPname&"','"&point&"',"&meterid&",'"&datadef&"'"
'	response.end
	cnn.reassign premise, meterPname, point, meterid, datadef
	Set cnn = nothing
	response.redirect "gbdpreassign.asp"
else 

	Dim premise_point
	premise_point = Request("Premise_point")
	%>
<link rel="Stylesheet" href="../styles.css" type="text/css">
</head>
	<body BGCOLOR="#eeeeee" LINK="#0000CC" VLINK="#0000CC" TEXT="#000000">
	<form method="POST" name="form1" action="gbdpreassign.asp">
	  <table border=0 cellpadding="3" cellspacing="0" width="400">
		<tr bgcolor="#3399cc"> 
		  <td><span class="standardheader">Data Point Assignment</span></td>
		  <td>&nbsp;</td>
		</tr>
		<tr valign="top"> 
		  <td>
				<select name="premise_point" size="10" id="select" onclick="loadmeters()">
				<%
					dim rst1, cnn1, db_premise_point
					set cnn1 = server.createobject("ADODB.connection")
					set rst1 = server.createobject("ADODB.recordset")
					cnn1.open application("cnnstr_genergy2")
					rst1.open "SELECT Distinct meterPname, Premise, point FROM garbage ORDER BY premise, meterPName, point", cnn1
					while not rst1.EOF 
					db_premise_point = rst1("premise") & "_" & rst1("meterPname") & "_" & rst1("point")
				%>
				<option value="<%=db_premise_point%>" <%if trim(ucase(cstr(premise_point))) = trim(ucase(cstr(db_premise_point))) then %> selected <%end if %>><%=rst1("premise")%>.<%=rst1("meterPname")%>.<%=rst1("point")%></option>
				<%  rst1.movenext
					wend
					rst1.close
				%>
				</select> 

    </td>
		  <td>
      <table border=0 cellpadding="3" cellspacing="0">
      <tr>
        <td>1. </td>
        <td>Select premise point at left</td>
      </tr>
      <tr>
        <td>2. </td>
        <td>
			  <select name="meterid">
			  <option value="">Select Meter</option>
				<% 
				if premise_point <> "" then 
					Dim pre_pnt 
					pre_pnt = split(premise_point, "_")
					rst1.open "SELECT meterid, meternum, c.* FROM meters m INNER JOIN tblleasesutilityprices lup ON lup.leaseutilityid=m.leaseutilityid LEFT JOIN custom_oucAccount c ON c.billingid=lup.billingid WHERE c.premiseid='" & pre_pnt(0) & "' ORDER BY meternum", cnn1
					
					while not rst1.EOF
					%>
				<Option value="<%=rst1("meterid")%>"><%=rst1("meternum")%></Option>
				<%
					rst1.movenext
					wend 
					rst1.close
				end if	
				%>
			  </select><br>
      </td>
		</tr>
		<tr>
		  <td>3. </td>
      <td>
			  <select name="datadef">
			  <option value="">Select Point Definition</option>
				<%
				rst1.open "SELECT * FROM pointdefs", cnn1
					while not rst1.EOF
					%>
				<Option value="<%=rst1("id")%>"><%=rst1("pointdesc")%> (<%=rst1("dpointname")%>) 
				</Option>
				<%
					rst1.movenext
					wend 
					rst1.close
				%>
			  </select><br>
  		</td>
		</tr>
		<tr>
		  <td>4. </td>
  		<td>
      <input name="formsubmitted" type="hidden" value="0">
      <input type="submit" name="Submit" value="ASSIGN" onclick="document.all['loadFrame'].style.visibility='visible';formsubmit()">
      </td>
    </tr>
    </table>		  
    </td>
		</tr>
	  </table>
	</form>
<div id="loadFrame" style="visibility:hidden; position:absolute;left:200;top:150;background-color:lightyellow;border-width:1px;border-style:solid">
<table border="0" cellpadding="5" cellspacing="0"><tr><td style="font-family:arial;font-size:16px;font-weight:bold">Loading Data...</td></tr></table>
</div>
	</body>
<% end if %>
</html>





