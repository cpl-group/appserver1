<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<%@Language="VBScript"%>
<!-- #include file = "Adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
	<head>
		<title>Metering Jobs Status</title>
		<style type="text/css">
			<!-- 
h2 {letter-spacing: 0.25cm}
p  {color: #0000ff        }
--></style>
		
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
			<style type="text/css">
td { font-size:smaller; }
</style>
	</head>
	<%  	
'Declare variables  
'http params
dim crdate,dp,c,o,total,currentjob,currentcustomer,gtotal

'dp = UCase(request("dp"))  'Make query string value for dept./jobno uppercase
'c = UCase(request("c"))    'Make query string value for company uppercase
'o = request("o")
'm = request("m")


'adodb vars
dim cnn, cmd, rs, prm
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")

cnn.CommandTimeout =0

' open connection
cnn.Open getConnect(0,0,"Intranet")
cnn.CursorLocation = adUseClient

cmd.CommandText = "metering_jobs_status"
cmd.CommandType = adCmdStoredProc
cmd.CommandTimeout = 0

' assign internal name to stored procedure
cmd.Name = "test"
Set cmd.ActiveConnection = cnn


'return set to recordset rs
cnn.test  rs
%>
	<body bgcolor="#FFFFFF" text="#000000" class="innerbody">
		<table border="0" cellpadding="3" cellspacing="0" width="100%">
			<tr bgcolor="#eeeeee">
				<td><b>Metering Jobs Status</b></td>
			</tr>
		</table>
		<%

'Response.write c
If rs.EOF = "True" Then  'If recordset is empty then no matching records.
   Response.Write("<br>")
   Response.Write("<br>")
   Response.Write("<div align=""center"" class=""notetext"">There are no matching records.</div>")
   Response.Write("<br>")
   Response.Write("<br>")
Else

Response.Write("<table  border=0 cellpadding=""3"" cellspacing=""1"" bgcolor=""#cccccc"">")
Response.Write("<TBODY>")

'Print out table headers from RecordSet
Response.Write("<TR bgcolor=""#228866"">")
For Each oField In rs.Fields 'Start of For loop
	Response.Write("<td nowrap><span class=""standardheader"">" & UCase(oField.Name) & "</span> </TH>")
Next 'End of For loop
Response.Write("</TR>")


'Print values of RecordSet in a table by row
Do While Not rs.EOF
Response.Write("<TR bgcolor=""#ffffff"">")

for i = 0 to rs.fields.Count - 1
if (rs.Fields.Item(i).Type = 6 or rs.Fields.Item(i).Type = 5) then 'if value is of type money
	Response.Write("<TD align=right>" & rs(i)& "</TD>")
else
	if rs.Fields.Item(i).ActualSize <=100  then  'wrap fields that are greater that 100 characters in length
		Response.Write("<TD align=left nowrap >" &  rs(i) & "</TD>")
	else 
		Response.Write("<TD align=left >" & rs(i) & "</TD>")
	end if
end if
next

Response.Write("</TR>")

rs.MoveNext 'Move to next record in RecordSet
Loop 'End of Do while not rs.EOF loop	

Response.Write("</TBODY>")   
Response.Write("</TABLE>") 'End of Table
	
rs.Close
Set rs = Nothing
cnn.Close
Set cnn = Nothing

End If	
	
Response.Write("* = Job has multiple tickets opened.")
%>
		<br>
		<br>
		
	</body>
</html>
