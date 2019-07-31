<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<%@Language="VBScript"%>
<!-- #include file = "Adovbs.inc" -->
<html>
<head>
<title>Job To Bill Report</title>

<style type="text/css">
<!-- 
h2 {letter-spacing: 0.25cm}
p  {color: #0000ff        }
-->
</style>

<script type="text/javascript">

function openWindow(jobno,company)
{

// Append jobno to http link
var urlLink     = "https://appserver1.genergy.com/um/war/jc/jc.asp?c=" + company + "&j=" + jobno
//var urlLink     = "https://appserver1.genergy.com/um/war/jc/jc.asp?c=" + jobno


// Open new window and customize window settings
window.open(urlLink,"window","scrollbars=no,width=900,height=600,resizeable")

}
</script>

</head>


<%  	
'Declare variables  
'http params
dim crdate,dp,c,o,total,currentjob,currentcustomer,gtotal

dp = UCase(request("dp"))  'Make query string value for dept./jobno uppercase
c = UCase(request("c"))    'Make query string value for company uppercase
o = request("o")
m = request("m")


'adodb vars
dim cnn, cmd, rs, prm
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")


' open connection
cnn.Open "driver={SQL Server};server=10.0.7.20;uid=sa;pwd=!general!;database=main;"
cnn.CursorLocation = adUseClient

rs.open "SELECT crdate FROM sysobjects WHERE name = 'GY_activity_ara_status'",cnn
crdate=rs(0)
rs.close

'Response.Write crdate
'response.end

' specify stored procedure 
cmd.CommandText = "sp_jtb"
cmd.CommandType = adCmdStoredProc

'Create Parameter
Set prm1 = cmd.CreateParameter("dep", advarchar, adParamInput, 2)
cmd.Parameters.Append prm1
Set prm2 = cmd.CreateParameter("month", advarchar, adParamInput, 2)
cmd.Parameters.Append prm2

' assign internal name to stored procedure
cmd.Name = "test"
Set cmd.ActiveConnection = cnn

'return set to recordset rs
cnn.test  dp, m, rs
%>

<body bgcolor="#FFFFFF" text="#000000">
<br>
<br>
<H2 align=center>Job To Bill Report For Type <% response.write "<p>" & dp & "</p>" %> 
                                       Since <% response.write "" & m & "/1/2002" %> </H2>
<br>
<br>
<hr>
<br>
<br>

<%

'Response.write c
If rs.EOF = "True" Then  'If recordset is empty then no matching records.
   Response.Write("<br>")
   Response.Write("<br>")
   Response.Write("<br>")
   Response.Write("<br>")
   Response.Write("<P colspan=10 align=center><B>There are no matching records.</B></P>")
   Response.Write("<br>")
   Response.Write("<br>")
   Response.Write("<br>")
   Response.Write("<br>")
Else

Response.Write("<TABLE Cols=3 align=center cellSpacing=3 cellPadding=2 width=100% border=0>")
Response.Write("<TBODY>")

'Print out table headers from RecordSet
Response.Write("<TR>")
For Each oField In rs.Fields 'Start of For loop
if oField.Name <> "type" then
Response.Write("<TH align=left colspan=2>" & oField.Name & " </TH>")
End If
Next 'End of For loop
Response.Write("</TR>")

'Declare empty row for space
Response.Write("<TR>")
Response.Write("<TD>&nbsp;</TD>")
Response.Write("</TR>")

'Print values of RecordSet in a table by row
Do While Not rs.EOF
Response.Write("<TR>")
Response.Write("<TD colspan=2 align=left>")
Response.Write("<a href=""javascript:openWindow('" & rs(0) & "','" & c & "')""> " & rs(0) & "</a>")
'Response.Write("<a href='javascript:openWindow()'> " & rs(0) & "</a>")
Response.Write("</TD>")
Response.Write("<TD colspan=2 align=left>" & rs(2) & "</TD>")

Response.Write("<TD colspan=2 align=left>" & rs(3) & "</TD>")
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
	
%>

<br>
<br>
<hr>
<br>
<br>

</body>
</html>
