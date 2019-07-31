<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">

<%@Language="VBScript"%>
<!-- #include file = "Adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
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

function openWindow(jobno,company,jid)
{
// Append jobno to http link
if (company=="IL") {
var urlLink     = "https://appserver1.genergy.com/um/war/jc/jc1.asp?c=" + company + "&ji=" + jobno +"&jid="+jid
}
else {
var urlLink     = "https://appserver1.genergy.com/um/war/jc/jc1.asp?c=" + company + "&jg=" + jobno+"&jid="+jid
}
//var urlLink     = "https://appserver1.genergy.com/um/war/jc/jc.asp?c=" + jobno


// Open new window and customize window settings
window.open(urlLink,"window","scrollbars=no,width=900,height=600,resizeable")
//document.location=urlLink 

}
</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<style type="text/css">
td { font-size:smaller; }
</style>
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
set rs = server.createobject("ADODB.Recordset")


' open connection
cnn.Open getConnect(0,0,"Intranet")
cnn.CursorLocation = adUseClient

rs.open "SELECT crdate FROM sysobjects WHERE name = 'GY_activity_ara_status'",cnn
crdate=rs(0)
rs.close

rs.open "select type as [Job Type], job as [Job Number], hours as [Hours posted since last invoice], max_date as [Last Invoice Date], amount as [Last Invoice Amount], customer_name as [Customer Name] from jtb_adv order by type, jobno desc",cnn
%>

<body bgcolor="#FFFFFF" text="#000000" class="innerbody">

<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#eeeeee">
    <td><b>Job To Bill Report For All Jobs</b></td>
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

Response.Write("<table width=""100%"" border=0 cellpadding=""3"" cellspacing=""1"" bgcolor=""#cccccc"">")
Response.Write("<TBODY>")

'Print out table headers from RecordSet
Response.Write("<TR bgcolor=""#228866"">")
For Each oField In rs.Fields 'Start of For loop
if oField.Name <> "type" then
Response.Write("<td><span class=""standardheader"">" & UCase(oField.Name) & "</span> </TH>")
End If
Next 'End of For loop
Response.Write("</TR>")


'Print values of RecordSet in a table by row
Do While Not rs.EOF
Response.Write("<TR bgcolor=""#ffffff"">")
Response.Write("<TD>" & rs(0) &"</TD>")
Response.Write("<TD><a href=""javascript:openWindow('" & rs(1) & "','" & c & "', '"& Mid(rs(1), 4) &"')""> " & Mid(rs(1), 4) & "</a> </TD>")
Response.Write("<TD>" & rs(2) & "</TD>")
Response.Write("<TD align=left>" & rs(3) & "</TD>")
Response.Write("<TD align=right>" & formatcurrency(rs(4)) & "</TD>")
Response.Write("<TD align=left>" & rs(5) & "</TD>")
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

</body>
</html>