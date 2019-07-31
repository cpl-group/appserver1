<!-- #include file="secure.inc" -->
<html>

<head>

<title>Genergy ERI Management</title>

<!-- #include file="./adovbs.inc" -->

<link rel="stylesheet" type="text/css" href="../holiday/holiday.css">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head><style type="text/css">
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
-->
</style>


<%

user = Request("userid")

Set cnn1 = Server.CreateObject("ADODB.Connection")
openStr="data Source=web;user Id=web"
cnn1.Open openStr

Set rst1 = Server.CreateObject("ADODB.Recordset")

sql="SELECT * FROM visitors WHERE (email = N'" & user & "')" 

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly

 %>

<body>

<div align="center">
 <table border="0" width="1042" height="21">
 <tr>
 <td width="155" height="21">

 <p align="right"><font size="5">Welcome MR.</font></p>

 </td>
 <center>
 <td width="873" height="21"><font size="5"><%=rst1("fname") & rst1("lname")%></font></td>
 </tr>
 </table>
 </center>
</div>

<p>&nbsp;</p>

<table border="0" width="100%" height="1" cellspacing="1" cellpadding="0">
<tr>
<td width="100%" height="1" valign="bottom" align="left">

<p align="left"><i>The following is the list of property(s) for your account</i></p>

<p align="left"><i>Click on available services to show additional information. Click 

<a href="http://www.genergy.com/login/client_login/client_login.html" target="_top">here</a>

 to end this session.</i></p>

<p align="left"><i>Update as&nbsp; <!--webbot bot="Timestamp" s-type="REGENERATED" s-format="%B %d, %Y" startspan -->April 16, 2002<!--webbot bot="Timestamp" endspan i-checksum="17445" --></i></td>
</tr>
</table>

<p align="right">&nbsp;</p>

<p align="right">&nbsp;</p>

<table border="1" cellpadding="3" cellspacing="4" width="100%" height="75">
<tr>
<th align="center" width="5%" height="32"><font size="2">Bldg ID</font></th>
<th align="center" width="20%" height="32"><font size="2">Property Name</font></th>
<th align="center" width="10%" height="32"><font size="2">Contact</font></th>
<th align="center" width="5%" height="32"><font size="2">IRI</font></th>
<th align="center" width="5%" height="32"><font size="2">LMP</font></th>
<th align="center" width="5%" height="32"><font size="2">PGI</font></th>
<th align="center" width="5%" height="32"><font size="2">Power Availability</font></th>
<th align="center" width="5%" height="32"><font size="2">Power Chart</font></th>
<th align="center" width="5%" height="32"><font size="2">Revenue Prof</font></th>
<th align="center" width="5%" height="32"><font size="2">SubMeter</font></th>
<th align="center" width="5%" height="32"><font size="2">Over Time HVAC</font></th>
<th align="center" width="5%" height="32"><font size="2">ERI</font></th>
<th align="center" width="5%" height="32"><font size="2">PLP</font></th>
<th align="center" width="5%" height="32"><font size="2">MEP</font></th>
<th align="center" width="5%" height="32"><font size="2">MSI</font></th>
</tr>
  <%
' Create and open ADO Connection object.

Set rst2 = Server.CreateObject("ADODB.Recordset")

sql = "SELECT * from table1 where (userid = N'" & user & "') order by bldgid desc;"

rst2.Open sql, cnn1, adOpenStatic, adLockReadOnly

' Loop thorough recordset object, displaying each record.

Do While Not rst2.EOF
 %>
<tr>
<td valign="top" align="center" width="6%" height="19">

<p align="center"><font size="2"><%Response.write "<a href=" & rst2("bldgid_link") & ">" & rst2("bldgid")%></font></td>
<td valign="top" align="center" width="24%" height="19">

<p align="center"><font size="2"><%=rst2("Bldg_name")%></font></td>
<td valign="top" align="center" width="20%" height="19">

<p align="center"><font size="2"><%=rst2("contact_name")%></font></td>

<%if rst2("iri") then %>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"><%Response.write "<a href=" & rst2("iri_link") & ">" & "Y"%>
<%else %>
</font>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"></font></td>
<%end if%>

<%if rst2("lmp") then %>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"><%Response.write "<a href=" & rst2("lmp_link") & ">" & "Y"%>
<%else %>
</font>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"></font></td>
<%end if%>

<%if rst2("pgi") then %>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"><%Response.write "<a href=" & rst2("pgi_link") & ">" & "Y"%>
<%else %>
</font>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"></font></td>
<%end if%>

<%if rst2("pow_ava") then %>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"><%Response.write "<a href=" & rst2("pow_ava_link") & ">" & "Y"%>
<%else %>
</font>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"></font></td>
<%end if%>

<%if rst2("pa_chart") then %>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"><%Response.write "<a href=" & rst2("pa_chart_link") & ">" & "Y"%>
<%else %>
</font>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"></font></td>
<%end if%>

<%if rst2("rev_prof") then %>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"><%Response.write "<a href=" & rst2("rev_prof_link") & ">" & "Y"%>
<%else %>
</font>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"></font></td>
<%end if%>

<%if rst2("meter") then %>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"><%Response.write "<a href=" & rst2("meter_link") & ">" & "Y"%>
<%else %>
</font>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"></font></td>
<%end if%>

<%if rst2("ovthvac") then %>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"><%Response.write "<a href=" & rst2("ovthvac_link") & ">" & "Y"%>
<%else %>
</font>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"></font></td>
<%end if%>

<%if rst2("eri") then %>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"><%Response.write "<a href=" & rst2("eri_link") & ">" & "Y"%>
<%else %>
</font>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"></font></td>
<%end if%>

<%if rst2("plp") then %>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"><%Response.write "<a href=" & rst2("plp_link") & ">" & "Y"%>
<%else %>
</font>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"></font></td>
<%end if%>

<%if rst2("mep") then %>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"><%Response.write "<a href=" & rst2("mep_link") & ">" & "Y"%>
<%else %>
</font>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"></font></td>
<%end if%>

<%if rst2("msi") then %>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"><%Response.write "<a href=" & rst2("msi_link") & ">" & "Y"%>
<%else %>
</font>
<td valign="top" align="center" width="5%" height="19">

<p align="center"><font size="2"></font></td>
<%end if%>

</tr>
 
  <%
  rst2.MoveNext  
Loop

' Close and destroy the recordset and connection objects.
rst1.Close
rst2.close

Set rst1 = Nothing
set rst2 = Nothing

cnn1.Close
Set cnn1 = Nothing
%>
</table>

<p align="right">&nbsp;</p>

<p align="right">&nbsp;</p>


</body>

</html>
