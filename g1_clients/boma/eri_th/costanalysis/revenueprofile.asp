<html>
<head>
<%@Language="VBScript"%>
<%
user=Session("name")
		
Dim bldg
bldg=Request.QueryString("bldgnum")
Dim year
year=Request.QueryString("year")
Dim userid
userid=Request.Querystring("userid")
if bldg<>"" then 		
%>

<title>Revenue Profile</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function viewrevprof(bldg, year, userid) {
	var temp
		temp="revenueprofiledev.asp?bldgnum=" + bldg +"&year=" + year +"&userid="+userid
		document.frames.admin.location=temp
} 
function loadypidlist(bldg,pid) {
	var temp = "revbldglistdev.asp?bldg=" + bldg + "&pid="+pid
	document.location = temp
}
function bldglist(pid){
document.location="revbldglistdev.asp?pid=" + pid
}
function unreported(bldg, year, userid){
	var temp = "unreported.asp?bldg=" + bldg + "&year="+year+"&userid="+userid
 	 window.open(temp,"", "scrollbars=no, width=500, height=600, resizeable, status")

}
</script>
</head>
<%
Dim cnn1
Dim rst1
Dim strsql

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
strsql="Select strt from buildings where bldgnum = '" & Bldg &"'"

rst1.Open strsql, cnn1, adOpenStatic
if not rst1.EOF then 
	bldgname=rst1("strt")
end if
rst1.close
%>

<body bgcolor="#FFFFFF" text="#000000" link="#FFFFFF" vlink="#FFFFFF">
<img src="<%="makechart.asp?bldgNUM=" & bldg & "&year=" & year %>">
<% end if %>
</body>
</html>