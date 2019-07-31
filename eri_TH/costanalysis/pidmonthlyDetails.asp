<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE file="revFunctions.asp"-->
<%
dim date1, date2, b, utype, numberofRows, pid
b = request.querystring("b")
pid = request.querystring("pid")
date1 = request.querystring("date1")
date2 = request.querystring("date2")
utype = request.querystring("utype")
numberofRows = checkprefs()

dim ArrDataSeriesERI(12)
dim ArrDataSeriesSubMetered(12)
dim ArrDataSeriesExpenses(12)
dim ArrDataSeriesUnreportedExp(12)
dim ArrDataSeriesUnreportedRev(12)
dim ArrDataSeriesMac(12)
dim ArrDataSeriesPlp(12)
dim ArrDataSeriesNet(12)
Dim ArrPieRevenue(3)
Dim ArrPieExpenses(3)

dim ArrDataSeriesERI_one(12)
dim ArrDataSeriesSubMetered_one(12)
dim ArrDataSeriesExpenses_one(12)
dim ArrDataSeriesUnreportedExp_one(12)
dim ArrDataSeriesUnreportedRev_one(12)
dim ArrDataSeriesMac_one(12)
dim ArrDataSeriesPlp_one(12)
dim ArrDataSeriesNet_one(12)
Dim ArrPieRevenue_one(3)
Dim ArrPieExpenses_one(3)


dim i
initarrays()
call getdataSetsPortfolio(date1, pid, utype, "0")
for i = 1 to 12
	ArrDataSeriesERI_one(i) = ArrDataSeriesERI(i)
	ArrDataSeriesSubMetered_one(i) = ArrDataSeriesSubMetered(i)
	ArrDataSeriesExpenses_one(i) = ArrDataSeriesExpenses(i)
	ArrDataSeriesUnreportedExp_one(i) = ArrDataSeriesUnreportedExp(i)
	ArrDataSeriesUnreportedRev_one(i) = ArrDataSeriesUnreportedRev(i)
	ArrDataSeriesMac_one(i) = ArrDataSeriesMac(i)
	ArrDataSeriesPLP_one(i) = ArrDataSeriesPLP(i)
	ArrDataSeriesNet_one(i) = ArrDataSeriesNet(i)
next

if date2<>"" then
	dim ArrDataSeriesERI_two(12)
	dim ArrDataSeriesSubMetered_two(12)
	dim ArrDataSeriesExpenses_two(12)
	dim ArrDataSeriesUnreportedExp_two(12)
	dim ArrDataSeriesUnreportedRev_two(12)
	dim ArrDataSeriesMac_two(12)
	dim ArrDataSeriesPlp_two(12)
	dim ArrDataSeriesNet_two(12)
	Dim ArrPieRevenue_two(3)
	Dim ArrPieExpenses_two(2)
	initarrays()
	call getdataSets(date2, b, utype, "0")
	for i = 1 to 12
		ArrDataSeriesERI_two(i) = ArrDataSeriesERI(i)
		ArrDataSeriesSubMetered_two(i) = ArrDataSeriesSubMetered(i)
		ArrDataSeriesExpenses_two(i) = ArrDataSeriesExpenses(i)
		ArrDataSeriesUnreportedExp_two(i) = ArrDataSeriesUnreportedExp(i)
		ArrDataSeriesUnreportedRev_two(i) = ArrDataSeriesUnreportedRev(i)
		ArrDataSeriesMac_two(i) = ArrDataSeriesMac(i)
		ArrDataSeriesPLP_two(i) = ArrDataSeriesPLP(i)
		ArrDataSeriesNet_two(i) = ArrDataSeriesNet(i)

		if trim(ArrDataSeriesERI_two(i)) then ArrDataSeriesERI_two(i) = 0
		if trim(ArrDataSeriesSubMetered_two(i)) then ArrDataSeriesSubMetered_two(i) = 0
		if trim(ArrDataSeriesExpenses_two(i)) then ArrDataSeriesExpenses_two(i) = 0
		if trim(ArrDataSeriesUnreportedExp_two(i)) then ArrDataSeriesUnreportedExp_two(i) = 0
		if trim(ArrDataSeriesUnreportedRev_two(i)) then ArrDataSeriesUnreportedRev_two(i) = 0
		if trim(ArrDataSeriesMac_two(i)) then ArrDataSeriesMac_two(i) = 0
		if trim(ArrDataSeriesPLP_two(i)) then ArrDataSeriesPLP_two(i) = 0
		if trim(ArrDataSeriesNet_two(i)) then ArrDataSeriesNet_two(i) = 0
	next
end if
%>
<html>
<head>
<title></title>
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

<body bgcolor="#FFFFFF" text="#000000" onload="parent.closeLoadBox('loadFrame2')" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<table width="706" border="0" cellspacing="0" cellpadding="0">
	<tr><td bgcolor="#000000" width="50%"><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2"><b>Monthly Details</b></font></td>
		<td bgcolor="#000000" width="50%" align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b><a href="javascript:parent.loadoptions()" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Return To Options</a></b></font><font color="#FFFFFF" face="Arial, Helvetica, sans-serif" size="2">&nbsp;</font></td>
	</tr>
</table>
<table border="0" cellspacing="0" cellpadding="0" align="center">
<tr><td valign="top" width="106">
	<table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
	<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"><td>&nbsp;</td></tr>
	<%if ArrPrefs(exps) then%><tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"><td>Expenses</td></tr><%end if%>
	<%if ArrPrefs(subm) then%><tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"><td>Submeter</td></tr><%end if%>
	<%if ArrPrefs(eri) then%><tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"><td>ERI</td></tr><%end if%>
	<%if ArrPrefs(urar) then%><tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"><td>Revenue&nbsp;Adjustment</td></tr><%end if%>
	<%if ArrPrefs(urae) then%><tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"><td>Expense&nbsp;Adjustment</td></tr><%end if%>
	<%if ArrPrefs(mac) then%><tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"><td>Mac&nbsp;Adjustment</td></tr><%end if%>
	<%if ArrPrefs(plp) then%><tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"><td>PLP</td></tr><%end if%>
	<%if ArrPrefs(net) then%><tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"><td>Net</td></tr><%end if%>
	</table>
</td><td valign="top">
<div style="width:600; overflow:auto; height: <%=numberofRows*15+33%>;">
<table border="1" cellspacing="0" cellpadding="0" bordercolor="#CCCCCC">
<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;">
<%for i = 1 to 12%>
	<td align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=left(monthname(i),3)%>&nbsp;<%=right(date1,2)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
	<%if date2<>"" then%>
		<td align="center">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=left(monthname(i),3)%>&nbsp;<%=right(date2,2)%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
	<%end if%>
<%next%>
</tr>

<%if ArrPrefs(exps) then%>
	<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"> 
	<%for i = 1 to 12
		response.write "<td align=""right"">"&formatcurrency(ArrDataSeriesExpenses_one(i),1)&"K</td>"
		if date2<>"" then response.write "<td align=""right"">"&formatcurrency(ArrDataSeriesExpenses_two(i),1)&"K</td>"
	next%>
	</tr>
<%end if%>

<%if ArrPrefs(subm) then%>
	<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"> 
	<%for i = 1 to 12
		response.write "<td align=""right"">"&formatcurrency(ArrDataSeriesSubMetered_one(i),1)&"K</td>"
		if date2<>"" then response.write "<td align=""right"">"&formatcurrency(ArrDataSeriesSubMetered_two(i),1)&"K</td>"
	next%>
	</tr>
<%end if%>

<%if ArrPrefs(eri) then%>
	<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"> 
	<%for i = 1 to 12
		response.write "<td align=""right"">"&formatcurrency(ArrDataSeriesERI_one(i),1)&"K</td>"
		if date2<>"" then response.write "<td align=""right"">"&formatcurrency(ArrDataSeriesERI_two(i),1)&"K</td>"
	next%>
	</tr>
<%end if%>

<%if ArrPrefs(urar) then%>
	<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"> 
	<%for i = 1 to 12
		response.write "<td align=""right"">"&formatcurrency(ArrDataSeriesUnreportedRev_one(i),1)&"K</td>"
		if date2<>"" then response.write "<td align=""right"">"&formatcurrency(ArrDataSeriesUnreportedRev_two(i),1)&"K</td>"
	next%>
	</tr>
<%end if%>

<%if ArrPrefs(urae) then%>
	<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"> 
	<%for i = 1 to 12
		response.write "<td align=""right"">"&formatcurrency(ArrDataSeriesUnreportedExp_one(i),1)&"K</td>"
		if date2<>"" then response.write "<td align=""right"">"&formatcurrency(ArrDataSeriesUnreportedExp_two(i),1)&"K</td>"
	next%>
	</tr>
<%end if%>

<%if ArrPrefs(mac) then%>
	<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"> 
	<%for i = 1 to 12
		response.write "<td align=""right"">"&formatcurrency(ArrDataSeriesMac_one(i),1)&"K</td>"
		if date2<>"" then response.write "<td align=""right"">"&formatcurrency(ArrDataSeriesMac_two(i),1)&"K</td>"
	next%>
	</tr>
<%end if%>

<%if ArrPrefs(plp) then%>
	<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"> 
	<%for i = 1 to 12
		response.write "<td align=""right"">"&formatcurrency(ArrDataSeriesPLP_one(i),1)&"K</td>"
		if date2<>"" then response.write "<td align=""right"">"&formatcurrency(ArrDataSeriesPLP_two(i),1)&"K</td>"
	next%>
	</tr>
<%end if%>

<%if ArrPrefs(net) then%>
	<tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;"> 
	<%for i = 1 to 12
		dim color
		color = "black"
		if ArrDataSeriesNet_one(i)<0 then color = "red"
		response.write "<td style=""color:"& color &""" align=""right"">"&formatcurrency(abs(ArrDataSeriesNet_one(i)),1)&"K</td>"
		if date2<>"" then 
			color = "black"
			if ArrDataSeriesNet_one(i)<0 then color = "red"
			response.write "<td style=""color:"& color &""" align=""right"">&nbsp;"&formatcurrency(abs(ArrDataSeriesNet_two(i)),1)&"K</td>"
		end if
	next%>
	</tr>
<%end if%>
</table></div>
</td></tr></table>
<%if ArrPrefs(net) then%>
<table border="0" cellspacing="0" cellpadding="0"><tr style="font-family: Arial, Helvetica, sans-serif; font-size: 10;" valign="top">
	<td><img src="pidnetgainchart.asp?a=<%=ArrDataSeriesNet_one(1)%>,<%=ArrDataSeriesNet_one(2)%>,<%=ArrDataSeriesNet_one(3)%>,<%=ArrDataSeriesNet_one(4)%>,<%=ArrDataSeriesNet_one(5)%>,<%=ArrDataSeriesNet_one(6)%>,<%=ArrDataSeriesNet_one(7)%>,<%=ArrDataSeriesNet_one(8)%>,<%=ArrDataSeriesNet_one(9)%>,<%=ArrDataSeriesNet_one(10)%>,<%=ArrDataSeriesNet_one(11)%>,<%=ArrDataSeriesNet_one(12)%>" width="300" height="110"></td>
	<td>&nbsp;&nbsp;</td>
</tr></table>
<%end if%>

</body>
</html>
