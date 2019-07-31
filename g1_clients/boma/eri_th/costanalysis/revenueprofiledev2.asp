<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="adovbs.inc"-->
<%
Dim bldg
bldg=Request.QueryString("bldgnum")
Dim year
year=Request.QueryString("year")
Dim Bldgname
Dim userid
userid=Request.Querystring("userid")
if bldg = "" then 
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>

<%response.end
end if


'---------------------Values-----------------
Dim Title(12) 'Pair title
Title(1) = "Jan"
Title(2) = "Feb"
Title(3) = "Mar"
Title(4) = "Apr"
Title(5) = "May"
Title(6) = "Jun"
Title(7) = "Jul"
Title(8) = "Aug"
Title(9) = "Sep"
Title(10) = "Oct"
Title(11) = "Nov"
Title(12) = "Dec"

Dim StackTitle(5) 
StackTitle(1) = "UnReported"
StackTitle(2) = "Expenses"
StackTitle(3) = "Submeter"
StackTitle(4) = "ERI"
StackTitle(5) = "Net"

Dim i(12, 5) 'i(Bar#, Stack)
Dim Values(12, 5)
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


strsql = "SELECT eri_data.dbo.eri_total.eri_amt AS ERI_rev, UtilityBill.TotalBillAmt AS Expenses, SUM(tblBillByPeriod.TotalAmt) AS SubMetered, BillYrPeriod.BillPeriod, tblRPentries_sum.Expr1 AS UnReportedAmts FROM UtilityBill FULL OUTER JOIN BillYrPeriod FULL OUTER JOIN tblRPentries_sum ON  BillYrPeriod.BillPeriod = tblRPentries_sum.period AND BillYrPeriod.BldgNum = tblRPentries_sum.bldgnum AND BillYrPeriod.BillYear = tblRPentries_sum.year ON UtilityBill.ypId = BillYrPeriod.ypId FULL OUTER JOIN eri_data.dbo.eri_total ON BillYrPeriod.BldgNum = eri_data.dbo.eri_total.bldg_no AND BillYrPeriod.BillPeriod = eri_data.dbo.eri_total.BillPeriod AND BillYrPeriod.BillYear = eri_data.dbo.eri_total.BillYear FULL OUTER JOIN tblBillByPeriod ON  UtilityBill.ypId = tblBillByPeriod.ypId WHERE (BillYrPeriod.BldgNum = '"& bldg &"') AND (BillYrPeriod.BillYear = '"& year &"') GROUP BY eri_data.dbo.eri_total.eri_amt, UtilityBill.TotalBillAmt,  BillYrPeriod.BillPeriod, tblRPentries_sum.Expr1 ORDER BY BillYrPeriod.BillPeriod"

rst1.Open strsql, cnn1, adOpenStatic

if rst1.EOF then

%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#0099FF"> 
      <div align="left"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">No 
        Revenue Profile has been compiled for <%=year%> for <%=bldgname%></font></b></font></div>
    </td>
  </tr>
</table>
<%

else 

Dim numRecords
numRecords = rst1.RecordCount
if numRecords > 12 then numRecords=12 end if
Dim x
dim tot
dim tots

tot = 0
for x = 1 to numRecords
	if rst1("UnreportedAmts") <> "" then i(x, 1) = Formatnumber(rst1("UnreportedAmts")) else i(x,1)=0 end if
	if rst1("Expenses") <> "" then i(x, 2) = Formatnumber(rst1("Expenses")) else i(x,2)=0 end if
	if rst1("Submetered") <> "" then i(x, 3) = Formatnumber(rst1("Submetered")) else i(x,3)=0 end if
	if rst1("eri_rev") <> "" then i(x, 4) = Formatnumber(rst1("eri_rev")) else i(x,4)=0 end if
	
	i(x, 5) = ccur(i(x,2)) + ccur(i(x,3))+ ccur(i(x,4)) - ccur(i(x,1))

	if ccur(i(x,1)) < 0 then 
		tots =  i(x,2) + (i(x,1)*-1)
	else
		tots =  i(x,2) - (ccur(i(x,1)))
	end if
rst1.MoveNext
tot=tot+tots
next

	rst1.Close
	Set rst1 = Nothing

	cnn1.Close
	Set cnn1 = Nothing

for x= numRecords+1 to 12
	i(x, 5) = 0
	i(x, 4) = 0
	i(x, 2) = 0
	i(x, 3) = 0
	i(x, 1) = 0
next

numRecords = 12
Dim Maxvalue 
Maxvalue = tot /2
Dim TopPadding
TopPadding = 40
Dim SidePadding 
SidePadding = 120
Dim BarPadding
BarPadding = 50
Dim Barwidth 
Barwidth = 23
Dim LineLengths
LineLengths = 70
'---------Advanced-----------------------------

Dim NumBars 'Must be the same as or lower than i(x, 3)
NumBars = numRecords	     'Must be the same as or lower than Title(x)
	     'Must assign values

Dim NumStacks  'Must be the same as or lower than i(6, x) 
NumStacks = 5     'Must have color(from 1 to numbars).gif
	     'Must assign values

%>
<HTML>
<HEAD>
<STYLE TYPE="text/css">
<!--
.Bargraph {
	position:absolute;
	top:0px;
	left:0px;
	z-index:2;
	visibility:hidden;
	font: 10pt MS Comic Sans,arial,sans-serif;
	}
-->
</STYLE>

</HEAD>
<BODY>
<%
x=0
Dim r
Dim s
Dim t 'Counter Variables
Dim max
Dim maxmoney
Dim step 'Stack height
Dim startTop
max=0
maxmoney=0
for x = 1 to numbars
for r = 1 to numstacks
values(x, r) = i(x, r)

if values(x,r) > maxmoney then 
	maxmoney = values(x,r) 
end if

i(x, r) = ((i(x, r) / Maxvalue) * 1000)
	if i(x,r) > max then
		max=i(x,r)
	end if
next
next
maxmoney = (maxmoney/1000) 
Dim height
Dim Amt
Dim index
dim currheight
amt = 0
index = 0

%>

<%
t = 0
Dim PosLeft
PosLeft = SidePadding + 30
for x = 1 to numbars
startTop=0 
for r = 1 to numstacks
t = t + 1

	Step = 0
	If r =2  then
		For s = 1 to (r-1)
			Step = FormatNumber((Step + i(x, 3))/7,0)
		Next
	End If


if i(x, r) <> 0 then
%>
<DIV CLASS=Bargraph ID=bg<%=t%>>
<%if r = 2 or (r=1 and i(x,r) < 0) then
	dim abs
	
	if r=1 then 
		abs = i(x,r) * -1
	else
		abs = i(x,r)
	end if	
	%>
	
<img src="color<%=r%>.gif" Alt=<%=FormatCurrency(Values(x, r))%> Height=<%=(abs +  3)%> Width=<%=Barwidth  + 10%> border="1" style="border-style: solid; border-color: #0000FF" class="bar" id=section<%=r%>>
<%else%>
<%if values(x,r) > 0 then  
if r <> 5 then 
%>
<img src="color<%=r%>.gif" Alt=<%=FormatCurrency(Values(x, r))%> Height=<%=(i(x, r) + 4)%> Width=<%=16%> border="1" style="border-style: solid; border-color: #0000FF" class="bar" id=section<%=r%>>

<%
end if
end if
end if%>
</DIV>


<SCRIPT LANGUAGE="javascript">
		var BarGraphOBJ = eval("document.all['bg<%=t%>']");
		BarGraphOBJ.style.posLeft = (<%if r=2 then Response.Write PosLeft-5 else response.write PosLeft+3 end if%>);
		<% 
			currheight = FormatNumber(((400 - i(x, r) + TopPadding) - Step),0)
			if r=2 or (r=1 and i(x,r) < 0) then 
						If (Amt < FormatNumber(Values(x, r),0)) then
									height = currheight - TopPadding/2
									Amt = FormatNumber(Values(x, r),0)
							else
								if Amt = 0 then 
										height = currheight - TopPadding/2
										Amt = FormatNumber(Values(x, r),0)
								end if	
						end if
			end if
					
		%>
 		BarGraphOBJ.style.posTop = (<%=FormatNumber(((400 - i(x, r) + TopPadding) - Step),0)%>)
		BarGraphOBJ.style.visibility = "visible";
</SCRIPT>







<%
end if

If r = 1 then 
	t = t + 1
%>

	<DIV CLASS=Bargraph ID=Title<%=t%> >
	<%=Title(x)%>
	</DIV>
<SCRIPT LANGUAGE="javascript">
		var BarGraphOBJ = eval("document.all['Title<%=t%>']");
		BarGraphOBJ.style.posLeft = (<%=PosLeft%>);
		BarGraphOBJ.style.posTop = (400 + 10 + <%=TopPadding%>);
		BarGraphOBJ.style.visibility = "visible";
</SCRIPT>
<% 
End If
	Next
PosLeft = PosLeft + BarPadding
Next
Dim l
Dim j
j=1
%>

<DIV CLASS=Bargraph ID=line>
<p><img src="line.gif" width=<%=numbars * barpadding + 90%> height=3> </p>
</DIV>
<SCRIPT LANGUAGE="javascript">
		var BarGraphOBJ = eval("document.all['line']");
		BarGraphOBJ.style.posLeft = (<%=Sidepadding - 65%>);
		BarGraphOBJ.style.posTop = (400 + <%=TopPadding%> - 7 );
		BarGraphOBJ.style.visibility = "visible";
</SCRIPT>
//<DIV CLASS=Bargraph ID=lineinit>
//<p>$<%=Formatnumber((Maxmoney),0)%>k <img src="line.gif" width=<%=numbars * barpadding + 90%> height=1> </p>
//</DIV>
//<SCRIPT LANGUAGE="javascript">
//		var BarGraphOBJ = eval("document.all['lineinit']");
//		BarGraphOBJ.style.posLeft = (<%=Sidepadding - 100%>);
//		BarGraphOBJ.style.posTop = <%=height%>;
//		BarGraphOBJ.style.visibility = "visible";
//</SCRIPT>
<%
'for l=2 to 7%> 
//<DIV CLASS=Bargraph ID=line<%=l%>>
//<p>$<%=Formatnumber((Maxmoney-(Maxmoney*(l*.07))),0)%>k <img src="line.gif" width=<%=numbars * barpadding + 90%> height=1> </p>
//</DIV>
//<SCRIPT LANGUAGE="javascript">
//		var BarGraphOBJ = eval("document.all['line<%=l%>']");
//		BarGraphOBJ.style.posLeft = (<%=Sidepadding - 100%>);
//		BarGraphOBJ.style.posTop =  <%=height%> + <%=j*(height/7)%>;
//		BarGraphOBJ.style.visibility = "visible";
//</SCRIPT>
<%
'j= j + 1
'next%>

<%for x = 1 to Numstacks%>
<DIV CLASS=Bargraph ID=legendline<%=x%>>
<img src="color<%=x%>.gif" width=<%=numbars * barpadding + 90%> height=2> 
</DIV>
<SCRIPT LANGUAGE="javascript">
		var BarGraphOBJ = eval("document.all['legendline<%=x%>']");
		BarGraphOBJ.style.posLeft = (<%=Sidepadding - 65%>);
		BarGraphOBJ.style.posTop = (400 + <%=TopPadding%> - 7 + <%=x%> * 23);
		BarGraphOBJ.style.visibility = "visible";
</SCRIPT>

<DIV CLASS=Bargraph ID=legend<%=x%>>
  <table  width=<%=104 + numbars * (barpadding)%>>
    <tr>
<% if x = 5 then %>
      <td width = 12%> &nbsp;<b><Font size=1><%=stacktitle(x)%></font></b> 
        <%else%>
      <td width = 12%>
	<img src="color<%=x%>.gif" width=10 height=10>&nbsp;<Font size=1><%=stacktitle(x)%></font>
	<%end if%>
	</td>

<%for r = 1 to NumBars%>
      <td  width = 7.6% valign=top> 
        <%if x <> 4 then %>
        <Font size=1><%="$" & FormatNumber((values(r, x)/1000),0)& "k"%></font> 
        <%else
		 if FormatNumber((values(r, x)/1000),0) < 0 then 
				dim num
				num = FormatNumber((values(r, x)/1000),0) * (-1)
			%>
        <font size="1" color="#FF0000"><%="($" & num & "k)"%></font> 
        <% else %>
        <Font size=1><%="$" & FormatNumber((values(r, x)/1000),0)& "k"%></font> 
        <%
		end if
		end if%>
      </td>
<%next%> 

</tr>
</table>
</DIV>
<SCRIPT LANGUAGE="javascript">
		var BarGraphOBJ = eval("document.all['legend<%=x%>']");
		BarGraphOBJ.style.posLeft = (<%=Sidepadding - 70%>);
		BarGraphOBJ.style.posTop = (400 + <%=TopPadding%> + 25 * <%=x%>);
		BarGraphOBJ.style.visibility = "visible";
</SCRIPT>
<%next%>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#0099FF"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">Monthly 
      Electrical Revenue &amp; Expense Profile For <%=bldgname%> For </font><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF"><%=year%></font></b></font></b></font></td>
  </tr>
</table>
<%end if

%>
<p align="left">&nbsp;</p>
</BODY>
</HTML>