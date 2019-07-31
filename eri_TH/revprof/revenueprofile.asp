<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit %>
<!--#include file="adovbs.inc"-->
<%
Dim bldg
bldg=Request.QueryString("bldgnum")
Dim year
year=Request.QueryString("year")
Dim Bldgname
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

Dim StackTitle(4) 
StackTitle(4) = "Net"
StackTitle(2) = "Submeter"
StackTitle(3) = "ERI"
StackTitle(1) = "Expenses"

Dim i(12, 4) 'i(Bar#, Stack) (1-New, 2-Cancelled, 3-Frozen)
Dim Values(12, 4)
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

strsql = "SELECT BillYrPeriod.ypId,eri_data.dbo.eri_total.eri_amt AS ERI_rev, UtilityBill.TotalBillAmt AS Expenses,SUM(tblBillByPeriod.TotalAmt) AS SubMetered FROM eri_data.dbo.eri_total full JOIN BillYrPeriod ON eri_data.dbo.eri_total.bldg_no = BillYrPeriod.BldgNum AND eri_data.dbo.eri_total.BillPeriod = BillYrPeriod.BillPeriod AND eri_data.dbo.eri_total.BillYear = BillYrPeriod.BillYear full JOIN UtilityBill ON BillYrPeriod.ypId = UtilityBill.ypId full JOIN   tblBillByPeriod ON UtilityBill.ypId = tblBillByPeriod.ypId WHERE (BillYrPeriod.BldgNum = '"& bldg &"') AND (BillYrPeriod.BillYear = '"& year &"') GROUP BY eri_data.dbo.eri_total.eri_amt, UtilityBill.TotalBillAmt, BillYrPeriod.ypId order by billyrperiod.ypid"


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
	if rst1("Expenses") <> "" then i(x, 1) = rst1("Expenses") else i(x,1)=0 end if
	if rst1("Submetered") <> "" then i(x, 2) = rst1("Submetered") else i(x,2)=0 end if
	if rst1("eri_rev") <> "" then i(x, 3) = rst1("eri_rev") else i(x,3)=0 end if
	i(x, 4) = (i(x,2) + i(x,3)) - i(x,1)
	tots = rst1("Expenses")	
rst1.MoveNext
tot=tot+tots
next

	rst1.Close
	Set rst1 = Nothing

	cnn1.Close
	Set cnn1 = Nothing

for x= numRecords+1 to 12
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
NumStacks = 4     'Must have color(from 1 to numbars).gif
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

<BODY>
<%
x=0
Dim r
Dim s
Dim t 'Counter Variables

Dim step 'Stack height

for x = 1 to numbars
for r = 1 to numstacks
values(x, r) = i(x, r)
i(x, r) = ((i(x, r) / Maxvalue) * 1000)
next
next
%>

<%
t = 0
Dim PosLeft
PosLeft = SidePadding + 30
for x = 1 to numbars
for r = 1 to numstacks
t = t + 1 
	Step = 0
	If r =2  then
		For s = 1 to (r-1)
			Step = Step + i(x, 3)
		Next
	End If

if i(x, r) <> 0 then
%>
<DIV CLASS=Bargraph ID=bg<%=t%>>
<%if r = 1 then%>
<img src="color<%=r%>.gif" Alt=<%=FormatCurrency(Values(x, r))%> Height=<%=(i(x, r) + 3)%> Width=<%=Barwidth  + 10%> border="1" style="border-style: solid; border-color: #0000FF" class="bar" id=section<%=r%>>
<%else%>
<%if values(x,r) > 0 then  
if r <> 4 then %>
<img src="color<%=r%>.gif" Alt=<%=FormatCurrency(Values(x, r))%> Height=<%=(i(x, r) + 4)%> Width=<%=Barwidth%> border="1" style="border-style: solid; border-color: #0000FF" class="bar" id=section<%=r%>>
<%
end if
end if
end if%>
</DIV>


<SCRIPT LANGUAGE="javascript">
		var BarGraphOBJ = eval("document.all['bg<%=t%>']");
		BarGraphOBJ.style.posLeft = (<%if r=1 then Response.Write PosLeft-5 else response.write PosLeft end if%>);
		BarGraphOBJ.style.posTop = (<%=(400 - i(x, r) + TopPadding) - Step%>);
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
%>

<DIV CLASS=Bargraph ID=line1>
<img src="line.gif" width=<%=numbars * barpadding + 90%> height=2> 
</DIV>
<SCRIPT LANGUAGE="javascript">
		var BarGraphOBJ = eval("document.all['line1']");
		BarGraphOBJ.style.posLeft = (<%=Sidepadding - 65%>);
		BarGraphOBJ.style.posTop = (400 + <%=TopPadding%> - 7);
		BarGraphOBJ.style.visibility = "visible";
</SCRIPT>

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
<% if x = 4 then %>
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
<%end if%>
<p align="left">&nbsp;</p>
</BODY>
</HTML>