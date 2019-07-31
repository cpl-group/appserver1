<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<title>Meter Details - Usage History</title>
</head>
<%
bldg = Request("B")
if trim(bldg)="" then bldg = Request("Bldg")
if trim(bldg)="" then bldg = Request("building")
meterid = request("M")
if trim(meterid)="" then meterid = request("meterid")
if  Instr(meterid,"SVR") then meterid = split(meterid,"-")(1) 
%>
<script> 
function loadentry(luid, ypid){

	var temp = 'pgibill.asp?b=<%=bldg%>&l=' +luid+'&Y='+ypid

	
	document.location = temp
}
function lmp(bldg,meterid) {
	theURL="/genergy2/eri_th/lmp/lmp.asp?hideOptions=true&meterid=" + meterid+"&bldg="+bldg+"&lmp=1&utility=2&interval=0"
	openwin(theURL,800,475)
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}

function viewBill(){
	//buildpdf = (buildpdf!=false?true:false);
  var url = "http://pdfmaker.genergyonline.com//pdfmaker/pdfBatchPrint.asp?genergy2=true&devIP=appserver1.genergy.com&billurl=&building=919&logo=invoice_logo_1.jpg&logoh=0&logow=0&pid=24&lid=3057&bperiod=5&byear=2008&y=51931&ypid=51931&l=3057&detailed=false&utilityid=2&buildpdf=true";
  //alert(url);
  billpdf = window.open(url,'','width=600,height=500,resizable=yes');
}
</script>
<body bgcolor="#FFFFFF" onLoad="top.window.focus()">
<%

Set cnn1 = Server.CreateObject("ADODB.Connection")
set cmd1 = server.createobject("ADODB.Command")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open getLocalConnect(bldg)

sql="SELECT top 1  ypid FROM tblbillbyperiod where reject=0 and bldgnum='" & bldg & "' group by ypid order by ypid desc"

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
%>
<table width="100%" height="14%" border="0" cellspacing="0" cellpadding="0">
  <%sql="SELECT meternum from meters where meterid=" & meterid &""
 rst2.Open sql, cnn1, adOpenStatic, adLockReadOnly%>
  <tr> 
    <td bgcolor="#336699" width="23%" height="50%"><span class=standardheader><b><font color="#FFFFFF" size="6%">Meter 
      Details for <%=rst2("meternum")%></font></b></b><font color="#FFFFFF" size="4">Usage 
      Information</font></span></td>
    <td bgcolor="#336699" width="23%" align="right"><a href="Javascript:lmp('<%=bldg%>','<%=meterid%>')" style="text-decoration:none;color:white"><b>View 
      Load Profile</b></a></td><b></b><b></b>
  </tr>
  <tr> 
    <td width="46%" Height="24%" colspan="2" cellspacing="6" >&nbsp</td>
  </tr>
  <%rst2.close%>
</table>
<table border="1" width="100%" height="2" cellspacing="0" cellpadding="0">
  <tr bgcolor="#6699CC"> 
    <td width="14%" height="1%" align="center"><b><font size="2" face="Arial">Bill Period</font></b></td> 
    <td width="14%" height="1%" align="center"><b><font size="2" face="Arial">Date Last Read</font></b></td> 
    <td width="8%" height="1%" align="center"><b><font size="2" face="Arial">Previous</font></b></td>
    <td width="20%" height="1%" align="center"><b><font size="2" face="Arial">Current</font></b></td>
    <td width="13%" height="1%" align="center"><b><font size="2" face="Arial">On Peak</font></b></td>
    <td width="12%" height="1%" align="center"><b><font size="2" face="Arial">Off Peak</font></b></td>
    <td width="13%" height="1%" align="center"><b><font size="2" face="Arial">Kwhr</font></b></td>
    <td width="12%" height="1%" align="center"><b><font size="2" face="Arial">Demand</font></b></td>
  </tr>
</table>
  <table border="0" width="100%" height="1" cellpadding="0" cellspacing="0">
<%
while not rst1.eof
cmd1.ActiveConnection = cnn1
cmd1.CommandText = "select * from tblmetersbyperiod m, tblbillbyperiod b where reject = 0 and m.meterid=" & Meterid & " and m.bill_id=b.id and b.ypid=" & rst1("ypid")
cmd1.CommandType = 1
Set rst2 = cmd1.Execute
if not rst2.eof then
%>

<div align="left">

    <tr valign="top" onMouseOver="this.style.backgroundColor = 'lightgreen'" onMouseOut="this.style.backgroundColor = 'white'" onClick="javascript:viewBill()"> 
 <td width="14%" height="1%" align="center">
      <div align="right"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=rst2("billyear")%>/<%=rst2("billperiod")%>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp<%=rst2("datelastread")%></font></b></div></td>
      
      <td width="8" height="1%" align="right"><b><font face="Arial, Helvetica, sans-serif" size="2"></font></b></td>
      <td width="8%" height="1%" align="right"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst2("Prev"),0)%></font></b></td>
      <td width="20%" height="1%" align="right"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst2("current"),0)%></font></b></td>
      <td width="13%" height="1%" align="right"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst2("onpeak"),0)%></font></b></td>
      <td width="12%" height="1%" align="right"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst2("offpeak"),0)%></font></b></td>
      <td width="13%" height="1%" align="right"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst2("used"),0)%></font></b></td>
      <td width="12%" height="1%" align="right"><b><font size="2" face="Arial, Helvetica, sans-serif"><%=formatnumber(rst2("demand_P"),2)%></font></b></td>
    </tr>


</div>


<%
rst1.movenext
else
rst1.movenext
end if
wend
set cnn1 = nothing
%>
<tr>
    <td colspan = 8 style="border-top: #000000 solid 2px;"><font face="Arial, Helvetica, sans-serif" size="2"><b>Click bill period above for detailed billing information</b></td>
</tr></font>
  </table>
</body>

</html>
