<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<script>
function process(item, flag, d, job){
	var count=document.temp.count.value-1
	document.location="corpinvoicefilter.asp?job="+job+"&flag="+flag+"&date="+d+"&item="+item
}
</script>
</head>
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
cnn1.Open getConnect(0,0,"engineering")
bldgnum=request("bldgnum")

sql="select address, sqft, rev_date from tlbldg where bldgnum='"& bldgnum &"'"

rst1.Open sql, cnn1, 0, 1, 1
%>

<body bgcolor="#FFFFFF" text="#000000">
<%
if not rst1.eof then
	
%>

<table width="100%" border="0" bgcolor="#3399CC">
  <tr >
    <td height="2" width="71%"><font face="Arial, Helvetica, sans-serif"><i><b> 
      Capacity </b></i></font> </td>
  </tr>
</table> 

<div align="right"> </div>

<table width="100%" border="0"> 
  <tr bgcolor="#CCCCCC">
    <td width="29%"><font face="Arial, Helvetica, sans-serif">Building Address</font></td>
    <td width="21%"><font face="Arial, Helvetica, sans-serif">SQFT</font></td>
  </tr>
  <%
      address=trim(rst1("address"))
	  sqft=trim(rst1("sqft"))
	  revdate=trim(rst1("rev_date"))
  %>

  <tr><form>
      <td width="29%"><a href="capbldginfo.asp?address=<%=address%>&bldgnum=<%=bldgnum%>&sqft=<%=sqft%>&revdate=<%=revdate%>"><font face="Arial, Helvetica, sans-serif"><%=address%></font></a></td>
      <td width="21%"><font face="Arial, Helvetica, sans-serif"><%=sqft%></font></td>
  </form></tr>
  
  <%
end if
rst1.close
  %>
  
</table>
</body>
</html>
