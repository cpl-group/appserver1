<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
 '   if isempty(Session("name")) then
%>
<!--<script>
top.location="../index.asp"

</script>-->
<%
     ' Response.Redirect "http://www.genergyonline.com"
   ' else
    '  if Session("eri") < 2 then 
     '   Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

       ' Response.Redirect "../main.asp"
     ' end if  
   ' end if    
    
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">
<script>
function openpopup(){
//configure "Open Logout Window
parent.document.location.href="../index.asp";
}
function loadpopup(){
openpopup()
}

function reload(val){
  var temp="tenant_survey.asp?bldg="+val
    document.title.location=temp
}
</script>
<script language="JavaScript" type="text/javascript">
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
</script>
</head>

<body bgcolor="#eeeeee" text="#000000">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699CC"> 
  <td><span class="standardheader">ERI Manager | Tenant Survey</span></td>
</tr>
</table>
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#eeeeee"> 
  <td valign="top" style="border-bottom:1px solid #cccccc;"> 
  <%
  dim cnn1, rst1, sqlstr
  
  Set cnn1 = Server.CreateObject("ADODB.Connection")
  Set rst1 = Server.CreateObject("ADODB.recordset")
  cnn1.open getConnect(0,0,"Engineering")
  sqlstr = "select * from buildings order by strt"
  rst1.Open sqlstr, cnn1
  %>
  <select name="bldg" onChange=reload(this.value)>
  <%
  while not rst1.EOF%>
  <option <%if isBuildingOff(rst1("bldgnum")) then%>class="grayout"<%end if%> value="<%=rst1("bldgnum")%>"><%=rst1("strt")%></option><%
  rst1.movenext
  wend  
  rst1.close  
  set cnn1 = nothing
  %>
  </select>
  <input type="button" name="Button" value="View Building" onClick=reload(bldg.value)>
  </td>
</tr>
<tr valign="top">
  <td style="border-top:1px solid #ffffff;" height="560">
  <IFRAME name="title" width="100%" height="160" src="null.htm" scrolling="no" frameborder=0 border=0></IFRAME> 
  <IFRAME name="tenant" width="100%" height="85" src="null.htm" scrolling="no" marginwidth="0" marginheight="0" frameborder=0 border=0></IFRAME> 
  <IFRAME name="details" width="100%" height="280" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0></IFRAME> 
  </td>
</tr>
</table>
</body>
</html>
