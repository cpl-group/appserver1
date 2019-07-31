<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<script language="JavaScript" type="text/javascript">
var lastrow = "";
function redirect(bldgnum, f, currentrow){
   parent.frames.riser.location="capriser.asp?bldgnum="+bldgnum+"&floor="+f  
   parent.frames.detail.location="capdetail.asp?bldgnum="+bldgnum+"&floor="+f+"&item=floor"
  if (currentrow != lastrow) { 
    if (lastrow != "") { lastrow.style.backgroundColor = "white" }
    currentrow.style.backgroundColor = "#ccffcc"; 
    }
  lastrow = currentrow;
  
}
function rowOver(targetrow){
  targetrow.style.backgroundColor = "lightgreen";
}

function rowOut(targetrow){
  var tempcolor = "white";
  if (targetrow == lastrow) { tempcolor = "#ccffcc"; }
  targetrow.style.backgroundColor = tempcolor;
}
</script>

<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<script language="JavaScript" type="text/javascript">
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">

<%

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"engineering")

bldgnum=secureRequest("bldgnum")
riser=secureRequest("riser")
wsqft=0
if riser="" then
    
    
  sql = "select * from tblfloor where bldgnum='"& bldgnum &"' order by orderno"
  label = "All floors in this building"
else
    sql = "select distinct a.fl_name,f.sqft, a.wsqft,f.orderno,f.include  from tblassociation a join tblfloor f on a.bldgnum=f.bldgnum and a.fl_name=f.fl_name where a.riser_name='"& riser &"' and a.bldgnum='"& bldgnum &"' order by f.orderno"
  label ="Floors associated with riser "& riser
end if
rst1.Open sql, cnn1, 0, 1, 1

if rst1.eof then
%>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
  <tr>
    <td>
      <div align="center"> 
        <p><font face="Arial, Helvetica, sans-serif"><i>
    <%
    if riser="" then
    %>
    No Floor listed for the building </i></font></p>
        <%
    else
    %>
        No Floor for this Riser 
        <%
    end if
    %>
      </div>
    </td>
  </tr>
</table>
<%
else
%>

<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#eeeeee">
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><b><%=label%></b></td>
</tr>
</table>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#dddddd" style="font-weight:bold;"> 
  <td width="30%">Floor</td> 
   <td width="30%">Order No</td> 
  <td width="30%">SQFT</td>
  <td width="40%">WSQFT</td>
</tr>
<% 
do until rst1.EOF 
%>
<form name="form1" method="post" action="">
<% floor=trim(rst1("fl_name")) %>
<tr onMouseOver="rowOver(this)" style="cursor:hand" onMouseOut="rowOut(this)" onClick="redirect('<%=bldgnum%>', '<%=replace(floor,"'","\'")%>',this)"> 
  <td>  
  <input type="hidden" name="floor" value="<%=trim(rst1("fl_name"))%>">
  <%=floor%>
  </td>
  <td><%=rst1("orderno")%></td>
  <td><%=rst1("sqft")%></td>
  <td>  
  <%
  if riser <> "" then
      wsqft=rst1("wsqft")
  else
    sql2="select sum(wsqft)as wsqft from tblassociation where bldgnum='"& bldgnum &"' and fl_name='"& replace(floor,"'","''") &"' group by fl_name"
    rst2.Open sql2, cnn1, 0, 1, 1
    if not rst2.eof then
      wsqft=rst2("wsqft")
    end if
    rst2.close
  end if
  %>   
  <%=wsqft%>
  </td>  
</tr>
</form>
<%
rst1.movenext
loop
%>
</table>
<%
end if
rst1.close
set cnn1=nothing
%>
</body>
</html>
