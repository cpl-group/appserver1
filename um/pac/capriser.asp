<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<script language="JavaScript" type="text/javascript">
var lastrow = "";
function redirect(bldgnum, r, currentrow){
  parent.frames.detail.location="capdetail.asp?bldgnum="+bldgnum+"&riser="+r+"&item=riser"  
  parent.frames.floor.location="capfloor.asp?bldgnum="+bldgnum+"&riser="+r  
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

cnn1.Open getConnect(0,0,"engineering")

bldgnum=secureRequest("bldgnum")
floor=secureRequest("floor")
if floor="" then
	sql = "select *,voltage_drop*100 as vpercent from tblriser where bldgnum='"& bldgnum &"' order by riser_name"
	label ="All risers in this building"
else
	sql = "select distinct a.riser_name, r.* ,r.voltage_drop*100 as vpercent from tblassociation a join tblriser r on a.bldgnum=r.bldgnum and a.riser_name=r.riser_name where a.fl_name='"&floor&"' and a.bldgnum='"& bldgnum&"' order by a.riser_name"
	label = "Risers associated with floor "&floor
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
		if floor="" then
		%>
          No Risers in this building 
          <%
		else
		%>
          No Riser for this Floor 
          <%
		end if
		%>
          </i></font></p>
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
  <tr valign="bottom" bgcolor="#dddddd" style="font-weight:bold;"> 
    <td width="12%">Riser Name</td>
    <td width="2%">Size</td>
    <td width="6%">Metal</td>
    <td width="7%">Insulation</td>
    <td width="6%">Sets</td>
    <td width="6%">Volts</td>
    <td width="6%">Sw Frame</td>
    <td width="6%">Sw Fuse</td>
    <td width="6%">Power Factor</td>
    <td width="6%">Avg Length</td>
    <td width="6%">Wire Capacity</td>
    <td width="10%">Power Capacity</td>
    <td width="6%">Amps</td>
	<td width="10%">Voltage Drop</td>
    <td width="15%">Note</td>
  </tr>
  <% 
  dim rowcolor
	do until rst1.EOF 
  %>
  <form name="form1" method="post" action="">
    <tr onMouseOver="rowOver(this);" style="cursor:hand" onMouseOut="rowOut(this);" onClick="redirect('<%=bldgnum%>', '<%=trim(rst1("riser_name"))%>',this);"> 
      <td><%if rst1("volts")>210 then %><font color="#FF0000"><%else%><font color="#0000FF"><%end if%><%=rst1("riser_name")%></font></td>
      <td><%if rst1("volts")>210 then %><font color="#FF0000"><%else%><font color="#0000FF"><%end if%><%=rst1("size")%> </font></td>
      <td><%if rst1("volts")>210 then %><font color="#FF0000"><%else%><font color="#0000FF"><%end if%><%=rst1("metal")%></font></td>
      <td><%if rst1("volts")>210 then %><font color="#FF0000"><%else%><font color="#0000FF"><%end if%><%=rst1("insulation")%> </font></td>
      <td><%if rst1("volts")>210 then %><font color="#FF0000"><%else%><font color="#0000FF"><%end if%><%=rst1("sets")%></font></td>
      <td><%if rst1("volts")>210 then %><font color="#FF0000"><%else%><font color="#0000FF"><%end if%><%=rst1("volts")%> </font></td>
      <td><%if rst1("volts")>210 then %><font color="#FF0000"><%else%><font color="#0000FF"><%end if%><%=rst1("sw_frame")%> </font></td>
      <td><%if rst1("volts")>210 then %><font color="#FF0000"><%else%><font color="#0000FF"><%end if%><%=rst1("sw_fuse")%></font></td>
      <td><%if rst1("volts")>210 then %><font color="#FF0000"><%else%><font color="#0000FF"><%end if%><%=rst1("power_factor")%> </font></td>
      <td><%if rst1("volts")>210 then %><font color="#FF0000"><%else%><font color="#0000FF"><%end if%><%=rst1("riser_length")%></font></td>
      <td><%if rst1("volts")>210 then %><font color="#FF0000"><%else%><font color="#0000FF"><%end if%><%=rst1("wire_capacity")%> </font></td>
      <td><%if rst1("volts")>210 then %><font color="#FF0000"><%else%><font color="#0000FF"><%end if%><%=rst1("p_capacity")%></font></td>
      <td><%if rst1("volts")>210 then %><font color="#FF0000"><%else%><font color="#0000FF"><%end if%><%=rst1("amps")%></font></td>
	  <td><%if rst1("volts")>210 then %><font color="#FF0000"><%else%><font color="#0000FF"><%end if%><%=rst1("vpercent")%>%</font></td>
      <td><%if rst1("volts")>210 then %><font color="#FF0000"><%else%><font color="#0000FF"><%end if%><%=rst1("note")%></font></td>
      
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
