<%Option Explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim cnnSuper, sql,rst
set cnnSuper = server.createobject("ADODB.Connection")
set rst		 = server.createobject("ADODB.Recordset")

cnnSuper.open getConnect(0,0,"dbCore")

sql ="select bldgnum, ip from rm where enable = 1 and lm = 1"
rst.open sql, cnnSuper

%>
<html>
<head>
<title>Navigation Bar</title>
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">

<script>

function viewEFX(){
	var ip = document.buildingpicker.bldgnum.value
	var url = "http://10.0.8.225/cgi-bin/eso.tcl?ip="+ip
	document.frames.info.location.href=url
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body bgcolor="#eeeeee" onunload="try{editpane.close();}catch(exception){}">

<form name="buildingpicker" method="get">
<table border=0 height= "30" cellpadding="2" cellspacing="0" width="100%" bgcolor="336699" style="border-right:2px solid #000000;">
	<tr><td width="1%" style="border-top:1px solid #99ccff;border-right:1px solid #000000;">
				<select name="bldgnum" onchange="viewEFX();">
					<%
					if not rst.eof then
						do until rst.eof
							%>
							<option value="<%=rst("ip")%>"><%=rst("bldgnum")%></option>
							<%
							rst.movenext
						loop
					end if
					rst.close
					set rst= nothing
					%>
				</select>
        <input type="button" name="Button" value="View" onclick="viewEFX()"></td>
	</tr>
</table>  
</form>
<IFRAME name="info" width="100%" height="90%" src="" scrolling="auto" marginwidth="0" marginheight="0" frameborder=1 border=1> </IFRAME> 
</body>
</html>
