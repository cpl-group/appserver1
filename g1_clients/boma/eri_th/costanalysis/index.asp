<html>
<head>
<title>Cost &amp; Revenue Analysis</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function loadprofiles(type,pid){
 	switch(type){
	case "Revenue Profiles":
		document.location="revbldglist.asp?pid="+pid
		break
	case "Costs Profiles":
		document.location="cost.asp?pid="+pid
		break
	}
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF">
  <tr>
    <td>
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">COST 
        &amp; REVENUE ANALYSIS</font></b></font></div>
    </td>
  </tr>
</table>
<p> 
  <input type="hidden" name="pid" value="<%=Request.Querystring("pid")%>">
  <input type="button" name="Button" value="Revenue Profiles" onclick="loadprofiles(this.value,pid.value)">
  <input type="button" name="Button" value="Costs Profiles" onclick="loadprofiles(this.value,pid.value)">
</p>
</body>
</html>
