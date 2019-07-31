<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<html>
<head>
<title>G1 Console QA Module</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<STYLE>
<!--
A.ssmItems:link		{color:black;text-decoration:none;}
A.ssmItems:hover	{color:black;text-decoration:none;}
A.ssmItems:active	{color:black;text-decoration:none;}
A.ssmItems:visited	{color:black;text-decoration:none;}
//-->
</STYLE>
</head>
<body bgcolor="#eeeeee" text="#000000">
	<table width="100%" border="0" cellpadding="3" cellspacing="0">
	  <tr> 
		
    <td bgcolor="#6699cc"><span class="standard" style="margin:4px;line-height:9pt;color:#003399;">Loading gEnergyOne version <%=Request("version")%> for user <%=request("userid")%></span></td>
	  </tr>
	</table>
</body>
<%
Session("loginemail")=Request("userid")
Session("userid")=Request("userid")
response.redirect trim(Request("version"))	
%>
</html>
