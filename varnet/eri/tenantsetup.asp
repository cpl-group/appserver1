<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'Response.Redirect "http://www.genergyonline.com"
		else
			if Session("eri") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				Response.Redirect "../main.asp"
			end if	
		end if		
		
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function openpopup(){
//configure "Open Logout Window
parent.document.location.href="../index.asp";
}
function loadpopup(){
openpopup()
}
</script>
<STYLE>
<!--
A.ssmItems:link		{color:black;text-decoration:none;}
A.ssmItems:hover	{color:black;text-decoration:none;}
A.ssmItems:active	{color:black;text-decoration:none;}
A.ssmItems:visited	{color:black;text-decoration:none;}
//-->
</STYLE>

<script>
function addnew(){
// Load Add new tenant form
var bldg=document.choosebldg.bldg.value
document.frames.title.location.href="null.htm";
document.frames.piclist.location.href="ti_add.asp?bldg=" + bldg;
document.frames.info.location.href="tenantlist.asp?bldg=" + bldg;
}
function lookup(){
	document.frames.title.location.href="title.asp?bldg="+document.forms[0].bldg.value
}
</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC"> 
      <div align="center"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif">ERI 
        Manager - Tenant Setup</font></b></font></div>
    </td>
  </tr>
</table>
<table width="100%" border="0" height="100%">
  <tr> 
    <td valign="top" align="right" width="79%" height="600"> 
      <%
		dim cnn1
		dim rst1
		
		Set cnn1 = Server.CreateObject("ADODB.Connection")
		Set rst1 = Server.CreateObject("ADODB.recordset")
		cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"
		rst1.Open "buildings", cnn1, 0, 1, 2
		%>
      <form target=title method="POST" action="tenantsetup.asp" webbot-action="--WEBBOT-SELF--" name="choosebldg">
        <!--webbot bot="SaveResults" U-File="_private/form_results.txt"
		  S-Format="TEXT/CSV" S-Label-Fields="TRUE" startspan --><strong>[FrontPage Save Results Component]</strong><!--webbot bot="SaveResults" endspan i-checksum="6561" -->
        </div>
        <p align="left"> &nbsp; 
          <select name="bldg">
            <%
		  while not rst1.EOF
		  response.write "<option value='" & rst1("bldgnum") & "'>" & rst1("strt") & "</option>" & vbCrLf
		  rst1.movenext
		  wend  
		  rst1.close  
		  set cnn1 = nothing
		  %>
          </select>
          <input type="button" name="Submit" value="Lookup" onclick="lookup()">
          <%	  if Session("eri") > 2 then %>
          <input type="button" name="Submit2" value="Add New Tenant" onclick="addnew()">
          <%	   end if  	%>
        </p>
      </form>
      <p align="center"> <IFRAME name="title" width="100%" height="160" src="null.htm" scrolling="no"></IFRAME> 
        <IFRAME name="piclist" width="100%" height="300" src="null.htm" scrolling="auto" marginwidth="8" marginheight="16"></IFRAME> 
        <IFRAME name="info" width="100%" height="300" src="null.htm" scrolling="auto" marginwidth="8" marginheight="16"></IFRAME> 
      </p>
    </td>
  </tr>
</table>
</body>
</html>
