<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include VIRTUAL="/genergy2/secure.inc" -->
<html>
<head>
<title>G1 Console QA Module</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- <STYLE>
<!--
A.ssmItems:link		{color:black;text-decoration:none;}
A.ssmItems:hover	{color:black;text-decoration:none;}
A.ssmItems:active	{color:black;text-decoration:none;}
A.ssmItems:visited	{color:black;text-decoration:none;}
//- ->
</STYLE> -->
</head>
<% 	
lid = request("lid")	
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set userdetails = Server.CreateObject("ADODB.recordset")
Set rs = Server.CreateObject("ADODB.recordset")

cnn1.open getConnect(0,0,"dbCore")

strsql = "SELECT  * FROM clientsetup WHERE id = '" & lid & "'"

userdetails.Open strsql, cnn1, adOpenStatic
if not userdetails.EOF then 
	category 		= userdetails("category")
	catid 			= userdetails("catid")		
	region			= userdetails("region")
	regioncount 	= userdetails("regioncount")
	bldgid			= userdetails("bldgid")
	bldgName		= userdetails("bldgName")
	serviceid 		= userdetails("serviceid")
	serviceurl		= userdetails("serviceurl")
	customlabel 	= userdetails("customlabel")
	servicetarget 	= userdetails("servicetarget")
	effdate		 	= userdetails("effdate")
	lockdate		= userdetails("lockdate")
	id 				= userdetails("id")
	' Added by Tarun 07/18/2008
	displayOrdSeq	= userdetails("ViewOrderSeq")
end if 

userdetails.close
		strsql = "SELECT  * FROM services order by servicelabel"
		
		rs.Open strsql, cnn1, adOpenStatic
		if not rs.EOF then 
		
		%>
<script language="javascript">
var urllist = new Array()
var paramlist = new Array()
 
<% 		x = 0	
		while not rs.EOF 
		if trim(rs("servicelink")) = "NA" then 
				servicelink = "Custom Link Required Below"
		else
			servicelink = rs("servicelink")
		end if
		%>
		urllist[<%=x%>] = '<%=servicelink%>' 
		paramlist[<%=x%>] = '<%=rs("parameters")%>'
<% 
		if trim(rs("id")) = trim(serviceid) then 
			currentService = rs("servicelink")
		end if
		x=x+1
		rs.movenext
		wend
		rs.movefirst
%>
function updateurl(){
	document.form1.urllink.value = urllist[document.form1.serviceid.selectedIndex]
	document.form1.serviceurl.value = paramlist[document.form1.serviceid.selectedIndex]
	
}
</script>
		<%
		end if 
%>
<link rel="Stylesheet" href="\genergy2\SETUP\setup.css" type="text/css">
<body bgcolor="#eeeeee" text="#000000">
	
<form name="form1" action="leafupdates.asp" method="post">
  <table width="100%" border="0" cellpadding="3" cellspacing="0">
    <tr> 
      <td width="48%"> <strong><span class="standard" style="margin:4px;line-height:9pt;color:#003399;">service
        <input type="hidden" name="id" value = "<%=id%>">
        </span></strong></td>
    </tr>
    <tr> 
      <td><span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
        <%
		if not rs.EOF then 
			
	%>
        <select name="serviceid" onChange="updateurl()">
          <% 
	while not rs.eof 
	%>
          <option value="<%=rs("id")%>" <% if cint(rs("id")) = cint(serviceid) then %> selected<%end if%>><%=rs("servicelabel")%> 
          (<%=rs("servicecode")%>)</option>
          <% 
	rs.movenext
	wend
	%>
        </select>
        <%
  		end if
	%>
        </span></td>
    </tr>
    <tr> 
      <td valign="top"><span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
        <input type="text" name="urllink" value="<%=currentService%>" size="100" style="width: 100%;">
        </span></td>
    </tr>
    <tr> 
      <td valign="top"><strong><span class="standard" style="margin:4px;line-height:9pt;color:#003399;">Current 
        Service URL</span></strong></td>
    </tr>
    <tr> 
      <td valign="top"><span class="standard" style="margin:4px;line-height:9pt;color:#000000;"><%=serviceurl%></span></td>
    </tr>
    <tr> 
      <td valign="top"><strong><span class="standard" style="margin:4px;line-height:9pt;color:#003399;">Edit 
        Service URL</span></strong></td>
    </tr>
    <tr> 
      <td valign="top"> <strong><span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
        <input type="text" name="serviceurl" size="100" style="width: 100%;" value="<%=serviceurl%>">
        </span></strong>
      </td>
    </tr>
	<tr> 
	<td> 
	<strong><span class="standard" style="margin:4px;line-height:9pt;color:#003399;">Custom Service Label</span></strong></td>
	</tr>
	<tr>
	<td>
	<input name="customlabel" type="text" size="50%" value = "<%=customlabel%>">
	</td>
	</tr>
	<tr> 
	<td> 
	<strong><span class="standard" style="margin:4px;line-height:9pt;color:#003399;">Display Order Sequence Number</span></strong></td>
	</tr>	
	<tr>
	<td>
	<input name="DisplayOrdSeq" type="text" size="50%" value = "<%=displayOrdSeq%>">
	</td>
	</tr>
    <tr> 
      <td><a href="javascript:document.form1.submit()">
        Save Update</a><span class="standard" style="margin:4px;line-height:9pt;color:#003399;">
<input type="hidden" name="mode" value = "updateleaf">
        | <a href="addentries.asp?mode=delentry&lid=<%=lid%>">Delete Entry</a> 
        </span></td>
    </tr>
    <tr> 
      <td> 
        <% if trim(bldgid) ="-1" then %>
        <a href="addentries.asp?mode=mts&lid=<%=lid%>" target="_self">Move To 
        Sub</a> 
        <%end if%>
      </td>
    </tr>
  </table>
</form>
</body>

</html>
