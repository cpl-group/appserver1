<%@Language="VBScript"%>
<!--#INCLUDE Virtual="/genergy2/secure.inc"-->
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
<% 	
lid = request("lid")	
userid = session("editemail")
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set userdetails = Server.CreateObject("ADODB.recordset")
Set rs = Server.CreateObject("ADODB.recordset")

cnn1.open getConnect(0,0,"dbCore")

strsql = "SELECT  * FROM clientsetup WHERE id = '" & lid & "'"
'response.write strsql

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
end if 
userdetails.close

%>
<script>
function movetocategory(){
	document.form1.mode.value = "updatebranch3a"
	document.form1.submit()
}
</script>

<body bgcolor="#eeeeee" text="#000000">
	
<form name="form1" action="leafupdates.asp" method="post">
  <table width="100%" border="0" cellpadding="3" cellspacing="0">
    <tr> 
      <td width="48%"> <strong><span style="color: #003399">Category Label</span></strong></td>
      <td width="52%"> <span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
        <input name="category" type="text" value="<%=category%>" size="50%">
        </span></td>
    </tr>
    <tr> 
      <td width="48%"> <strong><span style="color: #003399">Category ID</span></strong></td>
      <td> <span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
        <input name="pcatid" type="hidden" value="<%=catid%>"><%=catid%>
        </span></td>
    </tr>
    <tr> 
      <td colspan=2><a href="javascript:document.form1.submit()">Save Update</a><span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
        <input type="hidden" name="mode" value = "updatebranch3">
        | <a href="addentries.asp?mode=delbranch3&lid=<%=catid%>" target="_self">Delete 
        Category</a> | <a href="addentries.asp?mode=movecategory&catid=<%=catid%>&mv_dir=up" target="_self">Move 
        Up</a> | <a href="addentries.asp?mode=movecategory&catid=<%=catid%>&mv_dir=down" target="_self">Move 
        Down</a> </span></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td colspan = 2><a href="addentries.asp?mode=addbranch3&lid=<%=lid%>" target="_self">Add 
        Category</a> | 
        <% 
		strsql = "SELECT  Distinct category, catid FROM clientsetup WHERE userid = '" & userid & "' and category <> '-1'"
		rs.Open strsql, cnn1, adOpenStatic
		if not rs.eof then 
		%>
        <strong><span style="color: #003399">Move to Category </span></strong> 
        <span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
        <select name="catid" onchange="movetocategory()">
          <% while not rs.eof%>
          <option value="<%=rs("catid")%>" <%if trim(rs("category")) =trim(category) then%>selected<%end if%>><%=rs("category")%></option>
          <% rs.movenext
			wend
		%>
        </select>
        </span> 
        <%
		end if
		%>
      </td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
</body>

</html>
