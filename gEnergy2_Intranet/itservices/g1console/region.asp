<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
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
function movetoregion(){
	document.form1.mode.value = "updatebranch2a"
	document.form1.submit()
}
</script>
<body bgcolor="#eeeeee" text="#000000">
	
<form name="form1" action="leafupdates.asp" method="post">
  <table width="100%" border="0" cellpadding="3" cellspacing="0">
    <tr> 
      <td width="48%"> <strong><span style="color: #003399">Region Label</span></strong></td>
      <td width="52%"> <span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
        <input name="region" type="text" value="<%=region%>" size="50%">
        </span></td>
    </tr>
    <tr> 
      <td width="48%"> <strong><span style="color: #003399">Region Count</span></strong></td>
      <td> <span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
        <input type="hidden" name="pregioncount" value = "<%=regioncount%>">
        <%=regioncount%></span></td>
    </tr>
    <tr> 
      <td><strong></strong></td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td colspan=2><a href="javascript:document.form1.submit()">Save Update</a><span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
        <input type="hidden" name="mode" value = "updatebranch2">
        | <a href="addentries.asp?mode=delbranch2&lid=<%=regioncount%>" target="_self">Delete 
        Region</a> 
        <% if trim(region) <> trim(bldgname) then%>
        <% if customlabel = "-1" then %>
        | <a href="addentries.asp?mode=srvorder&lid=<%=regioncount%>" target="_self">Order 
        Region by Service</a> 
        <%else%>
        | <a href="addentries.asp?mode=bldgorder&lid=<%=regioncount%>" target="_self">Order 
        Region by Building</a> 
        <%end if%>
        <%end if%>
        </span> <a href="addentries.asp?mode=moveregion&lid=<%=regioncount%>&mv_dir=up&catid=<%=catid%>" target="_self">Move 
        Up</a> | <a href="addentries.asp?mode=moveregion&lid=<%=regioncount%>&mv_dir=down&catid=<%=catid%>" target="_self">Move 
        Down</a> </td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td colspan=2> 
        <% if category ="-1" then %>
        <a href="addentries.asp?mode=rmts&lid=<%=regioncount%>" target="_self">Make Child</a> 
        | 
        <%
		else
		%>
        <a href="addentries.asp?mode=rmtp&lid=<%=lid%>" target="_self">Make Parent</a> 
        | 
        <%end if%>
        <a href="addentries.asp?mode=addbranch2&lid=<%=lid%>" target="_self">Add 
        Region</a> | 
		<% if trim(customlabel) <> trim(bldgname) then %>
        <a href="addentries.asp?mode=addbranch1&lid=<%=lid%>" target="_self">Add 
        Branch</a> | 
        <% end if %>
        <a href="addentries.asp?mode=delbranch2&lid=<%=regioncount%>" target="_self"> 
        </a>
        <% 
		strsql = "SELECT  Distinct region, regioncount FROM clientsetup WHERE userid = '" & userid & "' and catid =" & catid
		rs.Open strsql, cnn1, adOpenStatic
		if not rs.eof then 
		%>
        <strong><span style="color: #003399">Move to Region</span></strong> <span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
        <select name="regioncount" onChange="movetoregion()">
          <% while not rs.eof%>
          <option value="<%=rs("regioncount")%>" <%if trim(rs("regioncount")) =trim(regioncount) then%>selected<%end if%>><%=rs("region")%></option>
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
    <tr> 
      <td>&nbsp; </td>
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
</body>

</html>
