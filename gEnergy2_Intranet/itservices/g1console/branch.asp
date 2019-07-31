<%@Language="VBScript"%>
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
	bldgorder		= userdetails("bldgorder")
end if 
userdetails.close

%>
<body bgcolor="#eeeeee" text="#000000">
	
<form name="form1" action="leafupdates.asp" method="post">
  <table width="100%" border="0" cellpadding="3" cellspacing="0">
    <tr> 
      <td width="48%"><strong><span style="color: #003399">Branch Label</span></strong></td>
      <td width="52%"><span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
        <input name="bldgname" type="text" value="<%=bldgname%>" size="50%">
        </span></td>
    </tr>
    <tr> 
      <td width="48%"><strong><span style="color: #003399">Branch ID</span></strong></td>
      <td><span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
        <input name="bldgid" type="text" id="bldgid" value="<%=bldgid%>" size="50%">
        </span></td>
    </tr>
    <tr>
      <td><strong><span style="color: #003399">Branch Region</span></strong></td>
      <td><% 
		strsql = "SELECT  Distinct region, regioncount FROM clientsetup WHERE userid = '" & userid & "' and catid = " & catid
		rs.Open strsql, cnn1, adOpenStatic
		if not rs.eof then 
		%><span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
        <select name="regioncount">
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
        <span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
        </span> </td>
    </tr>
    <tr> 
      <td colspan=2><a href="javascript:document.form1.submit()"> Save Update</a><span class="standard" style="margin:4px;line-height:9pt;color:#003399;"> 
        <input type="hidden" name="mode" value = "updatebranch1">
        <input type="hidden" name="pBldgid" value = "<%=bldgid%>">
        <input type="hidden" name="pRegioncount" value = "<%=regioncount%>">
        </span> | <a href="addentries.asp?mode=delentry&lid=<%=lid%>"></a><a href="addentries.asp?mode=delbranch1&lid=<%=bldgid%>" target="_self">Delete 
        Branch</a> | <a href="addentries.asp?mode=movebuilding&lid=<%=bldgorder%>&mv_dir=up&regioncount=<%=regioncount%>" target="_self">Move 
        Up</a> | <a href="addentries.asp?mode=movebuilding&lid=<%=bldgorder%>&mv_dir=down&regioncount=<%=regioncount%>" target="_self">Move 
        Down</a> </td>
      <td>&nbsp;</td>
    </tr>
    <tr> 
      <td colspan =2> 
        <% if  region ="-1" then %>
        <a href="addentries.asp?mode=bmts&lid=<%=lid%>" target="_self">Move To 
        Sub</a> | 
        <%else %>
        <a href="addentries.asp?mode=bmtp&lid=<%=lid%>" target="_self">Move To 
        Parent</a> | 
        <%end if%>
        <a href="addentries.asp?mode=<% if trim(region) = "-1" then %>addbranch1a<%else%>addbranch1<% end if %>&lid=<%=lid%>" target="_self">Add 
        Branch</a> | <a href="addentries.asp?mode=addentry&lid=<%=lid%>">Add New 
        Entry</a> </td>
      <td> </td>
    </tr>
    <tr> 
      <td colspan =2>&nbsp; </td>
      <td></td>
    </tr>
  </table>
</form>
</body>

</html>
