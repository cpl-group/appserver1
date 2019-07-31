<%Option Explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Operations Log</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">
<script>
function bldgInfo(bldgnum) {

	var temp
	temp="capbldginfo.asp?bldgnum="+bldgnum
	document.frames.capacity.location=temp
}


</script>
</head>
<%
dim bldgnum, items, sql
bldgnum = Request("bldgnum")
items = Request("items")
'response.write(items)
Dim cnn1, rst1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"engineering")

sql="select b.bldgnum, strt as address from tlbldg tl inner join ["&Application("CoreIP")&"].dbCore.dbo.buildings b on tl.bldgnum=b.bldgnum order by strt"

rst1.Open sql, cnn1, 0, 1, 1
if not rst1.eof then


%>

<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC"> 
      <div align="center"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif">Power 
        Capacity Setup</font></b></font></div>
    </td>
  </tr>
</table>
<form name="form1">
<table width="100%" border="0" align="center">
  <tr> 
    <td align="left" height="36"> 
        <font face="Arial, Helvetica, sans-serif">
        Search for Building 
		</font>
		
        <select name="bldgnum" size="1" onChange='bldgInfo(this.value)'>
		  <option>========</option>
		  <%
		  do until rst1.eof
		  %>
		 
		  
            <option  <%if isBuildingOff(rst1("bldgnum")) then%>class="grayout"<%end if%> value="<%=Trim(rst1("bldgnum"))%>"><font face="Arial, Helvetica, sans-serif"><%=rst1("address")%>
					
		  </font>
		  <%
		  rst1.movenext
		  loop
		  %> 
        </select>
        <font face="Arial, Helvetica, sans-serif">
        <input type="button" name="Submit3" value="New Building" onClick='javascript:capacity.location="capnewbldg.asp"'>
        </font> 
    </table>
</form>
<IFRAME name="capacity" width="100%" height="100%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME>

</body>
<%
end if
rst1.close
%>
</html>