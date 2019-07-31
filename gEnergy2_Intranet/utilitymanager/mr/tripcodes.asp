<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
server.ScriptTimeout=300

dim tripcode, bldgnum 
dim whereStr

tripcode = request("tripcode")
bldgnum = request("bldgnum")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rs	 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"dbCore")

if (tripcode <> "") then whereStr = whereStr + "where tripcode="+tripcode + " "
if (bldgnum <> "" and bldgnum <> "All") then whereStr = whereStr + " and a.bldgnum='" + bldgnum + "'"

sqltemp = "(select * from buildings where bldgnum in (select bldgnum from dbo.super_tripcodes))" 
sqlstr = "select tripcode, a.bldgnum, strt,utility, utilityid from " & sqltemp & " a inner join dbo.super_tripcodes s on a.bldgnum = s.bldgnum inner join  dbo.tblutility u on s.uid = u.utilityid " + whereStr + " order by tripcode, a.bldgnum" 

rst1.Open sqlstr, cnn1, 0, 1, 1
%>
<html>
<head>
<title></title>
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">
<style type="text/css">
.tblunderline { border-bottom:1px solid #cccccc; }
</style>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"><head>
<script>
function updateentry(key){
	parent.document.frames.tripedit.location = "tripdetail.asp?key=" + key
	parent.document.all.te.style.visibility="visible"

}
function deletetrip(tripdate, tripcode){

	if (confirm("Delete trip date "+tripdate+" for trip code "+tripcode+"?")){
	parent.document.frames.tripedit.location = "tripmodify.asp?tripcode=" + tripcode +"&modify=DeleteAllTrips&tripdate=" + tripdate
	parent.document.all.te.style.visibility="hidden"
	}
}
function showall(){
	var func = eval('document.all.alltrips')
	func.innerHTML = (func.innerHTML == '[-]' ? '[+]' : '[-]');
	var displaytype = (func.innerHTML != '[-]' ? 'none':'block');
	var tag = document.all//('note162');
	for (i = 0; i < tag.length; i++){
		if (tag[i].name == 'tripset') tag[i].style.display = displaytype
		if (tag[i].name == 'tripfunc') tag[i].innerHTML = func.innerHTML
	} 
}
function trip(id){
	var tag = document.getElementById(id) 
	tag.style.display = (tag.style.display == "block" ? "none" : "block");
	var func = eval('document.all.func'+id)
	func.innerHTML = (func.innerHTML == '[-]' ? '[+]' : '[-]');
}
</script>
<body bgcolor="#eeeeee" leftmargin="0" topmargin="0" class="innerbody"> 
<table border=0 cellpadding="3" cellspacing="0" bgcolor="#ffffff" width="100%">
  <tr bgcolor="#ffffff" valign="middle"> 
    <td width="2%" align="center" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">&nbsp;</td>
    <td width="10%" align="center" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">T.C.</td>
    <td width="10%" align="center" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Building 
      Number</td>
    <td width="33%" align="left" bgcolor="#eeefff" class="tblunderline">Building 
      Address</td>
    <td width="5%" align="left" bgcolor="#eeefff" class="tblunderline" ><div align="center">Utility</div></td>
  </tr>
</table>  
<%  
if not rst1.eof then
	Do until rst1.EOF 
%> 
<table border=0 cellpadding="3" cellspacing="0" bgcolor="#ffffff" width="100%">
  <tr <%if isBuildingOff(rst1("bldgnum")) then%>class="grayout"<%end if%> bgcolor="#ffffff" valign="middle"> 
    <td width="2%" align="center" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%if not(isBuildingOff(rst1("bldgnum"))) then%><span id="func<%=rst1("bldgnum")%><%=rst1("utilityid")%>" name = "tripfunc" style="cursor:hand;text-decoration:none;" onclick="trip('<%=rst1("bldgnum")%><%=rst1("utilityid")%>')"><%end if%>[+]</span></td>
    <td width="10%" align="center" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("tripcode")%></td>
    <td width="10%" align="center" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("bldgnum")%>&nbsp;</td>
    <td width="33%" align="left" bgcolor="#eeefff" class="tblunderline"><%=rst1("strt")%>&nbsp;</td>
    <td width="5%" align="left" bgcolor="#eeefff" class="tblunderline"> <%=rst1("utility")%>&nbsp;</td>
  </tr>
</table>
<div id="<%=rst1("bldgnum")%><%=rst1("utilityid")%>" name="tripset" style="width:100%; height:100;border-bottom:1px solid #cccccc;display:none;">
  <table width="100%" border=0 cellpadding="3" cellspacing="0" bgcolor="#ffffff">
    <tr valign="bottom"> 
      <td colspan="4" align="center" bgcolor="#f0f0e0" class="tblunderline" style="border-left:1px solid #e3e3d3;">TRIP 
        DETAILS</td>
    </tr>
    <tr valign="middle"> 
      <td align="center" class="tblunderline">Start Date&nbsp;</td>
      <td align="center" class="tblunderline">End Date (read date)&nbsp;</td>
      <td align="right" class="tblunderline">Bill Period / Bill Year</td>
    </tr>
    <%
	sqlstr =  "select * from billyrperiod where (year(datestart) >= year(getdate()) or year(dateend)=year(getdate())) and bldgnum = '" & rst1("bldgnum") & "' and utility = " &  rst1("utilityid") 
	rs.Open sqlstr, getLocalConnect(rst1("bldgnum"))
	if not rs.eof then
		Do until rs.EOF 
	%>
    <form name=form1 method="post" action="">
      <tr valign="middle"> 
        <td align="right" class="tblunderline"><%=rs("datestart")%>&nbsp;</td>
        <td align="right" class="tblunderline"><%=rs("dateend")%>&nbsp;</td>
        <td align="right" class="tblunderline"><%=rs("billperiod")%>/<%=rs("billyear")%>&nbsp;</td>
      </tr>
    </form>
    <%  
    rs.movenext
    loop
end if
rs.close
%>
  </table>
</div>
  <%  
    rst1.movenext
    loop
end if
%>
<p>&nbsp;</p>
</body>
</html>
