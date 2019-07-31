<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%



	
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")

bldg=request("bldg")
rid=request("rid")
fid=request("fid")


hasrecords = true
sqlstr= "select  r.room as roomname, fl.floor as floorname,f.*,  ft.fix_catalog+' ' as fixture,DATEADD(week,(est_lamp_life/est_hr_wk) ,dlc)as estd , datediff(week,getdate(), (DATEADD(week,(est_lamp_life/est_hr_wk) , dlc))) as weeksr from fixture_types ft join fixtures f on ft.id=f.typeid  INNER JOIN room r ON r.id=f.room INNER JOIN floor fl ON r.floor=fl.id  where f.bldgnum='"&bldg&"'and f.room='"&rid&"' order by weeksr"
'response.write sqlstr
'response.end
rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.eof then
	rst1.close
	sqlstr="SELECT f.floor as floorname, r.room as roomname FROM floor f INNER JOIN room r ON r.floor=f.id where r.id="&rid
	rst1.open sqlstr, cnn1
	room=rst1("roomname")
	fl=rst1("floorname")
	hasrecords = false
else 
	room=rst1("roomname")
	fl=rst1("floorname")
end if
%>
<script>
try{top.applabel("Floor Management - View Fixtures in <%=room%> on <%=fl%> floor");}catch(exception){}
function findfixture(id){
  var temp = "fixtureinfo.asp?id=" +id  + "&bldg=<%=bldg%>&room=<%=room%>&floor=<%=fl%>&fid=<%=fid%>&rid=<%=rid%>";
  //alert(temp)
	document.location=temp
}
function editroom(bldg,room,fl,rid,fid){
//alert(utility)
  var temp = "newroom.asp?bldg=" +bldg+"&rid="+rid+"&floor="+fl+"&fid="+fid
  //alert(temp)
	document.location=temp
}

function newfix(bldg,room,fl,fid,rid){
//alert(utility)
  var temp = "newfixture.asp?bldg=" +bldg+"&room="+room+"&fl="+fl+"&fid="+fid+"&rid="+rid
  //alert(temp)
	document.location=temp
}

</script>

<link rel="Stylesheet" href="/genergy2/styles.css">

<body bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="3" cellspacing="0" style="border:solid 1px #ffffff;">
  <tr> 
    <td bgcolor="#cccccc" width="37%"> 
      <div align="left"><font face="Arial, Helvetica, sans-serif" color="#000000" size="2"><span class="standard"><b><a href="floorsearch.asp?bldg=<%=bldg%>" class="breadcrumb">Floor</a>: <%=fl%> | <a href="roomsearch.asp?bldg=<%=bldg%>&floor=<%=fl%>&fid=<%=fid%>" class="breadcrumb">Room</a>: <%=room%></b></span></font></div>
    </td>
    <td  width="12%" align="right" nowrap bgcolor="#cccccc"> 
      <font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><b>
        <input type="button" name="newf2" value="Add Fixture" onClick="newfix('<%=bldg%>','<%=room%>','<%=fl%>',<%=fid%>,'<%=rid%>')" class="standard">
        <input type="button" name="newf" value="Edit Room" onclick="editroom('<%=bldg%>','<%=room%>','<%=fl%>','<%=rid%>', '<%=fid%>')" class="standard">
        </b></span></font>
    </td>
  </tr>
</table>
<table width="100%" cellpadding="3" cellspacing="1" border="0">
  <tr> 
    <td colspan="4" ><font face="Arial, Helvetica, sans-serif" color="#000000" size="2"><span class="standard"><strong>Select 
      a fixture:</strong></span></font></td>
  </tr>
  <tr>
  	<td bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><b>Fixture Catalog Number</b></span></font></td>
  	<td bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><b>Est. Replacement Date</b></span></font></td>
  	<td bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><b>Weeks Remaining</b></span></font></td>
  	<td bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><b>Fixture Count</b></span></font></td>
  </tr>
  <% While not rst1.EOF and hasrecords%>
  <tr valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="javascript:findfixture('<%=rst1("id")%>')"> 
    <td><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><%=rst1("fixture")%></span></font></td>
	<td><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><%=rst1("estd")%></span></font></td>
	<td><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><% if not rst1("weeksr") > 0 then%> 0 (<%=abs(rst1("weeksr"))%> weeks overdue)<% else %><%=rst1("weeksr")%><% end if %></span></font></td>
	<td><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><%=rst1("fixtureqty")%></span></font></td>
  </tr>
  <%
	rst1.movenext
	Wend
		%>
	<tr>
	  
    <td colspan="4" bgcolor="#eeeeee">&nbsp;</td>
	</tr>
</table>
   
<%

rst1.close
set cnn1=nothing
%>


