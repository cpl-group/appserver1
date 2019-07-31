<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->

<%

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")

bldg=request.querystring ("bldg")
fid=request.querystring ("fid")

hasrecords = true
sqlstr= "select room, r.id, f.floor from room r INNER JOIN floor f ON f.id=r.floor where r.bldg='"&bldg&"'and r.floor="&fid
'response.write sqlstr
'response.end
rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1
if rst1.eof then
	rst1.close
	rst1.open "SELECT floor FROM floor f where bldg='"&bldg&"'and id="&fid, cnn1
	hasrecords = false
end if
fl=rst1("floor")
%>

<script>
try{top.applabel("Floor Management - View Rooms on <%=fl%> floor");}catch(exception){}
function findfixture(bldg,room,fl,rid,fid){
  var temp= "fixsearch.asp?bldg=" +bldg+"&room="+room+"&floor="+fl+"&fid="+fid+"&rid="+rid
	document.location=temp
}
function newfix(bldg,fl,fid){
//alert(utility)
  var temp = "newroom.asp?bldg=" +bldg+"&floor="+fl+"&fid="+fid
  //alert(temp)
	document.location=temp
}

</script>
<link rel="Stylesheet" href="/genergy2/styles.css">

<body bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="3" cellspacing="0" style="border:solid 1px #ffffff;">
  <tr> 
    <td bgcolor="#CCCCCC"> 
      <div align="left"><font face="Arial, Helvetica, sans-serif" color="#000000" size="2"><span class="standard"><b><a href="floorsearch.asp?bldg=<%=bldg%>" class="breadcrumb">Floor</a> : 
        <%=fl%> </b></span></font></div>
    </td>
  	<td align="right" bgcolor="#CCCCCC"><input type="button" name="newf2" value="Add New Room" onClick="newfix('<%=bldg%>','<%=fl%>','<%=fid%>')" class="standard">
      <input type="button" name="newf" value="Edit Floor" onclick="document.location='newfloor.asp?bldg=<%=bldg%>&fid=<%=fid%>'" class="standard"></td>
  </tr>
  </table>
  
      
<table width="100%" cellpadding="3" cellspacing="1" border="0">
  <tr> 
    <td><span class="standard"><input type="hidden" name="bldg" value="<%=bldg%>">
      <strong>Select a room:</strong></span></td>
  </tr>
  <% While not(rst1.EOF) and hasrecords%>
  <tr valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="javascript:findfixture('<%=bldg%>','<%=rst1("room")%>','<%=fl%>','<%=rst1("id")%>','<%=fid%>')"> 
    <td width="20%"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><%=rst1("room")%></span></font></td>
  </tr>
  <%
		x=x+1
		rst1.movenext
		Wend
		%>
	<tr bgcolor="#eeeeee">
	  
    <td bgcolor="#eeeeee">&nbsp;</td>
	</tr>
</table>
   
<%

rst1.close
%>


