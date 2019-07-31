<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<html>
<head>
<%

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")
bldg=request.querystring ("bldg")
cid=request.querystring ("cid")
sqlstr= "select sum(f.fixtureqty) as qty, fixture_types.id,fixture_types.description,fix_catalog+' '+lamp_catalog as fc from fixture_types left join (SELECT * FROM fixtures WHERE bldgnum='"&bldg&"') as f on fixture_types.id=f.typeid where client='"&cid&"'  group by fixture_types.id,fixture_types.description,fixture_types.fix_catalog,fixture_types.lamp_catalog"
'response.write sqlstr
'response.end

rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1

%>

<script>
try{top.applabel("Lighting Catalogue");}catch(exception){}
function findfixture(id,bldg){
  var temp= "fixtureview.asp?id=" + id + "&bldg=" + bldg
	location=temp
}
function newfix(cid){
//alert(utility)
  var temp = "fixtype.asp?cid="+cid
  //alert(temp)
	location=temp
}

</script>
<link rel="Stylesheet" href="/genergy2/styles.css">
</head>
<body bgcolor="#FFFFFF">
<table width="100%" cellpadding="3" cellspacing="0" border="0" style="border:1px solid #ffffff">
  <tr> 
    
    <td bgcolor="#FFFFFF" width="68%"> 
      <div align="left"><font face="Arial, Helvetica, sans-serif" color="#ffffff" size="2"><span class="standard"></span></font></div>
    </td>
	<td width="11%" bgcolor="#FFFFFF"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><b>
	  <input type="hidden" name="cid" value="<%=cid%>">
        <input type="button" name="newf" value="Add New Fixture" onclick="newfix(cid.value)" class="standard">
        </b></span></font></div></td>
  </tr>
  </table>
  
      
<table width="100%"  cellpadding="3" cellspacing="1" border="0">
  <tr valign="bottom" bgcolor="#cccccc"> 
    <td bgcolor="#cccccc" height="19" width="51%"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"> 
      <b>Fixture Catalog / Lamp Catalog Number</b></span></font></td>
	   
	   
    <td bgcolor="#cccccc" height="19" width="28%"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"> 
      <b>Description</b></span></font></td>
	  
    <td bgcolor="#cccccc" height="19" width="21%"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"> 
      <b>Fixture Total</b></span></font></td>
  </tr>
  <% While not rst1.EOF %>
  <tr valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="javascript:findfixture('<%=rst1("id")%>','<%=bldg%>')"> 
    <td width="51%"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><%=rst1("fc")%></span></font></td>

	<td width="28%"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><%=rst1("description")%></span></font></td>
	<td width="21%"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><%=rst1("qty")%></span></font></td>
  </tr>
  <%
		rst1.movenext
		Wend
		%>
</table>
   
<%

rst1.close
set cnn1=nothing
%>

</html>

