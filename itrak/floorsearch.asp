<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")

bldg=request.querystring ("bldg")
		sqlstr= "select * from floor where bldg='"&bldg&"' order by floor"
rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic
'response.write sqlstr
'response.end

rst1.Open sqlstr, cnn1, 0, 1, 1

%>

<html>
<head>
<script>

function findroom(bldg,floor1, fid){
var temp = "roomsearch.asp?bldg=" +bldg+"&floor="+floor1+"&fid="+fid
	document.location=temp
}

function newfix(bldg,fid){
//alert(utility)
  var temp = "newfloor.asp?bldg=" +bldg+"&fid="+fid
  //alert(temp)
	document.location=temp
}
try{top.applabel("Floor Management");}catch(exception){}
</script>
<link rel="Stylesheet" href="/genergy2/styles.css">
</head>

<body bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="3" cellspacing="0" style="border:solid 1px #ffffff;">
  <tr> 
    <td width="26%" align="left" bgcolor="#FFFFFF" nowrap><span class=standardheader><font color="#000000">Select 
      Building :</font></span> 
      <select name="select">
        <option value="#">123 Main Street</option>
      </select></td>
    <td width="74%" align="right" bgcolor="#FFFFFF"><input type="button" name="newf" value="Add New Floor" onClick="newfix('<%=bldg%>','<%=fid%>')" class="standard"></td>
  </tr>
</table>
  
      
<table width="100%" cellpadding="3" cellspacing="1" border="0">
  <tr bgcolor="#CCCCCC"> 
    <td bgcolor="#cccccc"><font face="Arial, Helvetica, sans-serif" color="#000000" size="2"><span class="standard"> 
      <input type="hidden" name="bldg" value="<%=bldg%>">
      Select a floor:</span></font></td>
  </tr>
  <% While not rst1.EOF %>
  <form name="form1" method="post" action="">
    <tr valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="javascript:findroom('<%=request.querystring("bldg")%>','<%=rst1("floor")%>','<%=rst1("id")%>')"> 
      <td width="8%"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><%=rst1("floor")%></span></font></td>
    </tr>
  </form>
  <%
		x=x+1
		rst1.movenext
		Wend
		%>
  <tr bgcolor="#eeeeee">
    <td>&nbsp;</td>
  </tr>
</table>
  

<%

rst1.close
 
%>
</html>


