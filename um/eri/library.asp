<%@Language="VBScript"%>
<!-- #include VIRTUAL="/genergy2/secure.inc" -->
<%
if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'	Response.Redirect "http://www.genergyonline.com"
end if	
	
If Session("eri")  >  2 then
	Set cnn1 = Server.CreateObject("ADODB.Connection")
	Set rst1 = Server.CreateObject("ADODB.recordset")
	cnn1.open getConnect(0,0,"Engineering")
	type1=Request("type1")
	item=Request("item")
	dir=Request("dir")
	count=0
	'Response.Write type1&" "&item&" "&dir
	if (type1="") then
		if(item ="") then
		    strsql = "SELECT  top 50 * FROM tblSurveylib where type='lighting'"
			
		else
		    if(dir = "A") then
			    strsql = "SELECT top 50 * FROM tblSurveylib where type='lighting' order by "&item&""
			else
			
				strsql = "SELECT top 50* FROM tblSurveylib where type='lighting' order by "&item&" desc"
			end if
		end if
	else
		if(item ="") then
		    
			strsql = "SELECT top 50 * FROM tblSurveylib where type='lighting'"   '&type1&"' "  
		else
			if(dir = "A") then
			    strsql = "SELECT top 50 * FROM tblSurveylib where type='lighting'"   '&type1&"' order by "&item&" "
			else
			
				strsql = "SELECT top 50 * FROM tblSurveylib where type='lighting'"    '&type1&"' order by "&item&" desc"
			end if
		end if
	
	end if
	'Response.Write strsql
	rst1.Open strsql, cnn1, adOpenStatic
	
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" type="text/javascript">
function filter(type1){
    
	var dir=document.form1.dir.value
	var item=document.form1.item.value
	window.location="library.asp?type1="+type1+"&dir="+dir+"&item="+item
}

function sortInOrder(dir, item){
	var type1=document.form1.type1.value
	if(dir == "A-Z"){
		dir="A"
	}else{
		dir="Z"
	}
	//document.form1.item.value=item
	//document.form1.dir.value=dir
	window.location="library.asp?type1="+type1+"&dir="+dir+"&item="+item	
}


</script>
<script language="JavaScript" type="text/javascript">
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
</script>
</head>

<body bgcolor="#eeeeee" text="#000000">
<form name="form1" method="post" action="">
<input type="hidden" name="type1" value="<%=type1%>">
<input type="hidden" name="dir" value="<%=dir%>">
<input type="hidden" name="item" value="<%=item%>">
</form>


<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr valign="bottom" bgcolor="#dddddd" style="font-weight:bold;">
  <td>Type</td>
  <td>Description</td>
  <td>Amps</td>
  <td>Volt</td>
  <td>PH</td>
  <td>PF</td>
  <td>Watt</td>
  <td>Month Factor</td>
  <td>Adj. Factor</td>
  <td></td>
</tr>
  <% 
    	  
	Do While Not rst1.EOF
%>
  <tr>
    <form name="form2" method="post" action="libraryupdate.asp">
   
	
    <td width="14%"><input type="text" name="type1" value="<%=Trim(rst1("type"))%>" size="15"></td>
    <td width="18%"> 
	
      <input type="text" name="description" value="<%=Trim(rst1("description"))%>" size="20"></td>
    <td width="10%"> 
      <input type="text" name="amps" value="<%=rst1("amps")%>" size="10"></td>
    <td width="8%"> 
      <input type="text" name="volt" value="<%=rst1("volt")%>" size="5"></td>
    <td width="8%"> 
      <input type="text" name="ph" value="<%=rst1("ph")%>" size="5"></td>
    <td width="8%"> 
      <input type="text" name="pf" value="<%=rst1("pf")%>" size="5"></td>
    <td width="10%"> 
      <input type="text" name="watt" value="<%=rst1("watt")%>" size="10"></td>
    <td width="7%"> 
      <input type="text" name="mf" value="<%=rst1("monthfactor")%>" size="5"></td>
    <td width="10%"> 
      <input type="text" name="adj" value="<%=rst1("adjfactor")%>" size="10">
    </td>
	<td>
      <input type="Submit" name="Submit" value="Update" style="border:1px outset #ddffdd;background-color:ccf3cc;">
	  <input type="hidden" name="temp" value="<%=type1%>">
	  <input type="hidden" name="dir" value="<%=dir%>">
	  <input type="hidden" name="item" value="<%=item%>">
	</td>
      
	</form>
</tr>

<%
	rst1.MoveNext
	count=count+1  
	Loop
end if

rst1.Close
Set rst1 = Nothing
cnn1.Close
Set cnn1 = Nothing
%>
<tr>
	<form name=form2 method="post" action="libraryadd.asp">
		
    <td width="14%"> 
      <input type="text" name="type1" value="" size="15">
    </td>
   
    <td width="18%"> 
      <input type="text" name="description" value="" size="20"></td>
    <td width="10%"> 
      <input type="text" name="amps" value="" size="10"></td>
    <td width="8%"> 
      <input type="text" name="volt" value="" size="5"></td>
    <td width="8%"> 
      <input type="text" name="ph" value="" size="5"></td>
    <td width="8%"> 
      <input type="text" name="pf" value="" size="5"></td>
    <td width="10%"> 
      <input type="text" name="watt" value="" size="10"></td>
    <td width="7%"> 
      <input type="text" name="mf" value="" size="5"></td>
    <td width="10%"> 
      <input type="text" name="adj" value="" size="10">
    </td>
	<td>
	  <input type="Submit" name="Submit" value="Add" style="border:1px outset #ddffdd;background-color:ccf3cc;">
	</td>
	<input type="hidden" name="temp" value="<%=type1%>">
	<input type="hidden" name="dir" value="<%=dir%>">
	<input type="hidden" name="item" value="<%=item%>">
	</form>
</tr>
</table>

</body>
</html>
