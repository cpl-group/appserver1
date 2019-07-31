<%@Language="VBScript"%>
<%
if isempty(Session("name")) then
	Response.Redirect "http://www.genergyonline.com"
end if	
	
If Session("eri")  >  2 then
	Set cnn1 = Server.CreateObject("ADODB.Connection")
	Set rst1 = Server.CreateObject("ADODB.recordset")
	cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"
	type1=Request("type1")
	item=Request("item")
	dir=Request("dir")
	count=0
	'Response.Write type1&" "&item&" "&dir
	if (type1="") then
		if(item ="") then
		    strsql = "SELECT * FROM tblSurveylib"
		else
		    if(dir = "A") then
			    strsql = "SELECT * FROM tblSurveylib order by "&item&""
			else
				strsql = "SELECT * FROM tblSurveylib order by "&item&" desc"
			end if
		end if
	else
		if(item ="") then
			strsql = "SELECT * FROM tblSurveylib Where type='"&type1&"' "  
		else
			if(dir = "A") then
			    strsql = "SELECT * FROM tblSurveylib Where type='"&type1&"' order by "&item&" "
			else
				strsql = "SELECT * FROM tblSurveylib Where type='"&type1&"' order by "&item&" desc"
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
<script>
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
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post" action="">
<input type="hidden" name="type1" value="<%=type1%>">
<input type="hidden" name="dir" value="<%=dir%>">
<input type="hidden" name="item" value="<%=item%>">
</form>


<table border="1" width="116%">
  <% 
    	  
	Do While Not rst1.EOF
%>
  <tr>
    <form name="form2" method="post" action="libraryupdate.asp">
   
	
    <td align="center" width="14%"> 
      <input type="text" name="type1" value="<%=Trim(rst1("type"))%>" size="15">
	</td>
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
      <input type="Submit" name="Submit" value="Update">
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
		
    <td align="center" width="14%"> 
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
	  <input type="Submit" name="Submit" value="Add">
	</td>
	<input type="hidden" name="temp" value="<%=type1%>">
	<input type="hidden" name="dir" value="<%=dir%>">
	<input type="hidden" name="item" value="<%=item%>">
	</form>
</tr>
</table>

</body>
</html>
