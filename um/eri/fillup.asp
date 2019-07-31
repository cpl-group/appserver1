<%@Language="VBScript"%>
<!-- #include file="./adovbs.inc" -->
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'			Response.Redirect "http://www.genergyonline.com"
		end if
		
		if Session("eri") > 2 then 	
		type1=Request.QueryString("type")
		
		description=Request.QueryString("description")
		count=Request("count")
		'Response.Write count
		'Response.Write type1&"@"&description
		
%>
<html>
<head>
<script language="JavaScript">
//function openpopup(){
//configure "Open Logout Window
//window.open("../logout.asp","","width=300,height=338")
//parent.document.location.href="../index.asp";
//}
function openpopup(type, id){
	//var str2=document.doublecombo.qty.value
	var temp="filter.asp?type1="+type;
	//alert(temp);
    window.open(temp,"", "scrollbars=yes", "width=300", "height=338" );

}

function reload2(str3){
    var str=document.doublecombo.type1.value
	var str2=document.doublecombo.qty.value
	var temp="surveyitems.asp?type1="+str+"&qty="+str2+"&description="+str3;
	//alert(temp);
	
//    window.location=temp;
	
}


</script>
<title>Untitled Document</title>

<meta http-equiv="Content-type1" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table>  
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")


cnn1.open getConnect(0,0,"Engineering")



sql = "SELECT * FROM tblSurveyLib WHERE(type='"& type1& "' and description='"& description &"')"

Set rst1 = Server.CreateObject("ADODB.Recordset")
rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly

If not rst1.EOF then 

%>
    <% 
    'count = 0 
    Do until rst1.EOF 
	name="forms"&"["&count&"]"
	if not (isnumeric(rst1("watt"))) then
		watt=0
	else
	    watt=rst1("watt")
	end if
	
	%>
	<tr>
	    <td width=4%>
		    <input type="text" name="type" value="<%Response.Write(type1)%>" size="8" >
			
	    </td>
	    <td>
		    <input type="text" name="description" size="" value="<%Response.Write(description)%>">
			
			<script>
			opener.document.forms[<%=count%>].description.value="<%Response.Write(description)%>"
			opener.document.forms[<%=count%>].amps.value="<%=rst1("amps")%>"
			opener.document.forms[<%=count%>].volt.value="<%=rst1("volt")%>"
			opener.document.forms[<%=count%>].ph.value="<%=rst1("ph")%>"
			opener.document.forms[<%=count%>].pf.value="<%=rst1("pf")%>"
			opener.document.forms[<%=count%>].mf.value="<%=rst1("monthfactor")%>"
			opener.document.forms[<%=count%>].watt.value="<%=watt%>"
			opener.document.forms[<%=count%>].totkw.value=opener.document.forms[<%=count%>].qty.value*<%=watt%>
			opener.document.forms[<%=count%>].adj.value="<%=rst1("adjfactor")%>"
			opener.document.forms[<%=count%>].adjkw.value=opener.document.forms[<%=count%>].qty.value*<%=(rst1("adjfactor")*watt)%>
			window.close()
			</script>	
        
    	<td width="2%">
			<input type="text" name="pf" size="6" value="<%=rst1("pf")%>">
		</td>
    	<td width="3%">
			<input type="text" name="mf" size="6" value="<%=rst1("monthfactor")%>"> 
		</td>
    	<td width="4%">
			<input type="text" name="watt" size="6" value="<%=rst1("watt")%>">	
		</td>
      	<td width="3%">
			<input type="text" name="adj" value="<%=rst1("adjfactor")%>" size="6">
      	
		</td>
	</tr>
    <%
    rst1.movenext
    Loop
	rst1.close
    %>
<%
end if
end if
cnn1.close
%>
</table>
<script>
//window.close()
</script>
</body>
</html>
