<%@Language="VBScript"%>
<!-- #include file="./adovbs.inc" -->
<%
		if isempty(Session("name")) then
			Response.Redirect "http://www.genergyonline.com"
		end if
		
		if Session("eri") > 2 then 	
		type1=Request.QueryString("type1")
		qty=Request("qty")
		des=Request.QueryString("description")
		tenant_no=Request("tenant_no")
		location=Request("location")
		orderno=Request("orderno")
		Response.Write tenant_no&" "&location&" "&orderno
%>
<html>
<head>
<script language="JavaScript">
function reload(str){
	var str2=document.doublecombo.qty.value
	var temp="http://appserver1.genergy.com/eri/surveyitems.asp?type1="+str+"&qty="+str2;
	//alert(temp);
    window.location=temp;
	
}

function reload2(str3){
    var str=document.doublecombo.type1.value
	var str2=document.doublecombo.qty.value
	var temp="http://appserver1.genergy.com/eri/surveyitems.asp?type1="+str+"&qty="+str2+"&description="+str3;
	//alert(temp);
	
    window.location=temp;
	
}

function setKwon(val){
    var kwon=document.doublecombo.adjkw.value*val
	//alert(kwon)
	document.doublecombo.kwon.value=kwon
}
function setKwoff(val){
    var kwoff=document.doublecombo.adjkw.value*val
	//alert(kwon)
	document.doublecombo.kwoff.value=kwoff
}
</script>
<title>Untitled Document</title>

<meta http-equiv="Content-type1" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="doublecombo">
  <table width="100%" border="1" cellspacing="2" cellpadding="0">
    <tr>
    <td width="4%" height="39">Qty</td>
    <td width="10%" height="39">Type</td>
    <td width="17%" height="39">Descriptionn</td>
    <td width="5%" height="39">Amps</td>
    <td width="3%" height="39">Volt</td>
    <td width="3%" height="39">PH</td>
    <td width="2%" height="39">PF</td>
    <td width="3%" height="39">MF</td>
    <td width="4%" height="39">Watt</td>
    <td width="6%" height="39">TotKw</td>
    <td width="3%" height="39">Adj %</td>
    <td width="3%" height="39">Adj Kw</td>
    <td width="4%" height="39">HOn</td>
    <td width="4%" height="39">HOff</td>
    <td width="6%" height="39">KwhOn</td>
    <td width="6%" height="39">KwhOff</td>
    <td width="8%" height="39"> 
      <div align="center">Intensive Equipment</div>
    </td>
    <td width="9%" height="39"> 
      <div align="center">Base Hrs Operation</div>
    </td>
  </tr>
  <tr>
    <td width="4%">
<%
if not isempty(Request.QueryString("qty")) then
'Response.Write Request.QueryString("qty")
%>  
  <input type="text" name="qty" value="<%=Response.Write (qty) %>" maxlength=" " size="8">
<%
else 
%>  
  <input type="text" name="qty" maxlength=" " size="8">     </td>
<%
end if
' end of line83
%>
    <td width="10%">

<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"

sql = "SELECT * FROM tblLoadtype"
Set rst1 = Server.CreateObject("ADODB.Recordset")

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
If not rst1.EOF then 

%>
      <select name="type1" size="1" onChange=reload(this.value)>
    <% 
    count = 0 
    Do until rst1.EOF 
    if(type1=rst1("description")) then
    %>
      <option value="<%=rst1("description")%>" selected><%=rst1("description")%></option>
    <%
    else
    %>
	  <option value="<%=rst1("description")%>" ><%=rst1("description")%></option>
    <%
    end if
    ' end of line 115
    count = count +1
    rst1.movenext
    Loop
    rst1.close
end if 
' end of line 106
%>
      </select>
      </td>
      <td width="17%"> 
<%
if NOT isempty(Request.QueryString("type1")) then
'Response.Write Request.Form("type1")
	Set rst2 = Server.CreateObject("ADODB.Recordset")
  
	sql2="SELECT description FROM tblSurveyLib WHERE (type='"& type1 &"')"
  
    rst2.Open sql2, cnn1, adOpenStatic, adLockReadOnly
  
    If not rst2.EOF then 
%>

      <select name="description" onChange=reload2(this.value) >
		<%
		if isempty(Request("description")) then
				
        %>
	    	<%
	    	Do until rst2.EOF 
		    %>
	      <option value="<%=rst2("description")%>" ><%=rst2("description")%></option>
	       <%
	    	rst2.movenext
	    	Loop
		else
		    Set rst4 = Server.CreateObject("ADODB.Recordset")
            sql4="SELECT description FROM tblSurveyLib WHERE (description!='"& des &"')"
            rst4.Open sql4, cnn1, adOpenStatic, adLockReadOnly
		    %>
		    <option value="<%=des%>" selected ><%=des%></option> 
            <%	
			If not rst4.EOF then
			DO until rst4.EOF
	        %>				
		    <option value="<%=rst4("description")%>" ><%=rst4("description")%></option>
  		<%
    		'end of line 166
			rst4.movenext
			Loop
			end if
		end if
		'end of line 152
	end if
	'end of 144
	rst2.close
end if	
'end of 136
%>
       </select>

<%
  if NOT isempty(Request("description")) then

  Set rst3 = Server.CreateObject("ADODB.Recordset")
  
  sql3="SELECT * FROM tblSurveyLib WHERE (description='"& des &"')"
  
  rst3.Open sql3, cnn1, adOpenStatic, adLockReadOnly
  
  If not rst3.EOF then 

%>
    </td>
      <td width="5%">
        <input type="text" name="amps" size="6" value="<%=rst3("amps")%>">
      </td>
      <td width="3%">
        <input type="text" name="volt" size="6" value="<%=rst3("volt")%>">
      </td>
      <td width="3%">
        <input type="text" name="ph" size="6" value="<%=rst3("ph")%>">
      </td>
    <td width="2%"><input type="text" name="pf" size="6" value="<%=rst3("pf")%>"></td>
    <td width="3%"><input type="text" name="mf" size="6" value="<%=rst3("monthfactor")%>"></td>
    <td width="4%"><input type="text" name="watt" size="6" value="<%=rst3("watt")%>"></td>
      <td width="6%">
        <input type="text" name="totkw" size="6" value="<%=(rst3("watt")* qty)%>">
      </td>
    <td width="3%"><input type="text" name="adj" value="<%=rst3("adjfactor")%>" size="6"></td>
      <td width="3%">
        <input type="text" name="adjkw" value="<%=(rst3("adjfactor")*(rst3("watt")* qty))%>"size="8">
      </td>
    <td width="4%"><input type="text" name="hon" size="6" onChange=setKwon(this.value)></td>
    <td width="4%"><input type="text" name="hoff" size="6" onChange=setKwoff(this.value)></td>
    <td width="6%"><input type="text" name="kwon" size="6"></td>  
	  <td width="6%">
        <input type="text" name="kwoff" size="6">
      </td>
	<div align="center">
    <td width="8%"><input type="checkbox" name="ie" size="6"></td>
	<td width="9%"><input type="checkbox" name="bho" size="6"></td>
	</div>
  </tr>
<%
end if
%>

<%
end if
%>
</table>
</form>
<p> 
  <input type="button" name="test" value="Go!"
onClick="go()">
 
</p>
<%
end if
%>
</body>
</html>
