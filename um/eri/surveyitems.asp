<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if isempty(Session("name")) then%>
<script>
top.location="../index.asp"
</script>
<%'Response.Redirect "http://www.genergyonline.com"
end if

dim xscroll, yscroll, bldg
bldg = request("bldg")
xscroll = 0
yscroll = 0
if trim(request("xscroll"))<>"" then if isnumeric(request("xscroll")) then xscroll = cint(request("xscroll"))
if trim(request("yscroll"))<>"" then if isnumeric(request("yscroll")) then yscroll = cint(request("yscroll"))

if Session("eri") > 2 then 	
%>
<html>
<head>
<script language="JavaScript">

function openpopup(type, count){
	//var str2=document.doublecombo.qty.value
	var temp="filter.asp?type1="+type+"&count="+count
	//alert(temp);
    window.open(temp,"", "scrollbars=yes, width=300, height=338" );

}
	
function Desc_alert(){
	alert("Please enter a Description")
}
function delete1(theform,xscroll, yscroll){//key,surveyid,xscroll, yscroll){
	if(confirm("Delete survey item?")){
		theform.action = "deleteitem.asp"
		theform.submit()
	}
	}
function setValue(nm)
{
}
</script>





<title>Survey Item Details</title>

<meta http-equiv="Content-type1" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">   
<style type="text/css">
.topline { border-top:1px solid #ffffff; }
.tblunderline td { border-bottom:1px solid #ffffff; }
</style>
</head>

<body bgcolor="eeeeee" text="#000000" topmargin="0" onload="scroll(<%=xscroll%>,<%=yscroll%>);">
<table border=0 cellspacing="1" cellpadding="3" bgcolor="#eeeeee" width="100%" class="tblunderline">
  <tr bgcolor="#dddddd"> 
    <td>Qty</td>
    <td>Type</td>
    <td>Description</td>
    <td>Amps</td>
    <td>Volt</td>
    <td>PH</td>
    <td>PF</td>
    <td>MF</td>
    <td>Watt</td>
    <td>TotKw</td>
    <td>Adj %</td>
    <td align="right">Adj Kw</td>
    <td>HOn</td>
    <td>HOff</td>
    <td>KwhOn</td>
    <td>KwhOff</td>
    <td>Int</td>
    <td>Base</td>
    <td>&nbsp; </td>
  </tr>
  <%
Set cnn1 = Server.CreateObject("ADODB.Connection")

cnn1.open getConnect(0,0,"Engineering")

sql = "SELECT * FROM tblSurveyItem WHERE(surveyid='"& request("survey_id") & "') order by id"


Set rst1 = Server.CreateObject("ADODB.Recordset")

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly


count = 0 
If not rst1.EOF then 

%>
  <% 
    
    Do until rst1.EOF 
	type1=rst1("type")
	name="forms"&"["&count&"]"
	%>
  <tr bgcolor="#eeeeee"> 
    <form name=form1 method="post" action="itemsupdate.asp" onsubmit="this.yscroll.value=window.document.body.scrollTop;">
      <input type="hidden" name="id" value="<%=Request("id")%>">
      <input type="hidden" name="orderno" value="<%=Request("orderno")%>">
      <input type="hidden" name="tenant_no" value="<%=Request("tenant_no")%>">
      <input type="hidden" name="location" value="<%=Request("location")%>">
      <input type="hidden" name="key" value="<%=rst1("id")%>">
      <td><input type="text" name="qty" value="<%=rst1("qty") %>" size="3" onChange="setValue(<%=count%>)" maxlength="3" ></td>
      <td> 
        <select name="type1" size="1" >
          <option value="<%=rst1("type")%>" selected><%=rst1("type")%></option>
          <option value="Lighting">Lighting</option>
          <option value="Equipment">Equipment</option>
          <option value="HVAC">HVAC</option>
        </select>
      </td>
      <td><input type="text" name="description2" size="20" value="<%=rst1("description")%>"></td>
      <td align="right"><input type="text" name="amps" size="6" value="<%=formatnumber(rst1("amps"),2)%>" maxlength="6"></td>
      <td align="right"><input type="text" name="volt" size="3" value="<%=rst1("volt")%>"  maxlength="3"></td>
      <td align="right"><input type="text" name="ph" size="2" value="<%=rst1("ph")%>" maxlength="2"></td>
      <td align="right"><input type="text" name="pf" size="4" value="<%=formatnumber(rst1("pf"),2)%>" maxlength="4"></td>
      <td align="right"><input type="text" name="mf" size="3" value="<%=rst1("monthfactor")%>" maxlength="3"></td>
      <td align="right"><input type="text" name="watt" size="6" value="<%=rst1("watt")%>" maxlength="10" ></td>
      <td align="right"><%=formatnumber(rst1("totalkw"),2)%></td>
      <td align="right"><input type="text" name="adj" value="<%=formatnumber(rst1("adjfactor"),2)%>" size="4" maxlength="4" ></td>
      <td align="right"><%=formatnumber(rst1("adjkw"),2)%></td>
      <td align="right"><input type="text" name="hon" value="<%=rst1("houron")%>" size="4" maxlength="4" ></td>
      <td align="right"> <input type="text" name="hoff" value="<%=rst1("houroff")%>" size="4" maxlength="4" ></td>
      <td align="right"><%=formatnumber(rst1("kwhon"),2)%></td>
      <td align="right"><%=formatnumber(rst1("kwhoff"),2)%></td>
      <td>  
        <% 
			if(rst1("intense")=true) then
			%>
        <input type="checkbox" name="ie" size="6" checked>
          <%
			else
			%>
          <input type="checkbox" name="ie" size="6" >
          <%
			end if
			
			%>
        </td>
        <td> 
           
          <%
			if(rst1("base")=true) then
			%>
          <input type="checkbox" name="bho" size="6" checked>
          <%
			else
			%>
          <input type="checkbox" name="bho" size="6">
          <%
			end if
			%>
          
        </td>
        
      <%if not(isBuildingOff(bldg)) then%><td><input type="Submit" name="Submit" value="Update"><input type="button" value="Delete" onclick="this.form.yscroll.value=document.body.scrollTop; delete1(this.form)"></td><%end if%>
		<input type="hidden" name="yscroll" value="">
    </form>
  </tr>
  <%
    count = count +1
    rst1.movenext
    Loop
	rst1.close
    %>
  <%
end if
	
%>
<%if not(isBuildingOff(bldg)) then%>
  <tr bgcolor="#dddddd"> 
    <form name="form2" method="post" action="itemsadd.asp" onsubmit="this.yscroll.value=window.document.body.scrollTop;">
      <input type="hidden" name="id" value="<%=Request("id")%>">
      <input type="hidden" name="orderno" value="<%=Request("orderno")%>">
      <input type="hidden" name="tenant_no" value="<%=Request("tenant_no")%>">
      <input type="hidden" name="location" value="<%=Request("location")%>">
      <input type="hidden" name="key" value="<%=id%>">
      <td><a tabindex=0><input type="text" name="qty" value="0" size="3" onChange="setValue(<%=count%>)" maxlength="3"></a></td>
      <td> 
        <select name="type1" size="1" >
          <option value="Lighting">Lighting</option>
          <option value="Equipment">Equipment</option>
          <option value="HVAC">HVAC</option>
          <option value="" selected>========</option>
        </select>
      </td>
      <td><input type="text" name="description" size="20" value="Desc" ></td>
      <td align="right"><input type="text" name="amps" size="6" value="0" maxlength="6"></td>
      <td align="right"><input type="text" name="volt" size="3" value="115" maxlength="3"></td>
      <td align="right"><input type="text" name="ph" size="2" value="1" maxlength="2"></td>
      <td align="right"><input type="text" name="pf" size="4" value="0.85" maxlength="4"></td>
      <td align="right"><input type="text" name="mf" size="3" value="12" maxlength="3"></td>
      <td align="right"><input type="text" name="watt" size="6" value="0" maxlength="10" ></td>
      <td align="right">&nbsp;</td>
      <td align="right"><input type="text" name="adj" value="1" size="4" maxlength="4" ></td>
      <td align="right">&nbsp;</td>
      <td align="right"><input type="text" name="hon" size="4" value="0" maxlength="4" ></td>
      <td align="right"><input type="text" name="hoff" size="4" value="0" maxlength="4" ></td>
      <td align="right">&nbsp;</td>
      <td align="right">&nbsp;</td>
      <td><input type="checkbox" name="ie2" size="6"></td>
      <td><input type="checkbox" name="bho" size="6"></td>
      <td><input type="Submit" name="Submit2" value="Add" style="padding-left:6px;padding-right:6px;"></td>
		<input type="hidden" name="yscroll" value="">
    </form>
  </tr>
  <%end if%>
  <%

end if
'cnn1.close
set cnn1=nothing
%>
</table>
</body>
</html>
