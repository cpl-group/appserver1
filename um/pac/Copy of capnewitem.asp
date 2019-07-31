<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
		if isempty(Session("name")) then
%>
<script>
//top.location="../index.asp"
//window.close()
</script>
<%
			'			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("ts") < 4 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				Response.Redirect "../main.asp"
			end if	
		end if	
user=Session("name")
uid="genergy\"&Session("login")
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
Set rst3 = Server.CreateObject("ADODB.Recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=Capacity_db;"
bldgnum=request("bldgnum")
item=request("item")
riser=request("riser")
floor=request("floor")
msg=request("msg")
if item="riser" then
	item1="fl_name"
	item2="riser_name"
	title="Floor"
	title2="Riser"
	val=riser
	sql="SELECT distinct "&item1&" as choice from tblfloor where bldgnum='"& bldgnum &"'"
else
    item1="riser_name"
	item2="fl_name"
	title="Riser"
	title2="Floor"
	val=floor
	sql="SELECT distinct "&item1&" as choice from tblriser where bldgnum='"& bldgnum &"' order by "&item1&""
end if
'if val="" then
	
'else
	sql3 = "SELECT distinct "&item1&" as choice from tblassociation where bldgnum='"& bldgnum &"' and "&item2&" ='"& val &"' "
'end if
'response.write sql
'response.end
rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
rst3.Open sql3, cnn1, adOpenStatic, adLockReadOnly
flag=0
if not rst1.eof then
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function openpopup(){
//configure "Open Logout Window

parent.document.location.href="../index.asp";
}
function loadpopup(){
openpopup()
}
function setUp(){
	document.forms[0].list.multiple=true
}
function redirect(bldgnum, i, f, r){
	//bldgnum=document.forms[0].bldgnum.value
	//item=document.forms[0].item.value
	//floor=document.forms[0].floor.value
	//riser=document.forms[0].riser.value
	
	if(i=="floor"){
		opener.parent.riser.location="capriser.asp?bldgnum="+bldgnum+"&floor="+f
	}else{
		opener.parent.floor.location="capfloor.asp?bldgnum="+bldgnum+"&riser="+r
	}
	window.close()
}

</script>

</head>
<body bgcolor="#FFFFFF" text="#000000">
<%
if msg <> "" then
%>
<font face="Arial, Helvetica, sans-serif" size="2">Riser_name</font>
<%
end if
%>
<form name="form1" method="post" action="capadditem.asp">
<input type="hidden" name="bldgnum" value="<%=bldgnum%>">
<input type="hidden" name="floor" value="<%=floor%>">
<input type="hidden" name="riser" value="<%=riser%>">
<input type="hidden" name="item" value="<%=item%>">
  <table width="100%" border="0" bgcolor="#3399CC">
    <tr>
      <td>
        <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF"><%=title%> 
          Association</font></b></font></div>
      </td>
    </tr>
  </table>
<br>
<div align="center">
  <table width="100%">
    <tr>
        <td width="40%"><b><font face="Arial, Helvetica, sans-serif"><%=title%>s 
          in this building: </font></b></td>
		<td width="20%">&nbsp;</td>
		<td width="40%"> &nbsp&nbsp </td>
  </tr>
  </table>

   
    <table width="100%" border="0" cellspacing="0" cellpadding="0">
   <%
	do until rst1.eof
    %><tr>
        <td><%=trim(rst1("choice"))%></td>
  </tr><%
	rst1.movenext
    loop
    rst1.close
    %>
</table>
    
    <table width="100%">
  <tr>
        <div align="center"><td width="40%"><b><font face="Arial, Helvetica, sans-serif"><%=title%>s 
          to <%=title2%>&nbsp<%=val%> </font></b></td><td width="20%">&nbsp;</td>
        <td width="40%"><b><font face="Arial, Helvetica, sans-serif"><%=title%>s 
          left</font></b></td></div>
  </tr>
  <tr>
        <td width="40%"> 
          <div align="center">
    <select multiple name="exist" size="5">
    <%
	if not rst3.eof then
	do until rst3.eof
	%>
	  <option value="<%=trim(rst3("choice"))%>"><%=trim(rst3("choice"))%></option>
	<%
	rst3.movenext
	loop
	rst3.close
	end if
	%>
    </select>
	</div>
  </td>
        <td width="20%"> 
          <div align="center"> 
    <input type="submit" name="submit" value="<-     Add"><br><br>
    <input type="submit" name="submit" value="->Delete"></div></td>	
        <td width="40%"> 
          <%
	Set rst2 = Server.CreateObject("ADODB.Recordset")
	if item="riser" then
		sql2 = "SELECT distinct "&item1&" as choice from tblfloor where bldgnum='"& bldgnum &"' and "&item1&" not in ("& sql3&" ) order by "&item1&""
    else
		sql2 = "SELECT distinct "&item1&" as choice from tblriser where bldgnum='"& bldgnum &"' and "&item1&" not in ("& sql3&" ) order by "&item1&""
	end if
	'response.write sql2
	'response.end
	rst2.Open sql2, cnn1, adOpenStatic, adLockReadOnly
	%>
          <div align="center">
	<%
	if not rst2.eof then
	%>
	<select multiple name="list" size="5">
	<%
   	    do until rst2.eof
	%>
	  <option value="<%=trim(rst2("choice"))%>"><%=trim(rst2("choice"))%></option>
	<%
	    rst2.movenext
	    loop
	%>
	</select>
	<%
	else
	%>
	Close window and add <%=title%> first
    <%
	end if
	%>
	</div>
</td></tr>
</table>	
</div>
	
<p>
   <input type="button" name="Submit" value="Close" onClick='javascript:redirect("<%=bldgnum%>", "<%=item%>", "<%=floor%>", "<%=riser%>" )'>
</p>
</form>


<p>&nbsp;</p>
<%
else
%>
<center><font face="Arial, Helvetica, sans-serif"> No <%=title%> available for this building</font></center>
<%
end if
%>
</body>
</html>
