<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("ts") < 4 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				Response.Redirect "../main.asp"
			end if	
		end if	

name=Trim(Session("name"))

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

function otherUser(){
    document.location="timeprint.asp?other=1"
}

function Printer(name){
//alert(name)
//    window.location="http://www.hotmail.com"
    window.location="timetemplate.asp?name="+name
    self.moveTo(0, 0)
    window.resizeTo(800, 600)
}
</script>

</head>
<body bgcolor="#FFFFFF" text="#000000">
<div align="center">
  <table width="55%" border="0">
  <%
    if not (Request("other")=1) Then
  %>
    <tr>
      <td width="45%"> 
        <div align="center"> 
          <input type="button" name="b1" value="Other User" onClick="otherUser()">
        </div>
    </td>
      <td width="38%"> 
        <div align="center">
	  <input type="hidden" name="temp" value="<%=name%>">
          <input type="button" name="b12" value="Current User" onClick="Printer(temp.value)">
        </div>
    </td>
      <td width="17%">&nbsp;</td>
  </tr>
  <%
  else
  Set cnn1 = Server.CreateObject("ADODB.Connection")
  Set rst1 = Server.CreateObject("ADODB.Recordset")
  Set rst2 = Server.CreateObject("ADODB.Recordset")
  cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"
  sql="select name from user_cost where name != '"& name &"' order by name"
  rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly
  %>
  <tr>
  	  <td width="45%"> 
        <div align="center">User Name </div>
      </td>
	  <td width="38%"> 
        <div align="center"> 
          <select name="username">
            <%
	Do until rst1.eof
	%>
            <option value="<%=rst1("name")%>"><%=rst1("name")%> </option>
            <%
	rst1.movenext
	loop
	%>
          </select>
        </div>
      </td>
	  <td width="17%"> 
        <div align="center">
          <input type="button" name="Submit2" value="Print" onClick="Printer(username.value)">
        </div>
      </td> 
    </tr>
  <%
  end if
  %>
</table>
<br><br>
</div>
<input type="button" name="Submit" value="Cancel" onClick="window.close()">
</body>
</html>
