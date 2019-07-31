<%@Language="VBScript"%>
<!-- #include file="./adovbs.inc" -->
<%
		if isempty(Session("name")) then
%>
<script>
opener.location="../index.asp"
window.close()
</script>
<%
			'			Response.Redirect "http://www.genergyonline.com"
		end if
		
		if Session("eri") > 2 then 	
		type1=Request.QueryString("type1")
		'des=Request.QueryString("description")
		count=Request("count")
		'Response.Write count
		
%>
<html>
<head>
<title>Descriptions</title>

<meta http-equiv="Content-type1" content="text/html; charset=iso-8859-1">
<script>
function newdescription(type,count){

     var temp = "addnewdesc.asp?type=" + type + "&count=" + count
     document.location.href = temp;

}
</script>


</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post" action="">
  <table width="100%" >
    <tr> 
      <td width="20%"> <font face="Arial, Helvetica, sans-serif"> 
	  <input type="hidden" name="type1" value="<%=type1%>">
	  <input type="hidden" name="count" value="<%=count%>">			
	<input type="button" name="Button" value="Add Description" onClick="newdescription(type1.value,count.value)">
        </font></td>
    </tr>
  </table>
</form>  
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")


cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"



sql = "SELECT description FROM tblSurveyLib WHERE(type='"& type1 & "') order by description"


Set rst1 = Server.CreateObject("ADODB.Recordset")
rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly


If not rst1.EOF then
    Do until rst1.EOF 
	%>
	<a href="fillup.asp?count=<%Response.Write(count)%>&type=<%Response.Write(type1)%>&description=<%=rst1("description")%>" ><%=rst1("description")%></a><br>
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
</body>
</html>
