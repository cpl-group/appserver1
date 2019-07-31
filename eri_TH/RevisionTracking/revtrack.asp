<%@Language="VBScript"%>
<html>

<head>
<!--#include file="../../UM/adovbs.inc" -->
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Revision Tracking</title>
</head>
<script> 
function loadentry(id){
	var temp = 'addrev.asp?action=2&id=' + id

	openpage(temp);
}
function openpage(page){
	var w = 500;
	var h = 300;
	var winl = (screen.width - w) / 2;
    var wint = (screen.height - h) / 2;

    winprops = 'height='+h+',width='+w+',top='+wint+',left='+winl+',status=yes,scrollbars='+scroll+',resizable=no'
     // open new window and use the variables to position it
	//popupwin=window.open(page,'login','WIDTH=400, HEIGHT=300, scrollbars=no,left='+x+',top='+y)
	popupwin=window.open(page,'login',winprops)
	popupwin.focus('login')
}
</script>

<body bgcolor="#FFFFFF">
<%
AppRev = 1


Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=security;"
strsql = "select (select count(*) from revtrack where pid in (select portfolio_id from clients where username = '"& Session("loginemail") &"')) as revlevel, * from revtrack where pid in (select portfolio_id from clients where username = '"& Session("loginemail") &"') group by revdate,guid,revdescriptor, id, pid order by revdate desc"
'response.write strsql
'response.end
rst1.Open strsql, cnn1, 0, 1, 1
if not rst1.eof then
%>
<table border="1" width="100%" cellpadding="0" cellspacing="0" bordercolor="#000000" height="46" align="center">
  <tr> 
    <td width="18%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" color="#000000" size="1">Rev. 
      Date </font></td>
    <td width="25%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" size="1">Revision</font></td>
    <td width="28%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" color="#000000" size="1">GUID</font></td>
    <td width="29%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" size="1" color="#000000">Revision 
      Description</font></td>
  </tr>
  <%
Revs = cint(rst1("revlevel")) -1
Do While Not rst1.EOF
 %>
  <tr valign="top" bgcolor="#CCCCCC"> 
    <td width="18%" align="left" height="0"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><i> 
        <%=rst1("revdate")%></i></font></div>
    </td>
    <td width="25%" align="center" height="0"><font face="Arial, Helvetica, sans-serif" size="2">v<%=Apprev%>.<%=Revs%></font></td>
    <td width="28%" align="center" height="0"><font face="Arial, Helvetica, sans-serif" size="2"><i> 
      <%=rst1("guid")%> </i></font></td>
    <td width="29%" align="center" height="0"> 
      <div align="left"><font face="Arial, Helvetica, sans-serif" size="2"><i> 
        <%=rst1("revdescriptor")%> </i></font></div>
    </td>
  </tr>
  <%
  Revs= Revs - 1
rst1.MoveNext  
Loop

rst1.Close
Set rst1 = Nothing
cnn1.Close
Set cnn1 = Nothing

else 
	Response.write "No Revision Records Found"

End if
 %>
</table>
</body>

</html>
