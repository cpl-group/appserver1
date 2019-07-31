<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
			Response.Redirect "http://www.genergyonline.com"
		end if		
%>
<html>

<head>
<!--#include file="../adovbs.inc" -->
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
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=security;"
strsql = "select * from revtrack where pid ='"& Request("pid") & "'"
rst1.Open strsql, cnn1, 0, 1, 1
%>
<a href="javascript:openpage('addrev.asp?pid=<%=Request("pid")%>')"><font face="Arial, Helvetica, sans-serif" size="2"><b>Add 
Revision</b></font></a> 
<%
if not rst1.eof then
%>
<table border="1" width="100%" cellpadding="0" cellspacing="0" bordercolor="#000000" height="46" align="center">
  <tr> 
    <td width="22%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" color="#000000" size="2">Rev. 
      Date </font></td>
    <td width="39%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" color="#000000" size="2">GUID</font></td>
    <td width="39%" align="center" bgcolor="#66CCFF" height="0"><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Revision 
      Description</font></td>
  </tr>
    <%
Do While Not rst1.EOF
 %>
        <tr valign="top" style="cursor:hand" onmouseover="this.style.backgroundColor='lightgreen'" onmouseout="this.style.backgroundColor='CCCCCC'; " onclick="javascript:loadentry('<%=rst1("id")%>')" bgcolor="#CCCCCC"> 
      <td width="22%" align="left" height="0"><font face="Arial"><i> <%=rst1("revdate")%> 
        </i></font></td>
      <td width="39%" align="center" height="0"><font face="Arial"><i> <%=rst1("guid")%> 
        </i></font></td>
      <td width="39%" align="center" height="0"><font face="Arial"><i> <%=rst1("revdescriptor")%> 
        </i></font></td>
  </tr>
  <%
rst1.MoveNext  
Loop

rst1.Close
Set rst1 = Nothing
cnn1.Close
Set cnn1 = Nothing


End if
 %>
</table>
</body>

</html>
