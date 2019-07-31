<%option explicit%>
<!--#include file="adovbs.inc"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'http params
dim crdate,c,total,j,strsql,tot1,sql,atable,prefix


if request("ji")="" then 
	j = request ("jg")
	aTable = "GY_MASTER_JOB"
	prefix = "jg"
else
	j = request ("ji")
	aTable = "IL_MASTER_JOB"
	prefix = "ji"
end if

c = request("c")


'adodb vars
dim cnn, cmd, rs, prm
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")


' open connection
cnn.Open getConnect(0,0,"Intranet")
cnn.CursorLocation = adUseClient

rs.open "SELECT crdate FROM sysobjects WHERE name = 'GY_master_po'",cnn
crdate=rs(0)
rs.close

' specify stored procedure to run based on company

select case c
case "IL" 
strsql = "Select * from ilite.dbo.times where right(jobno,4)='" & right(j,4) & "' order by date desc"
case "GY"
strsql = "Select * from times where jobno='" & right(j,4) & "' order by date desc"
case "NY"
strsql = "Select * from ilite.dbo.times where right(jobno,4)='" & right(j,4) & "' order by date desc"
end select
'response.write strsql
'response.write j
'response.end
rs.open strsql,cnn

%>

<html>
<head>

<script language="JavaScript1.2">
function po_item(c,p,j) {
	theURL="https://appserver1.genergy.com/um/war/po/po_item.asp?c="+c+"&p="+p+"&j=" +j
	openwin(theURL,800,400)
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, scroll=on, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}
</script>
<title>Genergy War Room - Time sheet Detail</title>
</head>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<style type="text/css">
td { font-size:smaller; }
</style>
<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#FFFFFF" class="innerbody">

    <table width="100%" border=0 cellpadding="3" cellspacing="1" bgcolor="#cccccc">
      <tr bgcolor="#228866" style="font-weight:bold;"> 
        <td width="9%"><span class="standardheader">Date</span></td>
        <td width="50%"><span class="standardheader">Description</span></td>
        <td width="11%"><span class="standardheader">Hours</span></td>
        <td width="11%"><span class="standardheader">Overtime</span></td>
        <td width="19%"><span class="standardheader">User</span></td>
      </tr>
      <%
total = 0

while not rs.EOF 
total = total + cdbl(rs("hours"))
tot1 = tot1 + cdbl(rs("overt"))

%>
      
      <tr bgcolor="#ffffff"> 
        <td> <%=rs("date")%> </td>
        <td> <%=left(rs("description"),55)%> </td>
        <td align="right"><%=formatnumber(rs("hours"),2)%> </td>
        <td align="right"><%=formatnumber(rs("overt"),2)%></td>
        <td><%=rs("matricola")%></td>
      </tr>
        <%
 rs.movenext


wend
%>
      <tr bgcolor="#ffffff"> 
        <td colspan="2"><b>Totals for job number <%=j%></td>
        <td align="right"><b><%=formatnumber(total,2)%></b></td>
        <td align="right"><b><%=formatnumber(tot1,2)%></b></td>
        <td height="23" width="19%">&nbsp;</td>
      </tr>
    </table>
	<br>
	<p>Update as of now!</p>
	<br><br>
	  <%
	  set cnn = nothing %>
	  
</body>
</html>
