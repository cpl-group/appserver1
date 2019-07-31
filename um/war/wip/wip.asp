<%option explicit%>
<!--#include file="adovbs.inc"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'http params
dim c,jt,j,strsql,tot1,sql,atable,prefix

c = request("c")
jt= request ("jt")

'adodb vars
dim cnn, cmd, rs, prm
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")




' open connection
cnn.Open getConnect(0,0,"Intranet")
cnn.CursorLocation = adUseClient

'rs.open "SELECT crdate FROM sysobjects WHERE name = 'GY_master_po'",cnn
'crdate=rs(0)
'rs.close

' specify stored procedure to run based on company
if c="gy" or c="GY" then
select case jt

case "ERI"
strsql = "Select * from master_job where company='GY' and status <>'closed' and type  like 'ES-ERI%'  order by id"
case "DWG"
strsql = "Select * from master_job where company='GY' and status <>'closed' and type  like 'TE-DWG%'  order by id"
case "MAC"
strsql = "Select * from master_job where company='GY' and status <>'closed' and type  like 'ES-MAC%'  order by id"
case "R&B"
strsql = "Select * from master_job where company='GY' and status <>'closed' and type  like 'ES-R&B%'  order by id"
case "RFP" 
strsql = "Select * from master_job where company='GY' and status <>'closed' and type  like '%-RFP-%'  order by id"
case "GY" 
strsql = "Select * from master_job where company='GY' and status <>'closed' and type  like 'GY-%'  order by id"
case "G1" 
strsql = "Select * from master_job where company='GY' and status <>'closed' and type  like 'G1-%'  order by id"
case "IT" 
strsql = "Select * from master_job where company='GY' and status <>'closed' and type  like 'IT-%'  order by id"

case "WIP"

strsql = "Select * from master_job where company='GY' and type not like '%R&B%' and type not like '%MAC%'and type not like '%DWG%'and type not like '%ALL%'and type not like '%ERI%' and type not like 'G1%' and type not like '%DWG%' and type not like '%RFP%' and status <> 'closed' order by id"

end select

end if


if c="il" or c="IL" then

select case jt
case "RFP" 
strsql = "Select * from master_job where company='IL' and status <>'closed' and type  like '%-RFP-%'  order by id"

case else
strsql = "Select * from master_job where company='IL' and status <>'closed' and type not like '%RFP%'  order by id"
end select


end if

if c="ny" or c="NY" then

select case jt
case "RFP" 
strsql = "Select * from master_job where company='NY' and status <>'closed' and type  like '%-RFP-%'  order by id"

case else
strsql = "Select * from master_job where company='NY' and status <>'closed' and type not like '%RFP%'  order by id"
end select


end if



'response.write  strsql
'response.end
rs.open strsql,cnn

%>

<html>
<head>

<script language="JavaScript1.2">
function job(j) {
	theURL="http://appserver1.genergy.com/genergy2_intranet/opsmanager/joblog/viewjob.asp?jid=" +j
	openwin(theURL,800,400)
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, scroll=on, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}
</script>
<title>Genergy War Room - Open jobs</title>
</head>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<style type="text/css">
td { font-size:smaller; }
</style>
<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#FFFFFF" class="innerbody">


    <table width="100%" border=0 cellpadding="3" cellspacing="1" bgcolor="#cccccc">
      <tr bgcolor="#228866" style="font-weight:bold;"> 
        <td width="9%"><span class="standardheader">Job #</span></td>
        <td width="25%"><span class="standardheader">Description</span></td>
        <td width="25%"><span class="standardheader">Job Address</span></td>
        <td width="11%"><span class="standardheader">PM</span></td>
        <td width="11%"><span class="standardheader">% Complete</span></td>
      </tr>
      <%


while not rs.EOF 



%>
      
      <tr bgcolor="#ffffff"> 
        <td> <a href="javascript:job('<%=rs("id")%>')"><%=rs("job")%></a></td>
        <td> <%=rs("description")%> </td>
        <td><%=rs("address_1") & " " &rs("address_2")%></td>
        <td> <%=rs("pm_last") & " " &rs("pm_first")%></td>
        <td align="right"><%=rs("percent_complete")%></td>
        <%
 rs.movenext


wend
%>
    </table>
      
        
    <p>&nbsp;&nbsp;Update as of now!</p>
    
	<p>&nbsp;</p>
	
	
	  <%
	  set cnn = nothing %>
	  
</body>
</html>
