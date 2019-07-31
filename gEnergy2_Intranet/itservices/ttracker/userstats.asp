<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
		dim cnn, rs, sql,totaltickets,L24,LWK,GWK, listlength, currentprocess, currentstatus
		
		set cnn = server.createobject("ADODB.Connection")
		set rs = server.createobject("ADODB.Recordset")
		' open connection
		cnn.open getConnect(0,0,"dbCore")
		
		if checkgroup("IT Services") or checkgroup("IT Consultants") or checkgroup("Department Supervisors") then 

        sql = "select t.userid, isnull(a.opencount,0) as opentics, isnull(b.closedcount,0) as closedtics, isnull(cast((cast(c.dl24 as decimal(18,2))/cast(b.closedcount as decimal(18,2))) as decimal(18,2)),0) as [%L24], isnull(cast((cast(d.lwk as decimal(18,2))/cast(b.closedcount as decimal(18,2))) as decimal(18,2)),0) as [%lwk], isnull(cast((cast(e.gwk as decimal(18,2))/cast(b.closedcount as decimal(18,2))) as decimal(18,2)),0) as [%gwk] from tickets t full join (select userid, count(*) as opencount from tickets where closed = 0 and runticket = 0 group by tickets.userid) a on t.userid=a.userid full join (select userid, count(*) as closedcount from tickets where closed = 1 and runticket = 0 group by tickets.userid) b on t.userid=b.userid full join (select userid, count(*) as dl24 from tickets where closed=1 and runticket = 0 and datediff(day, duedate, fixdate) <= 1  group by tickets.userid) c on t.userid=c.userid full join (select userid, count(*) as lwk from tickets where closed=1 and runticket = 0 and (datediff(day, duedate, fixdate) > 1 and datediff(day, duedate, fixdate) <=7) group by tickets.userid) d on t.userid=d.userid full join (select userid, count(*) as gwk from tickets where closed=1 and runticket = 0 and datediff(day, duedate, fixdate) > 7  group by tickets.userid) e on t.userid=e.userid group by t.userid, a.opencount,b.closedcount,c.dl24,d.lwk,e.gwk order by b.closedcount desc"
		else
			response.write "You do not have access to this view"
			response.end
		end if
		
		rs.open sql, cnn 
 		if not rs.EOF then 
			%>
			<title>Trouble Ticket User Statistics</title>
<link rel="Stylesheet" href="../../styles.css" type="text/css">
			<style type="text/css">
			<!--
			BODY {
			SCROLLBAR-FACE-COLOR: #dddddd;
			SCROLLBAR-HIGHLIGHT-COLOR: #ffffff;
			SCROLLBAR-SHADOW-COLOR: #eeeeee;
			SCROLLBAR-3DLIGHT-COLOR: #999999;
			SCROLLBAR-ARROW-COLOR: #000000;
			SCROLLBAR-TRACK-COLOR: #336699;
			SCROLLBAR-DARKSHADOW-COLOR: #333333;
			}
			
			td.red {color: red}
			-->
			</style>
			</head>
			
	<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#eeeeee">
      
<table border=0 cellpadding="3" cellspacing="1" width="100%">
  <tr bgcolor="#6699cc"> 
    <td width="10%" align="center"><span class="standardheader">UserID</span></td>
    <td width="13%" align="center"><span class="standardheader"> Open Tickets</span></td>
    <td width="13%" align="center"><span class="standardheader"> Closed Tickets 
      </span></td>
    <td width="21%" align="center" bgcolor="#66ff66">% Closed in Less Than 24 Hours</td>
    <td width="17%" align="center" bgcolor="#ffcc00">% Closed in Less Than 1 Week</td>
    <td width="19%" align="center" bgcolor="#cc0033"><font color="#FFFFFF">% Closed 
      in Greater Than 1 Week</font></td>
  </tr>
</table>
  <div style="width:100%; overflow:auto; height:70%;border-bottom:1px solid #cccccc;">
      
  <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#dddddd">
    <% 
			while not rs.EOF 
	%>
    <!--tr bgcolor="#ffffff" valign="top" onMouseOver="this.style.backgroundColor = '#6699cc'" style="cursor:crosshair" onMouseOut="this.style.backgroundColor = 'white'"-->
	<tr bgcolor="#ffffff" valign="top"> 
      <td width="10%" height="24"><%=rs("userid")%></td>
      <td width="13%" align="right"><div align="center"><%=rs("opentics")%></div></td>
      <td width="13%"><div align="center"><%=rs("closedtics")%></div></td>
      <td width="21%"><div align="center"><%=formatpercent(cdbl(rs("%l24")),0)%></div></td>
      <td width="17%" align="right"><div align="center"><%=formatpercent(cdbl(rs("%lwk")),0)%></div></td>
      <td width="19%" align="right"><div align="center"><%=formatpercent(cdbl(rs("%gwk")),0)%></div></td>
    </tr>
    <% 
		  rs.movenext
		  wend
    %>
  </table>  
</div> 
</body>
</html>
<%
			rs.close
		else
		response.write "<font size=1 face='Arial, Helvetica, sans-serif'> No tickets found</font>"
		end if 
%>