<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<html>
<head>
<title>Lamping Report</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/styles.css">
</head>

<body bgcolor="#FFFFFF" text="#000000">


<%
floor= Request.Querystring("floor")
bldgnum=Request.Querystring("bldgnum")
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
Set rst3 = Server.CreateObject("ADODB.recordset")
Set rst4 = Server.CreateObject("ADODB.recordset")
Set rst = Server.CreateObject("ADODB.recordset")



cnn1.Open getconnect(0,0,"engineering")

if floor<>"All" then
sqlstr = "select distinct F.room ,R.ROOM AS ROOMN,FF.FLOOR AS FLOOR from fixtures F JOIN ROOM R ON F.ROOM=R.ID JOIN FLOOR FF ON R.FLOOR=FF.ID where R.floor='"& floor &"' and BLDGNUM='"&bldgnum&"' GROUP BY F.room ,R.ROOM,FF.FLOOR"
else
sqlstr = "SELECT DISTINCT F.ID AS FID,F.FLOOR AS FLOOR,b.bldgname  FROM FLOOR F JOIN ROOM  R ON F.ID=R.FLOOR JOIN FIXTURES FX ON R.ID=FX.ROOM JOIN  facilityinfo b on fx.bldgnum=b.id WHERE BLDGNUM='"&bldgnum&"'  order by floor"
end if
'response.write sqlstr
'response.end

rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.EOF then
%>
<table width="100%" cellpadding="3" cellspacing="0" border="0">
  <tr>
    <td><span class="standard">No fixtures found.</span></td>
  </tr>
</table>
<%
else
if floor<>"All" then
%><form name="form2" method="post" action="">

<table width="100%" cellpadding="3" cellspacing="0" border="0" style="border:1px solid #ffffff">
  <tr>
      <td bgcolor="#336699"> <font face="Arial, Helvetica, sans-serif" color="#ffffff"><span class="standard"><b>Lamping 
        Report | Floor: <%=rst1("FLOOR")%></b></span></font></td>
  </tr>
  </table>
<%
flr=rst1("floor")
%>

<% While not rst1.EOF 
rm=rst1("room")%>
<table width="100%" border="0" cellpadding="3" cellspacing="0" style="border:1px solid #ffffff">
<tr bgcolor="#cccccc">
	<td width="55%"><font face="Arial, Helvetica, sans-serif"  size="2"><span class="standard"><b>Room: <%=rst1("roomN")%></b></span></font></td>
	<td width="45%" align="right"></td>
</tr>
</table>

<table width="100%" border="0" cellpadding="3" cellspacing="1">
<%
Set rst4 = Server.CreateObject("ADODB.recordset")
sqlstr = "select ft.*,f.* ,l.*, DATEADD(week,(est_lamp_life/est_hr_wk) , datelastchanged)as estd , datediff(week,getdate(), (DATEADD(week,(est_lamp_life/est_hr_wk) , datelastchanged))) as weeksr,datediff(week,getdate(),(DATEADD(week,(est_lamp_life/est_hr_wk) , datelastchanged)))* est_hr_wk as hoursr  from fixture_types ft join fixtures  f on ft.id=f.typeid join lamping_sch l on f.id=l.fid where f.room='"&rm&"' and bldgnum='"&bldgnum&"'  order by l.id desc"

rst4.Open sqlstr, cnn1, 0, 1, 1

%>

       <tr bgcolor="#eeeeee" valign="top">
	       
      <td width="40%" height="15"><font face="Arial, Helvetica, sans-serif" size="2"><span class="shrunkenheader">Fixture</span></font></td>
	       
      <td width="15%" height="15"><font face="Arial, Helvetica, sans-serif" size="2"><span class="shrunkenheader">Date 
        Last Changed</span></font></td>
	       
      <td width="15%" height="15"><font face="Arial, Helvetica, sans-serif" size="2"><span class="shrunkenheader">Estimated 
        Change Date</span></font></td>
	       
      <td width="15%" height="15"><font face="Arial, Helvetica, sans-serif" size="2"><span class="shrunkenheader">Remaining 
        Weeks</span></font></td>
	       
      <td width="15%" height="15"><font face="Arial, Helvetica, sans-serif" size="2"><span class="shrunkenheader">Remaining 
        Hours</span></font></td>
        </tr>
        <% While not rst4.EOF %>
        <tr align="left" valign="top"> 
            <td class="bottomline"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><%=rst4("fix_catalog")%>/<%=rst4("lamp_catalog")%></span></font></td>
            <td class="bottomline"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><%=rst4("datelastchanged")%></span></font></td>
            <td class="bottomline"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><%=rst4("estd")%></span></font></td>
            <td class="bottomline"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><%=rst4("weeksr")%></span></font></td>
            <td class="bottomline"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><%=rst4("hoursr")%></span></font></td>
		</tr>
        <% 
		rst4.movenext
		Wend
		rst4.close
%>
</table>

  </form>
 <%
rst1.movenext
Wend

rst1.close 



else%>

<form name="form1" method="post" action="">

<table width="100%" cellpadding="3" cellspacing="0" border="0" style="border:1px solid #ffffff">
<tr>
	<td bgcolor="#336699"><font face="Arial, Helvetica, sans-serif" color="#ffffff"><span class="standard"><b>Lamping Report | Building: <%=rst1("bldgname")%></b></span></font></td>
</tr>
</table>
       

    <% While not rst1.EOF %>

	<table width="100%" border="0" cellpadding="3" cellspacing="1">
    <tr> 
      <td><font face="Arial, Helvetica, sans-serif"><span class="standard"><b>Floor: 
        <a href="lampreport.asp?bldgnum=<%=bldgnum%>&floor=<%=rst1("FID")%>" class="floorlink"><%=rst1("floor")%></a></b></span></font></td>
    </tr>
	</table>
       
	
	  
<%fl=rst1("FID")%>


<%Set rst3 = Server.CreateObject("ADODB.recordset")
sqlstr = "select distinct F.room as room,R.ROOM AS ROOMN from fixtures F JOIN ROOM R ON F.ROOM=R.ID where R.floor='"&fl&"'"
rst3.Open sqlstr, cnn1, 0, 1, 1%>
<% While not rst3.EOF %>
	<table width="100%" border="0" cellpadding="3" cellspacing="0" style="border:1px solid #ffffff">
    <tr  bgcolor="#cccccc">
      <td colspan="2"><font face="Arial, Helvetica, sans-serif"><span class="standard"><b>Room: <%=rst3("roomN")%></b></span></font></td>
      <td colspan="3" align="right"></td>
    </tr>
	</table>
	<table width="100%" border="0" cellpadding="3" cellspacing="1">
       <tr bgcolor="#eeeeee" valign="top">
	       <td width="25%"><font face="Arial, Helvetica, sans-serif" size="2"><span class="shrunkenheader"><b>Fixture</b></span></font></td>
	       <td width="15%"><font face="Arial, Helvetica, sans-serif" size="2"><span class="shrunkenheader"><b>Date Last Changed</b></span></font></td>
	       <td width="15%"><font face="Arial, Helvetica, sans-serif" size="2"><span class="shrunkenheader"><b>Estimated Change Date</b></span></font></td>
	       <td width="15%"><font face="Arial, Helvetica, sans-serif" size="2"><span class="shrunkenheader"><b>Remaining Weeks</b></span></font></td>
	       <td width="15%"><font face="Arial, Helvetica, sans-serif" size="2"><span class="shrunkenheader"><b>Remaining Hours</b></span></font></td>
        </tr>

  
<%rm=rst3("room")
Set rst4 = Server.CreateObject("ADODB.recordset")
sqlstr = "select ft.*,f.* ,l.*, DATEADD(week,(est_lamp_life/est_hr_wk) , datelastchanged)as estd , datediff(week,getdate(), (DATEADD(week,(est_lamp_life/est_hr_wk) , datelastchanged))) as weeksr,datediff(week,getdate(),(DATEADD(week,(est_lamp_life/est_hr_wk) , datelastchanged)))* est_hr_wk as hoursr  from fixture_types ft join fixtures  f on ft.id=f.typeid join lamping_sch l on f.id=l.fid  where f.room='"&rm&"' and bldgnum='"&bldgnum&"'  order by l.id desc"
rst4.Open sqlstr, cnn1, 0, 1, 1
%>

        <% While not rst4.EOF %>
        <tr align="left" valign="top"> 
            <td class="bottomline"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><%=rst4("fix_catalog")%>/<%=rst4("lamp_catalog")%></span></font></td>
            <td class="bottomline"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><%=rst4("datelastchanged")%></span></font></td>
            <td class="bottomline"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><%=rst4("estd")%></span></font></td>
            <td class="bottomline"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><% if not rst4("weeksr") > 0 then%> 0 (<%=abs(rst4("weeksr"))%> weeks overdue)<% else %><%=rst4("weeksr")%><% end if %></span></font></td>
            <td class="bottomline"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"><%=rst4("hoursr")%></span></font></td>
		</tr>
        <% 
		rst4.movenext
		Wend
		rst4.close
		
rst3.movenext
Wend
%>
		
	</table>
<br>
  </form>

 <%
rst1.movenext
Wend
rst3.close
rst1.close 


end if
end if
%>

</body>
</html>