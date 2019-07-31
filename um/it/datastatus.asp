<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"dbCore")

sqlstr = "Select * from dbBilling.dbo.it_datastatus where meterid<>0 order by maxdate,bldgnum,utility"

rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1


if len(time) = 11 then 
	rtic = 11
	tic = 2
else
	rtic = 10
	tic = 1
end if

%>
<HTML>

<HEAD>
<TITLE>DATA STATUS MONITOR</TITLE>
<script>
<!--


//enter refresh time in "minutes:seconds" Minutes should range from 0 to inifinity. Seconds should range from 0 to 59
var limit="0:30"

if (document.images){
var parselimit=limit.split(":")
parselimit=parselimit[0]*60+parselimit[1]*1
}
function beginrefresh(){
if (!document.images)
return
if (parselimit==1)
window.location.reload()
else{ 
parselimit-=1
curmin=Math.floor(parselimit/60)
cursec=parselimit%60
if (curmin!=0)
curtime=curmin+" minutes and "+cursec+" seconds left until page refresh!"
else
curtime=cursec+" seconds left until page refresh!"
window.status=curtime
setTimeout("beginrefresh()",1000)
}
}

window.onload=beginrefresh
//-->
</script>
</HEAD>
<body id="datapage" bgcolor="#FFFFFF">

<form name="form1" method="post" action="">
  <table width="800" border="0" align="center" cellpadding="0" cellspacing="0" id="datatable" style="border:5px solid white">
    <tr> 
      <td bgcolor="#3399CC" height="36" width="13%"> 
        <div align="center"> <font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF" size="4"> 
          DATA STATUS MONITOR AS OF <%=date() & " " & time()%></font></b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="87%"> 
        <table width="100%" border="1" cellpadding="0" cellspacing="0" bordercolor="#999999">
          <tr> 
            <td bgcolor="#CCCCCC" width="25%" height="17"><div align="center"><strong><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Building 
                </font></strong></div></td>
            <td width="10%" height="17" bgcolor="#CCCCCC"> <div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif"><font color="#000000"> 
                Bldg # </font></font></strong></div></td>
            <td  width="10%" height="17" bgcolor="#CCCCCC"><div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">System 
                </font></strong></div></td>
            <td bgcolor="#CCCCCC"><div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif">Utility 
                </font></strong></div></td>
            <td width="15%" height="17" bgcolor="#CCCCCC"> <div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif"> 
                Meter Count</font></strong></div></td>
            <td width="15%" height="17" bgcolor="#CCCCCC"> <div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif"><font color="#000000">LMP 
                Meter ID </font></font></strong></div></td>
            <td width="15%" height="17" bgcolor="#CCCCCC"> <div align="center"><strong><font size="2" face="Arial, Helvetica, sans-serif"><font color="#000000">Last 
                Data Import</font></font></strong></div></td>
          </tr>
          <% While not rst1.EOF %>
          <tr <%if rst1("lm") = 1 then 
		  		system="Real-time"  
				response.write "bgcolor='#99CCCC'"
			else 
				system ="Day Behind"
				response.write "bgcolor='white'" 
				 end if %>> 
            <td width="5%"><font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("strt")%></a></font></td>
            <td width="2%"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("bldgnum")%></a></font></div></td>
            <td width="3%"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=system%></font></div></td>
            <td width="3%"> <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("utility")%></font></div></td>
            <td width="3%"> <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("metercnt")%></font></div></td>
            <td width="0%"> <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("meterid")%></font></div></td>
            <td width="92%"> <div align="right"><font face="Arial, Helvetica, sans-serif" 
			  <%if (instr(rst1("maxdate"),date()) > 0 and instr(left(right(rst1("maxdate"),rtic),tic),trim(left(time,tic))) > 0) then %> 
			  color="#000000" 
			  <% 

			  else
			  if rst1("lm") = 1 then 
			  %>
			  color="#FF0000" 
			  <%			  
			 problem = 1 
			 end if
			end if %> size="1"> <%=rst1("maxdate")%> 
                <%if rst1("maxdate") <> rst1("lmpmaxdate") then response.write "*" end if %></font> 
              </div></td>
          </tr>
          <%
		  	rst1.movenext
			wend%>
        </table>
      </td>
    </tr>
    <tr> 
      <td height="23" bgcolor="#3399CC"><font face="Arial, Helvetica, sans-serif" size="1">NOTE: 
        Real-time systems are highlighted green and Day-Behind systems are highlighted white</font></td>
    </tr>
  </table>
</form>
<%
rst1.close
%>
<script language="JavaScript1.2">
//function flashit(){
//if (!document.all)
//return
//if (document.bgColor="#FFFFFF")
//document.bgColor=<%if problem=1 then %>"red"<%else%> "#FFFFFF" <% end if %>
//else
//document.bgColor="#FFFFFF"
//}
//setInterval("flashit()", 3000)
//-->
</script>
</body>
