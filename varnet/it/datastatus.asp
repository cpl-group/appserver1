<%@Language="VBScript"%>
<%
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

sqlstr = " Select Strt, * from it_datastatus join buildings on it_datastatus.bldgnum=buildings.bldgnum order by strt "

rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1
if len(time) = 11 then 
	tic = 2
else
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
<body bgcolor="#FFFFFF">

<form name="form1" method="post" action="">
  <table width="800" border="0" id="datatable" style="border:5px solid white" align="center">
    <tr> 
      <td bgcolor="#3399CC" height="36" width="13%"> 
        <div align="center"> <font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF" size="4"> 
          DATA STATUS MONITOR AS OF <%=time%></font></b></font></div>
      </td>
    </tr>
    <tr> 
      <td width="87%"> 
        <table width="100%" border="0">
          <tr> 
            <td bgcolor="#CCCCCC" width="25%" height="17"><b><font face="Arial, Helvetica, sans-serif" size="2" color="#000000">Building 
              </font></b></td>
            <td width="25%" height="17" bgcolor="#CCCCCC"> 
              <div align="center"><b><font face="Arial, Helvetica, sans-serif" size="2"> 
                Meter Count</font></b></div>
            </td>
            <td width="25%" height="17" bgcolor="#CCCCCC"> 
              <div align="center"><font size="2"><b><font face="Arial, Helvetica, sans-serif"><font color="#000000">LMP 
                Meter ID </font></font></b></font></div>
            </td>
            <td width="25%" height="17" bgcolor="#CCCCCC"> 
              <div align="right"><font size="2"><b><font face="Arial, Helvetica, sans-serif"><font color="#000000">Last 
                Data Import</font></font></b></font></div>
            </td>
          </tr>
          <% While not rst1.EOF %>
          <tr <%if rst1("lm") = 1 then %> bgcolor="#99CCCC" <% end if %>> 
            <td width="2%"><font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("strt")%></a></font></td>
            <td width="3%"> 
              <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("metercnt")%></font></div>
            </td>
            <td width="0%"> 
              <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("meterid")%></font></div>
            </td>
            <td width="95%"> 
              <div align="right"><font face="Arial, Helvetica, sans-serif" 
			  <%if (instr(rst1("maxdate"),date()) > 0 and instr(rst1("maxdate"),trim(left(time,tic))) > 0 and rst1("lm")=1) or ((instr(rst1("maxdate"),date()-1) > 0 or instr(rst1("maxdate"),date()) > 0) and rst1("lm") = 0) then %> 
			  color="#000000" 
			  <% 
			  else
			  %>
			  color="#FF0000" 
			  <%			  
			 problem = 1 
			end if %> size="1"> <%=rst1("maxdate")%> 
                <%if rst1("maxdate") <> rst1("lmpmaxdate") then response.write "*" end if %></font>
              </div>
            </td>
          </tr>
          <%
		  	rst1.movenext
			wend%>
        </table>
      </td>
    </tr>
    <tr> 
      <td height="23" bgcolor="#3399CC"><font face="Arial, Helvetica, sans-serif" size="1">NOTE: 
        PT SYSTEMS ARE MONITORED FOR DATA EXISTING FOR THE PRIOR DAY, LMP SYSTEMS 
        ARE MONITORED FOR DATA EXISTING IN THE LAST HOUR</font></td>
    </tr>
  </table>
</form>
<%
rst1.close
%><script language="JavaScript1.2">
function flashit(){
if (!document.all)
return
if (datatable.style.borderColor=="white")
datatable.style.borderColor=<%if problem=1 then %>"red"<%else%> "white" <% end if %>
else
datatable.style.borderColor="white"
}
setInterval("flashit()", 500)
//-->
</script>

</body>
