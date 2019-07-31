<!-- #include file="./adovbs.inc" -->
<% 
leaseid= Request("leaseid")
profiletype=Request("profiletype")
portfolio=Request("portfolioid")
%> 
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function meterfill(leaseid,bldg, profiletype, portfolio){
	document.location.href="lmp.asp?leaseid=" + leaseid + "&bldg=" + bldg + "&profiletype=" + profiletype + "&portfolio=" + portfolio;
}
function loadmeter(param,script, dsn){
			var arrayparam=param.split("_")
			var id=arrayparam[0]
			var bldg=arrayparam[1]
			var temp= "http://www.genergy.com/cgi-bin/" + script + "?bldg=" + bldg + "&lid=" + id +"&dsn=" + dsn
			document.frames.lmp.location.href=temp;
	}
function bldglmp(id,script, dsn){
			var temp = "http://www.genergy.com/cgi-bin/" + script + "?lid=" + id +"&dsn=" + dsn
						
			document.frames.lmp.location.href=temp;
	}
</script>
</head><style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>


<body bgcolor="#FFFFFF" text="#000000">
<div align="center">
  <table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td bgcolor="#0099FF"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="4" color="#000000">Load 
          Management Profiles</font></div>
      </td>
    </tr>
  </table>
  
  <form method="post" action="" name="lmp">
    <div align="center"> 
      <h2> 
        <input type="hidden" name="lmp" value=<%=portfolio%>>
        <input type="hidden" name="profiletype" value=<%=Request("profiletype")%>>
        <font face="Arial, Helvetica, sans-serif" size="3"> 
        <%
Set cnn1 = Server.CreateObject("ADODB.Connection")
openStr="data Source=eri;user Id=web"
cnn1.Open openStr

Set rsCat1 = Server.CreateObject("ADODB.Recordset")

sql="SELECT Distinct Management FROM Buildings WHERE owner_id = N'" & request("portfolioid") & "'"

rsCat1.Open sql, cnn1, adOpenStatic, adLockReadOnly

Response.Write rsCat1("management") 

rscat1.close

%>
        </font></h2>
      <h2><font face="Arial, Helvetica, sans-serif" size="3"> 
		  <input type="hidden" name="script1" value="ptstart.cgi">
		  <input type="hidden" name="scriptlmp" value="proagrstart.cgi">
          <input type="hidden" name="dsn" value="sqlserverg1">
        <select name="meter" onChange="loadmeter(this.value,script1.value, dsn.value)">
          <% 	 
				Set cnn1 = Server.CreateObject("ADODB.Connection")
				Set rst1 = Server.CreateObject("ADODB.recordset")
				cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
		
		   		strsql = "SELECT DISTINCT Meters.MeterId, Meters.MeterNum, Meters.BldgNum FROM Meters INNER JOIN Buildings ON Meters.BldgNum = Buildings.BldgNum INNER JOIN tblProfile ON Meters.MeterId = tblProfile.MeterId WHERE (Buildings.portfolioid = '" & portfolio & "') AND (Meters.PP = 1)"
				
				response.write strsql
				
				rst1.Open strsql, cnn1, adOpenStatic
 %>
	          <option value="<%=rst1("meterid")%>">Select Building</option>
 <%				
				do until rst1.eof
			%>
          <option value="<%=rst1("meterid")%>_<%=rst1("bldgnum")%>"><font face="Arial, Helvetica, sans-serif" size="3"><%=rst1("meternum") %></font></option>
          <% 
				rst1.movenext
				loop
			%>
        </select>
        </font></h2>
      <h2> <font face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif" size="3"> 
        <input type="button" name="Button" value="Portfolio LMP" onClick="bldglmp(lmp.value,scriptlmp.value, dsn.value)">
        </font></font></font><font face="Arial, Helvetica, sans-serif" size="3"> 
        </font> </h2>
    </div>
    </form>
</div>
<IFRAME name="lmp" src="null.htm" width="100%" height="100%" scrolling="auto" marginwidth="8" marginheight="16"></IFRAME>
<%
rst1.close
%>
</body>
</html>
