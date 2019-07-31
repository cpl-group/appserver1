<html>
<head>
<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("um") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
		user=Session("name")
%>

<title>Operations Log</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function viewrpt(bldg,bprd,byr){

	if (bldg == "" || byr == "" || bprd == "" ){
			var temp= "Please Complete all Fields."
			alert(temp)	
	} else {
			var temp="meterproblemreport.asp?bldg="+bldg+"&year="+byr+"&period="+bprd
			document.frames.admin.location=temp
	}

}
</script>
</head>
<%
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

		
%>
<body bgcolor="#FFFFFF" text="#000000">
<font face="Arial, Helvetica, sans-serif"> </font> 
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC"> 
      <div align="center"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif">Meter 
        Problems Reports</font></b></font></div>
    </td>
  </tr>
</table>
<form name="form1" method="post" action="">
  <table width="100%" border="0">
    <tr>
      <td width="48%" height="2"><font face="Arial, Helvetica, sans-serif">Building 
        <select name="bldg">
          <%Set rst2 = Server.CreateObject("ADODB.recordset")
			sqlstr = "select bldgnum  as bldg, strt as bldgname from buildings"
   			rst2.Open sqlstr, cnn1, 0, 1, 1
			if not rst2.eof then
					do until rst2.eof
		%>
          <option value="<%=rst2("bldg")%>"><font face="Arial, Helvetica, sans-serif"><%=rst2("bldgname")%>,<%=rst2("bldg")%></font></option>
          <%
					rst2.movenext
					loop
					end if
					rst2.close
				%>
        </select>
        </font></td>
      <td width="52%" height="2">&nbsp;</td>
    </tr>
    <tr> 
      <td width="48%" height="2"> 
        <p><font face="Arial, Helvetica, sans-serif">Bill Year 
          <input type="text" name="bperiod" size="5" maxlength="4">
          Bill Period 
          <input type="text" name="byear" size="3" maxlength="2">
          </font> </p>
        <p> <font face="Arial, Helvetica, sans-serif"> 
          <input type="button" name="Button2" value="View Report" onClick="viewrpt(bldg.value,byear.value, bperiod.value)">
          </font></p>
      </td>
      <td width="52%" height="2"> 
        <div align="right"> <font face="Arial, Helvetica, sans-serif"> 
          <input type="button" name="Submit" value="Print Current View" onClick='javascript:document.frames.admin.focus();document.frames.admin.print()'>
          </font></div>
      </td>
    </tr>
  </table>
</form>
<p><IFRAME name="admin" width="100%" height="100%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME></p>
</body>
</html>