<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
pid 	= request("pid")
bldg 	= request("bldg")
utility = "electricity"
rpt 	= trim(request("rpt"))
NameStart = Len(rpt) + 4 

%>
<html>
<head>
<title><%=rptheader%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">		

<script>
function loadreport(bldg){
var graphtype,storedproc,storedproc2,by,b2,bldgname,rpt

storedproc  = document.runproc.storedproc.value
storedproc2  = document.runproc.storedproc2.value
by1 		= document.runproc.by1.value
by2 		= document.runproc.by2.value
bldgname 	= document.runproc.bldgname.value
rpt 		= document.runproc.rpt.value
vardollar 	= document.runproc.vardollar.value
varpercent 	= document.runproc.varpercent.value
if (document.runproc.graphtype[0].checked){
	graphtype = 6;
}else{
	graphtype = 1;
}


if (document.runproc.detailview.checked==false){
	document.all.secondset.style.display='none';
	document.frames.details.location.href="details.asp?storedproc=" + storedproc+ "&by1=" + by1 +"&by2=" + by2 + "&bldgname=" + bldgname +"&bldg=" + bldg+"&rpt=" + rpt;
	document.frames.info.location.href="charts.asp?storedproc=" + storedproc+ "&by1=" + by1 +"&by2=" + by2 + "&bldgname=" + bldgname +"&bldg=" + bldg + "&graphtype=" + graphtype +"&rpt=" + rpt;
}else{
	if (by2 == -1){
		alert("Please Provide A Comparison Year")
	} else {
		document.all.secondset.style.display='inline';
		document.frames.details.location.href="details.asp?storedproc=" + storedproc+ "&by1=" + by1 +"&by2=" + by2 + "&bldgname=" + bldgname +"&bldg=" + bldg+"&rpt=" + rpt;
		document.frames.info.location.href="charts.asp?storedproc=" + storedproc+ "&by1=" + by1 +"&by2=" + by2 + "&bldgname=" + bldgname +"&bldg=" + bldg + "&graphtype=" + graphtype +"&rpt=" + rpt;
		if (document.runproc.applyvariant.checked){
		document.frames.details2.location.href="details.asp?storedproc=" + storedproc2+ "&by1=" + by1 +"&by2=" + by2 + "&bldgname=" + bldgname +"&bldg=" + bldg+"&rpt=" + rpt + "&vardollar=" + vardollar + "&varpercent=" + varpercent+"&applyvariant=true";
		document.frames.info2.location.href="charts.asp?storedproc=" + storedproc2+ "&by1=" + by1 +"&by2=" + by2 + "&bldgname=" + bldgname +"&bldg=" + bldg + "&graphtype=" + graphtype +"&rpt=" + rpt + "&vardollar=" + vardollar + "&varpercent=" + varpercent+"&applyvariant=true";
		} else {
		document.frames.details2.location.href="details.asp?storedproc=" + storedproc2+ "&by1=" + by1 +"&by2=" + by2 + "&bldgname=" + bldgname +"&bldg=" + bldg+"&rpt=" + rpt;
		document.frames.info2.location.href="charts.asp?storedproc=" + storedproc2+ "&by1=" + by1 +"&by2=" + by2 + "&bldgname=" + bldgname +"&bldg=" + bldg + "&graphtype=" + graphtype +"&rpt=" + rpt;
		
		}
	}
}



}

</script>

</head>
<%

Set cnn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.recordset")
cnn.Open application("cnnstr_genergy1")
sqlstr = "select strt as bldgname from buildings where bldgnum = '"&bldg&"'"

rs.Open sqlstr, cnn,0

if not rs.eof then
	bldgname = rs("bldgname")
end if
rs.close

sqlstr = "select name from sysobjects where name like 'g1_"&rpt&"%' or name like 'g1_" & pid &"%'"

rs.Open sqlstr, cnn,0

%>
<body bgcolor="#eeeeee" text="#000000">
	
<table width="100%" border="0" cellpadding="3" cellspacing="0">
  <tr> 
    <td width="52%" bgcolor="#6699cc"><span class="standardheader">Building <%=rpt%> Analysis: 
      <%=bldgname%></span></td>
    <td width="48%" bgcolor="#6699cc"><div align="right">
        <input name="button" type="button" onclick="if(document.all['preferences'].style.display=='none'){document.all['preferences'].style.display='inline';}else{document.all['preferences'].style.display='none';}" value="Preferences">
      </div></td>
  </tr>
</table>
  <form method="POST" name="runproc">
	
  <table width="100%" border="0" cellpadding="3" cellspacing="0">
    <tr bgcolor="#eeeeee"> 
      <td valign="top" style="border-bottom:1px solid #cccccc;"> Select your Cost 
        Report and Comparison Years then click View Report</td>
    </tr>
    <tr> 
      <td style="border-top:1px solid #ffffff;"> <select name="storedproc">
          <%
	  	if not rs.eof then 
			while not rs.eof
	  %>
          <option value="<%=rs("name")%>"><%=Mid(replace(rs("name"), "_", " "), NameStart)%></option>
          <%
	  		rs.movenext
			wend
		end if
		rs.close
	  %>
        </select> 
        <% 
	sqlstr = "SELECT DISTINCT billyear as billyear FROM BillYrPeriod WHERE bldgnum = '"&bldg&"' AND utility = '"&utility&"' ORDER BY billyear DESC"
	rs.Open sqlstr, cnn,0

	%>
        <select name="by1">
          <%
	  	if not rs.eof then 
			while not rs.eof
	  %>
          <option value="<%=rs("billyear")%>"><%=rs("billyear")%></option>
          <%
	  		rs.movenext
			wend
		end if
		rs.movefirst
	  %>
        </select> <select name="by2">
          <option value="-1" selected>NONE</option>
          <%
	  	if not rs.eof then 
			while not rs.eof
	  %>
          <option value="<%=rs("billyear")%>"><%=rs("billyear")%></option>
          <%
	  		rs.movenext
			wend
		end if
		rs.close
	  %>
        </select>
        <input type="hidden" name="rpt" value="<%=rpt%>"> <input type="hidden" name="bldgname" value="<%=bldgname%>"> 
        <input type="button" name="View Report" value="View Report" onClick="loadreport('<%=bldg%>')"> 
      </td>
    </tr>
    <tr>
      <td style="border-top:1px solid #ffffff;">  <div id="preferences" style="display:none;">
  <br>
          <b>Building <%=rpt%> Analysis Preferences</b> 
          <table width="100%" height="94" border=0 cellpadding="3" cellspacing="0">
            <tr valign="top"> 
              <td width="538"> <table width="100%" border=0 cellpadding="3" cellspacing="0">
                  <tr> 
                    <td><input type="checkbox" name="detailview"></td>
                    <td>Compare to 
                      <select name="storedproc2">
                        <%
		  sqlstr = "select name from sysobjects where name like 'g1_"&rpt&"%' or name like 'g1_" & pid &"%'"
		  rs.Open sqlstr, cnn,0

	  	if not rs.eof then 
			while not rs.eof
	  %>
                        <option value="<%=rs("name")%>"><%=Mid(replace(rs("name"), "_", " "), NameStart)%></option>
                        <%
	  		rs.movenext
			wend
		end if
		rs.close
	  %>
                      </select></td>
                  </tr>
                  <tr>
                    <td valign="top"><input type="checkbox" name="applyvariant" onclick="document.runproc.detailview.checked = true;"></td>
                    <td><table width="202" border=0 cellpadding="3" cellspacing="0">
                        <tr> 
                          <td width="115"> <div align="left">Apply Variant $ </div></td>
                          <td width="75"> <div align="left"> 
                              <input name="vardollar" type="text" size="5" onclick="document.runproc.detailview.checked = true;document.runproc.applyvariant.checked = true;" >
                            </div></td>
                        </tr>
                        <tr> 
                          <td><div align="left">Apply Variant % </div></td>
                          <td> <div align="left"> 
                              <input name="varpercent" type="text" size="5" onclick="document.runproc.detailview.checked= true;document.runproc.applyvariant.checked = true;">
                            </div></td>
                        </tr>
                      </table></td>
                  </tr>
                </table></td>
              <td width="182"> <table border=0 cellpadding="3" cellspacing="0">
                  <tr> 
                    <td><input name="graphtype" type="radio" value="6" checked></td>
                    <td> Show Bar Graph </td>
                  </tr>
                  <tr> 
                    <td><input type="radio" name="graphtype" value="1"></td>
                    <td>Show Line Graph </td>
                  </tr>
                </table></td>
              <td width="202">&nbsp;</td>
            </tr>
            <tr valign="top">
              <td>Note: Variants only apply to your &quot;Compare to&quot; graph</td>
              <td>&nbsp;</td>
              <td>&nbsp;</td>
            </tr>
          </table>
        </div>
</td>
    </tr>
    <tr> 
      <td width="21%" style="border-top:1px solid #ffffff;">&nbsp; </td>
    </tr>
  </table>
  </form>
<div align="center">
  <IFRAME name="info" width="600" height="300" src="/null.htm" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0 align=center> 
  </IFRAME><br>
   <IFRAME id="detailframe" name="details" width="600" height="115" src="/null.htm" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0 align=center> 
  </IFRAME>
</div>
<br>
<div id ="secondset" align="center" style="display:none;">
  <IFRAME name="info2" width="600" height="300" src="/null.htm" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0 align=center> 
  </IFRAME><br>
   <IFRAME id="detailframe2" name="details2" width="600" height="115" src="/null.htm" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0 align=center> 
  </IFRAME>
</div>


</body>
</html>
