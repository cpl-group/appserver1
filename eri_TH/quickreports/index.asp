<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'N.Ambo added 5/19/2008; When 'Enter Historical Data' is clicked direct to the data entry page.
'This data entry page had been custom made for a particular client to aid in running cost analysis reports

pid 	= request("pid")
bldg 	= request("bldg")
utility = request("utility")
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
var graphtype,storedproc,storedproc2,by,b2,bldgname,rpt,budget

storedproc  = document.runproc.storedproc.value
storedproc2  = document.runproc.storedproc2.value
by1 		= document.runproc.by1.value
by2 		= document.runproc.by2.value
bldgname 	= document.runproc.bldgname.value
rpt 		= document.runproc.rpt.value
varpercent 	= document.runproc.varpercent.value
<% if trim(lcase(rpt)) = "cost"then%>
vardollar 	= document.runproc.vardollar.value
budget		= document.runproc.budget.value
<% end if%>
if (document.runproc.graphtype[0].checked){
	graphtype = 6;
}else{
	graphtype = 1;
}


if (document.runproc.detailview.checked==false){
	document.all.secondset.style.display='none';
	document.frames.details.location.href="details.asp?storedproc=" + storedproc+ "&by1=" + by1 +"&by2=" + by2 + "&bldgname=" + bldgname +"&bldg=" + bldg+"&rpt=" + rpt + "&budget="+budget;
	document.frames.info.location.href="charts.asp?storedproc=" + storedproc+ "&by1=" + by1 +"&by2=" + by2 + "&bldgname=" + bldgname +"&bldg=" + bldg + "&graphtype=" + graphtype +"&rpt=" + rpt + "&budget="+budget;
}else{
	if (by2 == -1){
		alert("Please Provide A Comparison Year")
	} else {
		document.all.secondset.style.display='inline';
		document.frames.details.location.href="details.asp?storedproc=" + storedproc+ "&by1=" + by1 +"&by2=" + by2 + "&bldgname=" + bldgname +"&bldg=" + bldg+"&rpt=" + rpt + "&budget="+budget;
		document.frames.info.location.href="charts.asp?storedproc=" + storedproc+ "&by1=" + by1 +"&by2=" + by2 + "&bldgname=" + bldgname +"&bldg=" + bldg + "&graphtype=" + graphtype +"&rpt=" + rpt + "&budget="+budget;
		if (document.runproc.applyvariant.checked){
		document.frames.details2.location.href="details.asp?storedproc=" + storedproc2+ "&by1=" + by1 +"&by2=" + by2 + "&bldgname=" + bldgname +"&bldg=" + bldg+"&rpt=" + rpt + "&vardollar=" + vardollar + "&varpercent=" + varpercent+"&applyvariant=true";
		document.frames.info2.location.href="charts.asp?storedproc=" + storedproc2+ "&by1=" + by1 +"&by2=" + by2 + "&bldgname=" + bldgname +"&bldg=" + bldg + "&graphtype=" + graphtype +"&rpt=" + rpt + "&vardollar=" + vardollar + "&varpercent=" + varpercent+"&applyvariant=true";
		} else {
		document.frames.details2.location.href="details.asp?storedproc=" + storedproc2+ "&by1=" + by1 +"&by2=" + by2 + "&bldgname=" + bldgname +"&bldg=" + bldg+"&rpt=" + rpt + "&budget="+budget;
		document.frames.info2.location.href="charts.asp?storedproc=" + storedproc2+ "&by1=" + by1 +"&by2=" + by2 + "&bldgname=" + bldgname +"&bldg=" + bldg + "&graphtype=" + graphtype +"&rpt=" + rpt + "&budget="+budget;
		
		}
	}
}



}
function updatebudget(budget,bldg){
	
	document.all.secondset.style.display='none';
	document.frames.info.location.href = 'savebudget.asp?budget='+budget+'&bldg='+bldg
}
function LoadHistoricData(){
	var frm = document.forms['runproc'];	
	document.all.secondset.style.display='none';
	document.location.href = 'historicaldataentry.asp?pid='+frm.pid.value+'&bldgNum='+frm.bldg.value;
}

</script>

</head>
<body bgcolor="#eeeeee" text="#000000">
<%
Set cnn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.recordset")

if trim(pid) <> "" and trim(bldg)="" then 
	cnn.open getMainConnect(pid)
else
	if trim(bldg)<>"" then 
	  cnn.Open getLocalConnect(bldg)
	  sqlstr = "select strt as bldgname from buildings b where bldgnum = '"&bldg&"'"
	
	  rs.Open sqlstr, cnn,0
	  
	  if not rs.eof then
		bldgname = rs("bldgname")
	  end if
	  rs.close
	end if
end if
%>
	
<table width="100%" border="0" cellpadding="3" cellspacing="0">
  <tr> 
    <td width="52%" bgcolor="#6699cc"><span class="standardheader">Building <%=rpt%> Analysis: 
      <%=bldgname%></span></td>
      <%if request("de") = 1 then 'N.Ambo 5/19/2008 - if de=1 then place a 'Data Entry' button at the top%>
      <td width="48%" bgcolor="#6699cc"><div align="right">
        <input name="button" type="button" onclick="LoadHistoricData()" value="Enter Historical Data" >
      </div></td>
      <%end if%>
    <td width="48%" bgcolor="#6699cc"><div align="right">
        <input name="button" type="button" onclick="if(document.all['preferences'].style.display=='none'){document.all['preferences'].style.display='inline';}else{document.all['preferences'].style.display='none';}" value="Preferences">
      </div></td>
  </tr>
</table>
  <form method="get" name="runproc" action="index.asp">
	
  <table width="100%" border="0" cellpadding="3" cellspacing="0">
    <tr bgcolor="#eeeeee"> 
      <td valign="top" style="border-bottom:1px solid #cccccc;"> Select your Cost 
        Report and Comparison Years then click View Report</td>
    </tr>
    <tr> 
      <td style="border-top:1px solid #ffffff;">
      <select name="storedproc">
	  <%
	  sqlstr = "select  distinct storedproc, label from customCA_query s where rpttype='"&trim(rpt)&"' and pid=0 or pid = "& pid & " order by label"
	  rs.Open sqlstr, cnn,0
	  
	  while not rs.eof
	  %>
            <option value="[<%=rs("storedproc")%>]|<%=replace(rs("label"), "_", " ")%>"><%=replace(rs("label"), "_", " ")%></option>
      <%
	  rs.movenext
			wend
      rs.close
	  %>
      </select> 
      <%sqlstr = "SELECT utilityid, utilitydisplay FROM tblutility WHERE utilityid=2 ORDER BY utilitydisplay"
      rs.Open sqlstr, getConnect(0,0,"dbCore"),0%>
      <!-- <select name="utility" onchange="submit()"> -->
      <%while not rs.eof%>
            <!-- <option value="<%=rs("utilityid")%>"<%if cint(rs("utilityid"))=cint(utility) then response.write " SELECTED"%>> --><%=rs("utilitydisplay")%><!-- </option> -->
      <%rs.movenext
			wend
      rs.close%>
      <!-- </select>  -->
      <%sqlstr = "SELECT DISTINCT billyear as billyear FROM BillYrPeriod WHERE bldgnum = '"&bldg&"' AND utility = '"&utility&"' ORDER BY billyear DESC"
      rs.Open sqlstr, cnn,0%>
      <select name="by1">
      <%
      if not rs.eof then 
      while not rs.eof%>
            <option value="<%=rs("billyear")%>"><%=rs("billyear")%></option>
      <%rs.movenext
			wend
		rs.movefirst
		end if
	  %>
        </select> <select name="by2">
          <option value="" selected>NONE</option>
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
              <td width="366"> <table width="271" border=0 align="left" cellpadding="3" cellspacing="0">
                  <tr>
                    <td colspan=3>Comparisons&nbsp;</td>
                  </tr>
                  <tr> 
                    <td><input type="checkbox" name="detailview"></td>
                    <td>Compare to</td>
                    <td><div align="right"> 
                        <select name="storedproc2">
                          <%
	  sqlstr = "select  distinct storedproc, label from customCA_query s where rpttype='"&trim(rpt)&"' and pid=0 or pid = "& pid & " order by label"
	  rs.Open sqlstr, cnn,0
	  
	  while not rs.eof
	  %>
                          <option value="[<%=rs("storedproc")%>]|<%=replace(rs("label"), "_", " ")%>"><%=replace(rs("label"), "_", " ")%></option>
                          <%
	  rs.movenext
			wend
      rs.close
	  %>
                        </select>
                      </div></td>
                  </tr>
                  <tr> 
                    <td width="20"><input type="checkbox" name="applyvariant" onClick="document.runproc.detailview.checked = true;"></td>
                    <td width="139" align="left">Apply Variant $</td>
                    <td width="146" align="right"><input name="vardollar" type="text" size="5" onclick="document.runproc.detailview.checked = true;document.runproc.applyvariant.checked = true;" ></td>
                  </tr>

                  <tr> 
                    <td>&nbsp;</td>
                    <td><div align="left">Apply Variant % </div></td>
                    <td> <div align="right"> 
                        <input name="varpercent" type="text" size="5" onclick="document.runproc.detailview.checked= true;document.runproc.applyvariant.checked = true;">
                      </div></td>
                  </tr>
                </table></td>
              <td width="225">
			  <table width="100%" border=0 cellpadding="3" cellspacing="0">
                  <tr> 
                    <td>Annual Energy Budget</td>
                  </tr>
                  <tr> 
				  <%
				  sqlstr = "select top 1 annualamt from BudgetsByBuilding where bldgid = '" &trim(bldg)& "' order by billyear desc, id desc"
				  rs.Open sqlstr, cnn,0
		
					if not rs.eof then 
							budgetamt = rs("annualamt")
					else
							budgetamt = 0
					end if
					rs.close
				  %>
				    <td width="156">$
<input name="budget" type="text" size="20" value="<%=budgetamt%>"></td>
                  </tr>
                  <tr> 
                    <td><span style="margin:9px"><input type="button" name="budgetupdt" value="Update Budget" onclick="updatebudget(budget.value,bldg.value)"></span></td>
                  </tr>
                </table> 
				</td>
              <td width="362"><table width="100%" border=0 cellpadding="3" cellspacing="0">
                  <tr> 
                    <td colspan=2>Graph Type</td>
                  </tr>
                  <tr> 
                    <td><input name="graphtype" type="radio" value="6" checked></td>
                    <td> Show Bar Graph </td>
                  </tr>
                  <tr> 
                    <td><input type="radio" name="graphtype" value="1"></td>
                    <td>Show Line Graph </td>
                  </tr>
                </table></td>
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
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
  </form>
<div align="center">
  <IFRAME name="info" width="600" height="300" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0 align=center> 
  </IFRAME><br>
   <IFRAME id="detailframe" name="details" width="90%" height="150" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0 align=center> 
  </IFRAME>
</div>
<br>
<div id ="secondset" align="center" style="display:none;">
  <IFRAME name="info2" width="600" height="300" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0 align=center> 
  </IFRAME><br>
   <IFRAME id="detailframe2" name="details2" width="600" height="150" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" frameborder=0 border=0 align=center> 
  </IFRAME>
</div>
</body>
</html>
