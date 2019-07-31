<%option explicit%>
<!-- #include virtual="/genergy2/secure.inc" -->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<html>
<%
dim cnn, rst, strsql
set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open application("cnnstr_main")

dim company,pid

company = request("company")
pid = request("pid")

%>
<title>Job Search</title>
<script language="JavaScript" type="text/javascript">
//<!--
function screencompany(company) {
    document.location.href="cis_browse.asp?company="+company	
}
function loadportfolio(ckey,loadflag,sterm) {
    if(loadflag==1) {
      document.location.href="cis_browse.asp?company=<%=company%>&pid="+ckey	
	  parent.searchbar.form1.search.value=sterm
	  }
	else {
	  parent.searchbar.form1.custid.value=ckey
	  parent.searchbar.form1.search.value=sterm
	  parent.searchbar.clearDefault()
	  parent.searchbar.form1.submit()
	  }
}
//display quickhelp
var helpIsOn = 0;
function toggleHelp(){
  if (helpIsOn) { 
    document.all.quickhelptext.style.display='none';
    helpIsOn = 0;
   } else { 
    document.all.quickhelptext.style.display='inline';
    helpIsOn = 1;
   }
}

//visual feedback functions for img buttons
function buttonOver(obj,clr){
  if (arguments.length == 1) { clr = "#336699"; }
  obj.style.border = "1px solid " + clr;
}

function buttonDn(obj,clr){
  if (arguments.length == 1) { clr = "#000000"; }
  obj.style.border = "1px solid " + clr;
}

function buttonOut(obj,clr){
  if (arguments.length == 1) { clr = "#eeeeee"; }
  obj.style.border = "1px solid " + clr;
}
//-->
</script>
<link rel="Stylesheet" href="../../styles.css" type="text/css">		
</head>

<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#eeeeee">
<form name="form1">
<table border=0 cellpadding="3" cellspacing="0" width="100%" style="border-top:1px solid #cccccc;">
<tr>
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">&nbsp;<b>Browse Jobs By Customer</b></td>
  <td align="right" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
<tr>
  <td style="border-top:1px solid #ffffff;">
<select name="company" onchange="screencompany(this.value)">
		<%if trim(company) = "" then %>
		<option value="">Select Company</option>
		<% end if %>
                <%
        rst.Open "select * from companycodes where active = 1 and code <> 'AC' order by name", cnn
        if not rst.eof then
        do until rst.eof
        %>
		<option value="<%=rst("code")%>" <%if trim(company) = trim(rst("code")) then%> selected <%end if%>><font face="Arial, Helvetica, sans-serif"><%=rst("name")%></font></option>
                <%    
        rst.movenext
        loop
        end if
        rst.close%>
     </select>  </td>
  <td align="right" style="border-top:1px solid #ffffff;"><a href="javascript:toggleHelp();" style="text-decoration:none;"><img src="/gEnergy2_Intranet/opsmanager/joblog/images/quick_help.gif" align="absmiddle" border="0">&nbsp;<b>Quick Help</b></a></td>
</tr>
<tr valign="top">
  <td colspan="2" height="255">
  <div id="quickhelptext" style="display:none;">
  <ul>
  <li>Click the radio button next to a company to list portfolios and individual customers. Customers are maintained separately for Genergy and I-Lite.
  </ul>
  </div>
 <% if trim(company)<>"AC" and trim(company)<>"" then %>
 
  <div id="customers" style="overflow:auto;width:100%;height:245px;border:1px solid #cccccc;">
  <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
  <%
    if pid="" then 'portfolios and ungrouped customers, switch js params accordingly
	
		dim cnnSuper
		set cnnSuper = server.createobject("ADODB.connection")
		
		cnnSuper.open application("cnnstr_SuperMod")
		dim mainIp, someSQL
		mainIp = getIpFromCnnStr(application("cnnStr_main"))
		someSQL = "select portfolio,name,'0' as customer from portfolio union select 'xxx' as portfolio,name,customer from [" & mainIp & "].main.dbo." & company & "_master_arm_customer where PID='PID' order by portfolio,name "
		'response.write someSQL
		'response.end
		rst.Open  someSQL, cnnSuper
		do until rst.eof
      %>
   <tr bgcolor="#ffffff" onMouseOver="this.style.backgroundColor = 'lightgreen'" style="cursor:hand" onMouseOut="this.style.backgroundColor = 'white'" onclick=<%   
        if rst("portfolio")<>"xxx" then
	    %>"loadportfolio('<%=rst("portfolio")%>',1,'<%=replace(left(trim(rst("name")),30),"'","")%>')"><td><%=left(trim(rst("name")),30)%></td></tr>
		<%else%>
		"loadportfolio('<%=rst("customer")%>',0,'<%=replace(left(trim(rst("name")),30),"'","")%>')"><td><%=left(trim(rst("name")),30)%></td></tr>
		<%
		end if
        rst.movenext
      loop
      rst.close
    
	else 'grouped customer case, same params for all
	
      rst.Open  "select name,customer from " & company & "_master_arm_customer where PID='"&pid&"' order by name ", cnn
      do until rst.eof  %>
   <tr bgcolor="#ffffff" onMouseOver="this.style.backgroundColor = 'lightgreen'" style="cursor:hand" onMouseOut="this.style.backgroundColor = 'white'" onclick="loadportfolio('<%=rst("customer")%>',0,'<%=left(trim(rst("name")),30)%>')"><td><%=left(trim(rst("name")),30)%></td></tr>
<%      rst.movenext
      loop
      rst.close
	end if
  %>
  </table>  
  </div>
<% end if 'end radio selected case%>
  </td>
</tr>
<tr>
  <td colspan=2 align="right" style="border-top:2px outset #ffffff;border-bottom:1px solid #cccccc;border-right:1px solid #cccccc;"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
</table>
</form>
</body>
</html>
<%
function getIpFromCnnStr(cnnStr)
	dim cnn_str_tokens, equalsSignLoc
	'response.write("param:" & cnnStr)
	cnn_str_tokens = split(cnnStr, ";", -1, 0)
	cnnStr = cnn_str_tokens(1)
	equalsSignLoc = inStr(cnnStr,"=")
	getIpFromCnnStr = trim(right(cnnStr, len(cnnStr) - equalsSignLoc ))
end function
%>