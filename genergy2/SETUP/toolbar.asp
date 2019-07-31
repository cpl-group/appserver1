<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, tid
pid = secureRequest("pid")
bldg = secureRequest("bldg")
tid = secureRequest("tid")

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(pid,bldg,"billing")

dim utility, lid
if trim(lid)<>"" then
	rst1.Open "SELECT * FROM tblleasesutilityprices WHERE leaseutilityid='"&lid&"'", cnn1
	if not rst1.EOF then
		utility = rst1("utility")
	end if
	rst1.close
end if

%>
<html>
<head>
<title>Portfolio View</title>
<script>
function loadPortfolio(pid)
{	document.location = "toolbar.asp?pid="+pid;
//parent.contentfrm.location = "groupView.asp?pid="+pid;
}
function portfolioEdit(){
  pid = document.forms[0].pid.value;
  document.location = "toolbar.asp?pid="+pid;
  parent.contentfrm.location = "portfolioedit.asp?pid="+pid;
}
function buildingSelect(pid,bldg)
{ 
	document.location = 'toolbar.asp?pid='+pid+'&bldg='+bldg;
//  parent.contentfrm.location = 'buildingedit.asp?pid='+pid+'&bldg='+bldg;
}
function buildingEdit(pid)
{	
  bldg = document.forms[0].building.value;
  parent.contentfrm.location = 'buildingedit.asp?pid='+pid+'&bldg='+bldg;
	document.location = 'toolbar.asp?pid='+pid+'&bldg='+bldg;
}
function groupView()
{	parent.contentfrm.location = 'groupView.asp?pid=<%=pid%>';
}
function tenantSelect(tid)
{	
  document.location.href = 'toolbar.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid='+tid;
//  parent.contentfrm.location = 'tenantedit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid='+tid;
}
function tenantEdit()
{	
  tid = document.forms[0].tenants.value;
  parent.contentfrm.location = 'tenantedit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid='+tid;
  document.location.href = 'toolbar.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid='+tid;
}
function leaseUtilityEdit()
{
  lid = document.forms[0].util.value;
	parent.contentfrm.location = 'leaseutilityedit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid='+lid;
}
function leaseUtilSelect(lid){
  parent.contentfrm.location.href = "contentfrm.asp?action=meteredit&pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid="+lid; 
}
function showMeters(lid){
  parent.contentfrm.location.href = "contentfrm.asp?action=showmeters&pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid="+lid; 
}

function loadbottomframes(){
  if (parent.contentfrm.length < 1) {
    parent.contentfrm.location = "contentfrm.asp";
  }
}

minimgOn = new Image();minimgOn.src = "images/btn-show_frame-000.gif";
minimgOff = new Image();minimgOff.src = "images/btn-hide_frame-000.gif";

function minframe(){
  if (parent.document.body.rows != "25,*") {
    parent.document.body.rows = "25,*";
    document.all.minimizerimg.src = minimgOn.src;
  } else {
    parent.document.body.rows = "100,*";
    document.all.minimizerimg.src = minimgOff.src;
  };
}

</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>
<body bgcolor="#eeeeee">
<FORM>
<table border=0 cellpadding="5" cellspacing="0" width="100%" bgcolor="#FFFFFF">
<tr>
  <td bgcolor="#000000"><span class="standardheader"><a href="index.asp" target="main" class="breadcrumb" style="text-decoration:none;"><img src="images/aro-left-000.gif" align="left" width="13" height="13" border="0">Utility Manager Setup</a></span></td>
  <td bgcolor="#000000" align="right"><a href="javascript:minframe();" id="minimizer" style="color:#33cc99;"><img name="minimizerimg" src="images/btn-hide_frame-000.gif" width="80" height="16" border="0"></a></td>
</tr>
</table>
<table border=0 cellpadding="2" cellspacing="0" width="100%">
<tr>
  <td height="29" width="50%" bgcolor="#eeeeee">
  <!-- begin portfolio pulldown -->
  <table border=0 cellpadding="1" cellspacing="0">
  <tr>
    <td><img src="images/setup_portfolio.gif" alt="set up portfolio" title="Set Up Portfolio" width="33" height="29" align="middle" border="0" onclick="portfolioEdit();"></td>
    <td>
    <select name="pid" onChange="loadPortfolio(this.value);">
<!--    [[select name="pid" onChange="loadPortfolio(this.value);loadbottomframes();"]]-->
    <option>Select Portfolio</option>
    <%
    rst1.open "SELECT * FROM portfolio ORDER BY name", getConnect(pid,bldg,"dbCore")
    do until rst1.eof
      %><option value="<%=rst1("id")%>"<%if trim(pid)=trim(rst1("id")) then Response.Write " SELECTED"%>><%=rst1("name")%></option><%
      rst1.MoveNext
    loop
    rst1.close
    %>
    </select>
    </td>
  </tr>
  </table>  
  <!-- end portfolio pulldown -->
  </td>
  <%
  if trim(bldg)<>"" then%>
  <td height="29" width="50%" bgcolor="#eeeeee">
  <!-- begin tenants pulldown -->
  <table border=0 cellpadding="1" cellspacing="0">
  <tr>
    <td><img src="images/setup_tenants.gif" alt="Set Up Accounts" title="Set Up Accounts" width="33" height="29" align="middle" border="0" onclick="tenantEdit();"></td>
    <td>
    <%
    rst1.Open "SELECT * FROM tblleases l WHERE bldgnum='"&bldg&"' ORDER BY billingname", cnn1
    if not rst1.EOF then%>
    <select name="tenants" onChange="tenantSelect(this.value);">
<!--
    [[select name="tenants" onChange="tenantSelect(this.value);loadbottomframes();"]]
-->
    <option value="">Select Account
    <%do until rst1.EOF%>
    <option value="<%=rst1("billingid")%>"<%if trim(tid)=trim(rst1("billingid")) then Response.Write " SELECTED"%>><%=rst1("billingname")%>
    <%rst1.movenext
    loop%>
    </select>
    <%
    else
      Response.Write "Click the icon at left to set up tenants for this building"
      Response.Write "<input type='hidden' name='tenants' value=''>"
    end if
    rst1.close%>
    </td>
  </tr>
  </table>  
  <!-- end tenants pulldown -->
  </td>
  <%
  end if
  %>
</tr>
<tr>
  <%
  if trim(pid)<>"" then%>
  <td height="29" bgcolor="#eeeeee">
  <!-- begin building pulldown -->
  <table border=0 cellpadding="1" cellspacing="0">
  <tr>
    <td><img src="images/setup_bldg.gif" alt="Set Up Building" title="Set Up Building" width="33" height="29" align="middle" border="0" onclick="buildingEdit('<%=pid%>');"></td>
    <td>
    <%
  '	Response.Write "<input type=""button"" value=""Add Building"" onclick=""buildingEdit(','');"">"
  '	Response.Write "<input type=""button"" value=""Manage Groups"" onclick=""groupView();"">"
    rst1.Open "SELECT * FROM buildings b WHERE portfolioid="&pid&" order by bldgname", cnn1
    if not rst1.EOF then%>
    <select name="building" onChange="buildingSelect('<%=pid%>',this.value);">
    <!--[[select name="building" onChange="buildingSelect('[[%=pid%]]',this.value);loadbottomframes();"]]-->
    <option value="">Select Building
    <%do until rst1.EOF%>
    <option value="<%=rst1("bldgnum")%>"<%if trim(bldg)=trim(rst1("bldgnum")) then Response.Write " SELECTED"%>><%=rst1("bldgname")%>
    <%rst1.movenext
    loop%>
    </select>
    <%
    else
      Response.Write "Click the icon at left to set up buildings for this portfolio"
      Response.Write "<input type='hidden' name='building' value=''>"
    end if
    rst1.close%>
    </td>
  </tr>
  </table>
  <!-- end building pulldown -->
  </td>
  <%
  end if
  %>

  <%if trim(bldg)<>"" then%>
  <td bgcolor="#eeeeee">
  <!-- begin utility pulldown -->
  <%
  if trim(tid)<>"" then%>
  <table border=0 cellpadding="1" cellspacing="0">
  <tr>
    <td><img src="images/setup_utils.gif" alt="Set Up Lease Utilities" title="Set Up Lease Utilities" width="33" height="29" align="middle" border="0" onclick="leaseUtilityEdit();"></td>
    <td>
    <%
  	rst1.Open "SELECT * FROM tblleasesutilityprices lup RIGHT JOIN tblutility tu ON lup.utility=tu.utilityid  WHERE billingid='"&tid&"'", cnn1
  	if not rst1.EOF then%>
    <select name="util" onChange="leaseUtilSelect(this.value);">
 		<%do until rst1.EOF%>
    <option value="<%=rst1("leaseutilityid")%>"<%if trim(utility)=trim(rst1("utilityid")) or trim(rst1("utilityid"))=2 then response.write " selected"%>><%=rst1("utilitydisplay")%></option>
    <%
    rst1.movenext
  loop%>
    </select>
    <script>
    //if (parent.contentfrm.length <= 1) { loadbottomframes(); }
    //leaseUtilSelect(document.forms[0].util.value);
    </script></td>
    <td><img src="images/btn-show_meters.gif" value="Show Meters" onclick="showMeters(document.forms[0].util.value);" alt="Show Meters" width="77" height="24" border="0">
  <%else
    Response.Write "Click the icon at left to set up lease utilities for this tenant"
    Response.Write "<input type='hidden' name='util' value=''>"
  end if
  rst1.close%>
    </td>
  </tr>
  </table>  
  <!-- end utility pulldown -->  
  <%end if%>
  </td>
  <% end if %>
</tr>
</table>

</FORM>
</body>
</html>
