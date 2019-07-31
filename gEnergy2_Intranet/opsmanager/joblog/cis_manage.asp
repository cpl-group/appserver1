<%option explicit%>
<!-- #include virtual="/genergy2/secure.inc" -->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<html>
<%
dim cnn, rst, strsql
set cnn = server.createobject("ADODB.connection")
set rst = server.createobject("ADODB.recordset")
cnn.open getConnect(0,0,"intranet")

dim cid, company,cName,cType,cTrade,cStatus, cAddress_street,cAddress_floor,cCity,cState,cZip,cPhone ,cFax_phone,date_established,cMaincontact, bAddress_street,bAddress_floor,bCity,bState,bZip,bMainContact, tcolor, cTitle, cContactID, email, contactmsg, customer

cid = secureRequest("cid")
company = secureRequest("company")
if company="" then company = "EM" 'rsm changed to EM
'company = "GY"

if trim(cid)<>"" then
  rst.Open "SELECT * FROM " & company & "_MASTER_ARM_CUSTOMER WHERE customer='"&cid&"'", cnn
  if not rst.EOF then
    cName   = rst("name")
    cType = rst("customer_type")
    cTrade  = rst("trade")
    cStatus = lcase(trim(rst("status")))
    Select Case cStatus
      case "active"
        tcolor = "#FFFFFF"
      case "inactive"
        tcolor = "#cccccc"
    end select 
    cAddress_street = rst("address_1")
    cAddress_floor = rst("address_2")
    cCity = rst("city")
    cState = rst("state")
    cZip = rst("zip_code")
    cPhone = rst("telephone")
    cFax_phone = rst("fax")
    date_established = rst("date_established")
    cMaincontact = rst("contact_1")
    bAddress_street = rst("billing_address_1")
    bAddress_floor = rst("billing_address_2")
    bCity = rst("billing_city")
    bState = rst("billing_state")
    bZip = rst("billing_zip_code")  
    bMaincontact= rst("billing_contact")
  end if
  rst.close
end if
%>
<title>Job Search</title>
<script language="JavaScript" type="text/javascript">
//<!--

function newcustomer() {
//	company = document.form1.company.value
	theURL = "cis_update.asp?mode=new&company=<%=trim(company)%>"
	openwin(theURL,600,350)
}

function opencontact(mode) 
{
//	company = document.form1.company.value
	theURL="updatecontact.asp?mode=" + mode + "&company=EM";
	openwin(theURL,500,450)
}

function customerdetail(cid) {
  theURL="cis_detail.asp?cid=" + cid + "&company=" + '<%=company%>'
  openwin(theURL,600,475)
}

function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}

function screencompany(company) {
    document.location.href="cis_manage.asp?company="+company	
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
  <td style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">&nbsp;<b>Manage Customers &amp; Contacts</b></td>
  <td align="right" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
  <input id="editjob" name="editjob" type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onclick="newcustomer();" value="New Customer">
    <input id="editjob" name="editjob" type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onclick="opencontact('new');" value="New Contact">
</tr>
<tr>
      <td style="border-top:1px solid #ffffff;"> 
        <select name="company" onchange="screencompany(this.value)" >
	
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
              </select>
  </td>
  <td align="right" style="border-top:1px solid #ffffff;"><a href="javascript:toggleHelp();" style="text-decoration:none;"><img src="/gEnergy2_Intranet/opsmanager/joblog/images/quick_help.gif" align="absmiddle" border="0">&nbsp;<b>Quick Help</b></a></td>
</tr>
<tr valign="top">
  <td colspan="2" height="255">
  <div id="quickhelptext" style="display:none;">
  <ul>
  <li>Click the radio button next to a company to show its customers. Customers and contacts are maintained separately for all entities.
  </ul>
  </div>
<% if trim(company) <>"AC" and trim(company) <> "" then %>
  <div id="customers" style="overflow:auto;width:100%;height:245px;border:1px solid #cccccc;">
  <table border=0 cellpadding="3" cellspacing="1" width="100%" bgcolor="#cccccc">
  <%
  rst.Open "SELECT distinct customer,name, status FROM " & company & "_MASTER_ARM_CUSTOMER order by name", cnn
    do until rst.eof
    cStatus = lcase(trim(rst("status")))
    Select Case cStatus
      case "active"
        tcolor = "#FFFFFF"
      case "inactive"
        tcolor = "#cccccc"
    end select 
     %>
   <tr bgcolor="<%=tcolor%>" onMouseOver="this.style.backgroundColor = 'lightgreen'" style="cursor:hand" onMouseOut="this.style.backgroundColor = '<%=tcolor%>'" onclick="customerdetail('<%=trim(rst("customer"))%>');"><td><%=left(trim(rst("name")),30) & " ("&rst("customer")&")"%></td></tr>
    <% rst.movenext
    loop
  rst.close
  %>
  </table>  
  </div>
<% end if %>
  </td>
</tr>
<tr>
  <td style="border-top:2px outset #ffffff;border-bottom:1px solid #cccccc;">
  <input id="editjob" name="editjob" type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onclick="newcustomer();" value="New Customer">
    <input id="editjob" name="editjob" type="button" class="standard" style="cursor:hand;background-color:#eeeeee;border:1px outset #ffffff;color:336699;" onclick="opencontact('new');" value="New Contact">
  </td>
  <td align="right" style="border-top:2px outset #ffffff;border-bottom:1px solid #cccccc;border-right:1px solid #cccccc;"><img src="/um/opslog/images/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
</tr>
</table>
</form>
</body>
</html>