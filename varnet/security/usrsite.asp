<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'Response.Redirect "http://www.genergyonline.com"
		else
			if  Session("ts") < 5 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				Response.Redirect "../main.asp"
			end if
		
		end if		
username=Request("username")
'Response.write(username)
index=0
%>
<html>

<head>
<!--#include file="../adovbs.inc" -->
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<script>
function fillup(name){
    document.location="usrdetail.asp?username="+name
	parent.frames.site.location="usrsite.asp?username="+name
}

function modify(index, bldg_name){
   //alert(choice)
   alert(bldg_name)
   temp="document.forms["+index+"].bldg_name.value"
   alert(temp)
}
</script>
<title>Tenant Selection</title>
</head>

<body bgcolor="#FFFFFF">
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open application("cnnstr_security")

rst1.open "SELECT * FROM tblCoreServices WHERE CSID in (SELECT csid FROM tbladdons)", cnn1
do until rst1.eof
	%><input type="button" name="action" value="<%=rst1("Label")%> Options" onclick="window.open('optionsList.asp?username=<%=username%>&csid=<%=rst1("csid")%>&label=<%=rst1("Label")%>', '<%=rst1("csid")%>', 'scrollbars=yes, width=250, height=400, toolbar=no');"><br><%
	rst1.movenext
loop
rst1.close
%>
<table border="0" width="100%" cellpadding="0" cellspacing="0" bordercolor="#000000" height="46" align="center">
<%
strsql = "SELECT * FROM clientsites where userid='"&username&"'"
rst1.Open strsql, cnn1, 0, 1, 1
flag=Request("flag")
	if not rst1.eof and flag = "" then
	col=0
    do until rst1.eof
	bldg=Trim(rst1("bldg_name"))
	if col mod 5=0 then
%>
  <tr> 
    <%
end if
%>
    <td> <font face="Arial, Helvetica, sans-serif"><a href="#<%=bldg%>"><%=rst1("bldg_name")%></a> 
      </font></td>
    <%
    rst1.movenext
	if col mod 5=0 then
%>
    <font face="Arial, Helvetica, sans-serif"></font>
    <%	
    end if
	col=col+1
	loop
	rst1.movefirst
%>
</table>
<br>
<div align="center"> 
  <%    
    do until rst1.eof
    %>
<form method="post" action="sitemodify.asp">
<input type="hidden" name="username" value="<%=username%>">
<input type="hidden" name="sitekey" value="<%=rst1("clientsitekey")%>">
    <table border="0" width="100%" cellpadding="0" cellspacing="0" bordercolor="#000000" height="46" align="center">
      <input type="hidden" name=count value="<%=index%>">
      <tr> 
        <td> <b><a name="<%=Trim(rst1("bldg_name"))%>"></a>Building Name</b> </td>
        <td> 
          <input type="text" name="bldg_name" value="<%=rst1("bldg_name")%>">
        </td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> Region </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="region" value="<%=rst1("region")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> Region Count </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="regioncount" value="<%=rst1("regioncount")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> Building ID </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="bldgid" value="<%=rst1("bldgid")%>">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> Contact Name </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="contact_name" value="<%=rst1("contact_name")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> Building Link 
          </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="bldgid_link" value="<%=rst1("bldgid_link")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PGI Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="pgi_link" value="<%=rst1("pgi_link")%>">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PGI Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="pgi_target" value="<%=rst1("pgi_target")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> ERI Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="eri_link" value="<%=rst1("eri_link")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> ERI Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="eri_target" value="<%=rst1("eri_target")%>">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> Meter Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="meter_link" value="<%=rst1("meter_link")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> Meter Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="meter_target" value="<%=rst1("meter_target")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> OVTHVAC link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="ovac_link" value="<%=rst1("ovthvac_link")%>">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> OVTHVAC Target 
          </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="ovthvac_target" value="<%=rst1("ovthvac_target")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> POW_AVA Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="pow_ava_link" value="<%=rst1("pow_ava_link")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> POW_AVA Target 
          </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="pow_ava_target" value="<%=rst1("pow_ava_target")%>">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PA Chart Link 
          </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="pa_chart_link" value="<%=rst1("pa_chart_link")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PA Chart Target 
          </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="pa_chart_target" value="<%=rst1("pa_chart_target")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> REV Prof Link 
          </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="rev_prof_link" value="<%=rst1("rev_prof_link")%>">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> REV Prof Target 
          </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="rev_prof_target" value="<%=rst1("rev_prof_target")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> IRI Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="iri_link" value="<%=rst1("iri_link")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> IRI Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="iri_target" value="<%=rst1("iri_target")%>">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> MEP Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="mep_link" value="<%=rst1("mep_link")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> MEP Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="mep_target" value="<%=rst1("mep_target")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PQ Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="pq_link" value="<%=rst1("pq_link")%>">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PQ Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="pq_target" value="<%=rst1("pq_target")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PLP Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="plp_link" value="<%=rst1("plp_link")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PLP Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="plp_target" value="<%=rst1("plp_target")%>">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> MSI Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="msi_link" value="<%=rst1("msi_link")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> MSI Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="msi_target" value="<%=rst1("msi_target")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> LMP Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="lmp_link" value="<%=rst1("lmp_link")%>">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> LMP Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="lmp_target" value="<%=rst1("lmp_target")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> CA Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="ca_link" value="<%=rst1("ca_link")%>">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> CA Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="ca_target" value="<%=rst1("ca_target")%>">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PGI </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <%
	  'response.write(rst1("pgi"))
	  'response.write(rst1("iri"))
	  'response.write(rst1("userid"))
	  if rst1("pgi")=true then
	  %>
          <input type="checkbox" name="pgi" checked value="1">
          <%
	  else
	  %>
          <input type="checkbox" name="pgi" value="1">
          <%
	  end if
	  %>
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> ERI </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <%
	  if rst1("eri")=true then
	  %>
          <input type="checkbox" name="eri" checked value="1">
          <%
	  else
	  %>
          <input type="checkbox" name="eri" value="1">
          <%
	  end if
	  %>
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> Meter </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <%
	  if rst1("meter")=true then
	  %>
          <input type="checkbox" name="meter" checked value="1">
          <% 
	  else
	  %>
          <input type="checkbox" name="meter" value="1">
          <%
	  end if
	  %>
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> OVTHVAC </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <%
	  if rst1("ovthvac")=true then
	  %>
          <input type="checkbox" name="ovthvac" checked value="1">
          <%
	  else
	  %>
          <input type="checkbox" name="ovthvac" value="1">
          <%
	  end if
	  %>
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> POW AVA </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <%
	  if rst1("pow_ava")=true then
	  %>
          <input type="checkbox" name="pow_ava" checked value="1">
          <%
	  else
	  %>
          <input type="checkbox" name="pow_ava" value="1">
          <%
	  end if
	  %>
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PA Chart </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <%
	  if rst1("pa_chart") then
	  %>
          <input type="checkbox" name="pa_chart" checked value="1">
          <%
	  else
	  %>
          <input type="checkbox" name="pa_chart" value="1">
          <%
	  end if
	  %>
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> REV Prof </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <%
	  if rst1("rev_prof")=true then
	  %>
          <input type="checkbox" name="rev_prof" checked value="1">
          <%
	  else
	  %>
          <input type="checkbox" name="rev_prof" value="1">
          <%
	  end if
	  %>
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> IRI </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <%
	  if rst1("iri")=true then
	  %>
          <input type="checkbox" name="iri" checked value="1">
          <%
	  else
	  %>
          <input type="checkbox" name="iri" value="1">
          <%
	  end if
	  %>
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> MEP </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <%
	  if rst1("mep")=true then
	  %>
          <input type="checkbox" name="mep" checked value="1">
          <%
	  else
	  %>
          <input type="checkbox" name="mep" value="1">
          <%
	  end if
	  %>
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PQ </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <%
	  if rst1("pq")=true then
	  %>
          <input type="checkbox" name="pq" checked value="1">
          <%
	  else
	  %>
          <input type="checkbox" name="pq" value="1">
          <%
	  end if
	  %>
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PLP </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <%
	  if rst1("plp")=true then
	  %>
          <input type="checkbox" name="plp" checked value="1">
          <%
	  else
	  %>
          <input type="checkbox" name="plp" value="1">
          <%
	  end if
	  %>
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> MSI </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <%
	  if rst1("msi")=true then
	  %>
          <input type="checkbox" name="msi" checked value="1">
          <%
	  else
	  %>
          <input type="checkbox" name="msi" value="1">
          <%
	  end if
	  %>
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> LMP </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <%
	  if rst1("lmp")=true then
	  %>
          <input type="checkbox" name="lmp" checked value="1">
          <%
	  else
	  %>
          <input type="checkbox" name="lmp" value="1">
          <%
	  end if
	  %>
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> CA </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <%
	  if rst1("ca")=true then
	  %>
          <input type="checkbox" name="ca" checked value="1">
          <%
	  else
	  %>
          <input type="checkbox" name="ca" value="1">
          <%
	  end if
	  %>
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"></font></td>
      </tr>
      <tr> 
        <td colspan="2"><font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="submit" name="choice" value="Save">
          <input type="submit" name="choice" value="Delete">
          <input type="button" name="Button" onclick="javascript:history.back()" value="Back">
          </font></td>
        <td colspan="2"><font face="Arial, Helvetica, sans-serif" size="1"></font></td>
        <td colspan="2"><font face="Arial, Helvetica, sans-serif" size="1"></font></td>
      </tr>
    </table>
</form>
  <hr>
  <font size="+4"><br>
  <%
	
rst1.movenext
index=index+1
loop
else
%>
<form method="post" action="sitemodify.asp">
<input type="hidden" name="username" value="<%=flag%>">
    <table border="0" width="100%" cellpadding="0" cellspacing="0" bordercolor="#000000" height="46" align="center">
      <input type="hidden" name=count value="<%=index%>">
      <tr> 
        <td> <font face="Arial, Helvetica, sans-serif"><b>Building Name</b> </font></td>
        <td> <font face="Arial, Helvetica, sans-serif"> 
          <input type="text" name="bldg_name">
          </font></td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> Region </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="region">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> Region Count </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="regioncount" >
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> Building ID </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="bldgid">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> Contact Name </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="contact_name">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> Building Link 
          </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="bldgid_link">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PGI Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="pgi_link">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PGI Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="pgi_target">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> ERI Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="eri_link">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> ERI Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="eri_target">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> Meter Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="meter_link">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> Meter Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="meter_target">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> OVTHVAC link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="ovac_link">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> OVTHVAC Target 
          </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="ovthvac_target">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> POW_AVA Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="pow_ava_link">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> POW_AVA Target 
          </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="pow_ava_target">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PA Chart Link 
          </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="pa_chart_link">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PA Chart Target 
          </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="pa_chart_target">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> REV Prof Link 
          </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="rev_prof_link">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> REV Prof Target 
          </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="rev_prof_target">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> IRI Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="iri_link">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> IRI Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="iri_target">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> MEP Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="mep_link">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> MEP Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="mep_target">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PQ Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="pq_link">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PQ Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="pq_target">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PLP Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="plp_link">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PLP Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="plp_target">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> MSI Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="msi_link">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> MSI Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="msi_target">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> LMP Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="lmp_link">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> LMP Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="lmp_target">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> CA Link </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="ca_link">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> CA Target </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="text" name="ca_target">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PGI </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <%
	  'response.write(rst1("pgi"))
	  'response.write(rst1("iri"))
	  'response.write(rst1("userid"))
	  %>
          <input type="checkbox" name="pgi" value="1">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> ERI </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="checkbox" name="eri" value="1">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> Meter </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="checkbox" name="meter" value="1">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> OVTHVAC </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="checkbox" name="ovthvac" value="1">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> POW AVA </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="checkbox" name="pow_ava" value="1">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PA Chart </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="checkbox" name="pa_chart" value="1">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> REV Prof </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="checkbox" name="rev_prof" value="1">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> IRI </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="checkbox" name="iri" value="1">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> MEP </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="checkbox" name="mep" value="1">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PQ </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="checkbox" name="pq" value="1">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> PLP </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="checkbox" name="plp" value="1">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> MSI </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="checkbox" name="msi" value="1">
          </font></td>
      </tr>
      <tr> 
        <td><font face="Arial, Helvetica, sans-serif" size="1"> LMP </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="checkbox" name="lmp" value="1">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"> CA </font></td>
        <td> <font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="checkbox" name="ca" value="1">
          </font></td>
        <td><font face="Arial, Helvetica, sans-serif" size="1"></font></td>
      </tr>
      <tr> 
        <td colspan="2"><font face="Arial, Helvetica, sans-serif" size="1"> 
          <input type="submit" name="choice" value="Add">
          <input type="button" name="Button" onclick="javascript:history.back()" value="Back">
          </font></td>
        <td colspan="2"><font face="Arial, Helvetica, sans-serif" size="1"></font></td>
        <td colspan="2"><font face="Arial, Helvetica, sans-serif" size="1"></font></td>
      </tr>
    </table>
</form>
<%
end if
rst1.close
%>
  </font> </div>
</body>

</html>
