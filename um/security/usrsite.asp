<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'Response.Redirect "http://www.genergyonline.com"
		end if		
username=Request("username")
'Response.write(username)
index=0
%>
<html>

<head>
<!--#include file="../adovbs.inc" -->
<script language="JavaScript" type="text/javascript">
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
</script>
<title>Tenant Selection</title>
<script language="JavaScript" type="text/javascript">
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
</script>
</head>

<body bgcolor="#eeeeee" class="innerbody">
<%

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"dbCore")
%>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699cc">
  <td><span class="standardheader"><%=username%>'s Buildings</span></td>
</tr>
<tr bgcolor="#dddddd">
  <td style="border-bottom:1px solid #cccccc;">
  <%
  rst1.open "SELECT * FROM tblCoreServices WHERE CSID in (SELECT csid FROM tbladdons)", cnn1
  do until rst1.eof
    %><input type="button" name="action" value="<%=rst1("Label")%> Options" onclick="window.open('optionsList.asp?username=<%=username%>&csid=<%=rst1("csid")%>&label=<%=rst1("Label")%>', '<%=rst1("csid")%>', 'scrollbars=yes, width=250, height=400, toolbar=no');" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;"> <%
    rst1.movenext
  loop
  rst1.close
  %>
  </td>
</tr>
<tr>
  <td style="border-top:1px solid #ffffff;">
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
  <%
  strsql = "SELECT * FROM clientsites where userid='"&username&"'"
  rst1.Open strsql, cnn1, 0, 1, 1
  flag=Request("flag")
    if not rst1.eof and flag = "" then
    col=0
  %>
  
    <%
      do until rst1.eof
    bldg=Trim(rst1("bldg_name"))
    if col mod 5=0 then
  %>
    <tr> 
      <%
  end if
  %>
      <td><img src="/genergy2/SETUP/images/aro-rt.gif" align="absmiddle" border="0">&nbsp;<a href="#<%=bldg%>"><%=rst1("bldg_name")%></a></td>
      <%
      rst1.movenext
    if col mod 5=0 then
  %>
      
      <%	
      end if
    col=col+1
    loop
    rst1.movefirst
  %>
  </table>
  
  </td>
</tr>
<tr>
  <td style="border-bottom:1px solid #cccccc;" height="8">&nbsp;</td>
</tr>
</table>
  <%    
    do until rst1.eof
    %>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<form method="post" action="sitemodify.asp">
<input type="hidden" name="username" value="<%=username%>">
<input type="hidden" name="sitekey" value="<%=rst1("clientsitekey")%>">
<input type="hidden" name=count value="<%=index%>">
<tr bgcolor="#dddddd">
  <td colspan="6"><a name="<%=Trim(rst1("bldg_name"))%>"></a><b><%=rst1("bldg_name")%></b></td>
</tr>
      <tr> 
        <td> Building Name </td>
        <td> 
          <input type="text" name="bldg_name" value="<%=rst1("bldg_name")%>">
        </td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr> 
        <td> Region </td>
        <td>  
          <input type="text" name="region" value="<%=rst1("region")%>">
          </td>
        <td> Region Count </td>
        <td>  
          <input type="text" name="regioncount" value="<%=rst1("regioncount")%>">
          </td>
        <td> Building ID </td>
        <td>  
          <input type="text" name="bldgid" value="<%=rst1("bldgid")%>">
          </td>
      </tr>
      <tr> 
        <td> Contact Name </td>
        <td>  
          <input type="text" name="contact_name" value="<%=rst1("contact_name")%>">
          </td>
        <td> Building Link 
          </td>
        <td>  
          <input type="text" name="bldgid_link" value="<%=rst1("bldgid_link")%>">
          </td>
        <td> PGI Link </td>
        <td>  
          <input type="text" name="pgi_link" value="<%=rst1("pgi_link")%>">
          </td>
      </tr>
      <tr> 
        <td> PGI Target </td>
        <td>  
          <input type="text" name="pgi_target" value="<%=rst1("pgi_target")%>">
          </td>
        <td> ERI Link </td>
        <td>  
          <input type="text" name="eri_link" value="<%=rst1("eri_link")%>">
          </td>
        <td> ERI Target </td>
        <td>  
          <input type="text" name="eri_target" value="<%=rst1("eri_target")%>">
          </td>
      </tr>
      <tr> 
        <td> Meter Link </td>
        <td>  
          <input type="text" name="meter_link" value="<%=rst1("meter_link")%>">
          </td>
        <td> Meter Target </td>
        <td>  
          <input type="text" name="meter_target" value="<%=rst1("meter_target")%>">
          </td>
        <td> OVTHVAC link </td>
        <td>  
          <input type="text" name="ovthvac_link" value="<%=rst1("ovthvac_link")%>">
          </td>
      </tr>
      <tr> 
        <td> OVTHVAC Target 
          </td>
        <td>  
          <input type="text" name="ovthvac_target" value="<%=rst1("ovthvac_target")%>">
          </td>
        <td> POW_AVA Link </td>
        <td>  
          <input type="text" name="pow_ava_link" value="<%=rst1("pow_ava_link")%>">
          </td>
        <td> POW_AVA Target 
          </td>
        <td>  
          <input type="text" name="pow_ava_target" value="<%=rst1("pow_ava_target")%>">
          </td>
      </tr>
      <tr> 
        <td> PA Chart Link 
          </td>
        <td>  
          <input type="text" name="pa_chart_link" value="<%=rst1("pa_chart_link")%>">
          </td>
        <td> PA Chart Target 
          </td>
        <td>  
          <input type="text" name="pa_chart_target" value="<%=rst1("pa_chart_target")%>">
          </td>
        <td> REV Prof Link 
          </td>
        <td>  
          <input type="text" name="rev_prof_link" value="<%=rst1("rev_prof_link")%>">
          </td>
      </tr>
      <tr> 
        <td> REV Prof Target 
          </td>
        <td>  
          <input type="text" name="rev_prof_target" value="<%=rst1("rev_prof_target")%>">
          </td>
        <td> IRI Link </td>
        <td>  
          <input type="text" name="iri_link" value="<%=rst1("iri_link")%>">
          </td>
        <td> IRI Target </td>
        <td>  
          <input type="text" name="iri_target" value="<%=rst1("iri_target")%>">
          </td>
      </tr>
      <tr> 
        <td> MEP Link </td>
        <td>  
          <input type="text" name="mep_link" value="<%=rst1("mep_link")%>">
          </td>
        <td> MEP Target </td>
        <td>  
          <input type="text" name="mep_target" value="<%=rst1("mep_target")%>">
          </td>
        <td> PQ Link </td>
        <td>  
          <input type="text" name="pq_link" value="<%=rst1("pq_link")%>">
          </td>
      </tr>
      <tr> 
        <td> PQ Target </td>
        <td>  
          <input type="text" name="pq_target" value="<%=rst1("pq_target")%>">
          </td>
        <td> PLP Link </td>
        <td>  
          <input type="text" name="plp_link" value="<%=rst1("plp_link")%>">
          </td>
        <td> PLP Target </td>
        <td>  
          <input type="text" name="plp_target" value="<%=rst1("plp_target")%>">
          </td>
      </tr>
      <tr> 
        <td> MSI Link </td>
        <td>  
          <input type="text" name="msi_link" value="<%=rst1("msi_link")%>">
          </td>
        <td> MSI Target </td>
        <td>  
          <input type="text" name="msi_target" value="<%=rst1("msi_target")%>">
          </td>
        <td> LMP Link </td>
        <td>  
          <input type="text" name="lmp_link" value="<%=rst1("lmp_link")%>">
          </td>
      </tr>
      <tr> 
        <td> LMP Target </td>
        <td>  
          <input type="text" name="lmp_target" value="<%=rst1("lmp_target")%>">
          </td>
        <td> CA Link </td>
        <td>  
          <input type="text" name="ca_link" value="<%=rst1("ca_link")%>">
          </td>
        <td> CA Target </td>
        <td>  
          <input type="text" name="ca_target" value="<%=rst1("ca_target")%>">
          </td>
      </tr>
      <tr> 
        <td> PGI </td>
        <td>  
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
          </td>
        <td> ERI </td>
        <td>  
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
          </td>
        <td> Meter </td>
        <td>  
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
          </td>
      </tr>
      <tr> 
        <td> OVTHVAC </td>
        <td>  
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
          </td>
        <td> POW AVA </td>
        <td>  
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
          </td>
        <td> PA Chart </td>
        <td>  
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
          </td>
      </tr>
      <tr> 
        <td> REV Prof </td>
        <td>  
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
          </td>
        <td> IRI </td>
        <td>  
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
          </td>
        <td> MEP </td>
        <td>  
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
          </td>
      </tr>
      <tr> 
        <td> PQ </td>
        <td>  
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
          </td>
        <td> PLP </td>
        <td>  
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
          </td>
        <td> MSI </td>
        <td>  
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
          </td>
      </tr>
      <tr> 
        <td> LMP </td>
        <td>  
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
          </td>
        <td> CA </td>
        <td>  
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
          </td>
        <td></td>
      </tr>
      <tr> 
        <td colspan="2"> 
          <input type="submit" name="choice" value="Save" style="border:1px outset #ddffdd;background-color:ccf3cc;">
          <input type="submit" name="choice" value="Delete" style="border:1px outset #ddffdd;background-color:ccf3cc;">
          </td>
        <td colspan="2"></td>
        <td colspan="2" align="right"><img src="/images/intranet/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
      </tr>
</form>
    </table>
  <br>
  <%
	
rst1.movenext
index=index+1
loop
else
%>
<form method="post" action="sitemodify.asp">
<input type="hidden" name="username" value="<%=flag%>">
<table border=0 cellpadding="3" cellspacing="0" width="100%">
      <input type="hidden" name=count value="<%=index%>">
      <tr> 
        <td>Building Name</td>
        <td>  
          <input type="text" name="bldg_name">
          </td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
      <tr> 
        <td> Region </td>
        <td>  
          <input type="text" name="region">
          </td>
        <td> Region Count </td>
        <td>  
          <input type="text" name="regioncount" >
          </td>
        <td> Building ID </td>
        <td>  
          <input type="text" name="bldgid">
          </td>
      </tr>
      <tr> 
        <td> Contact Name </td>
        <td>  
          <input type="text" name="contact_name">
          </td>
        <td> Building Link 
          </td>
        <td>  
          <input type="text" name="bldgid_link">
          </td>
        <td> PGI Link </td>
        <td>  
          <input type="text" name="pgi_link">
          </td>
      </tr>
      <tr> 
        <td> PGI Target </td>
        <td>  
          <input type="text" name="pgi_target">
          </td>
        <td> ERI Link </td>
        <td>  
          <input type="text" name="eri_link">
          </td>
        <td> ERI Target </td>
        <td>  
          <input type="text" name="eri_target">
          </td>
      </tr>
      <tr> 
        <td> Meter Link </td>
        <td>  
          <input type="text" name="meter_link">
          </td>
        <td> Meter Target </td>
        <td>  
          <input type="text" name="meter_target">
          </td>
        <td> OVTHVAC link </td>
        <td>  
          <input type="text" name="ovac_link">
          </td>
      </tr>
      <tr> 
        <td> OVTHVAC Target 
          </td>
        <td>  
          <input type="text" name="ovthvac_target">
          </td>
        <td> POW_AVA Link </td>
        <td>  
          <input type="text" name="pow_ava_link">
          </td>
        <td> POW_AVA Target 
          </td>
        <td>  
          <input type="text" name="pow_ava_target">
          </td>
      </tr>
      <tr> 
        <td> PA Chart Link 
          </td>
        <td>  
          <input type="text" name="pa_chart_link">
          </td>
        <td> PA Chart Target 
          </td>
        <td>  
          <input type="text" name="pa_chart_target">
          </td>
        <td> REV Prof Link 
          </td>
        <td>  
          <input type="text" name="rev_prof_link">
          </td>
      </tr>
      <tr> 
        <td> REV Prof Target 
          </td>
        <td>  
          <input type="text" name="rev_prof_target">
          </td>
        <td> IRI Link </td>
        <td>  
          <input type="text" name="iri_link">
          </td>
        <td> IRI Target </td>
        <td>  
          <input type="text" name="iri_target">
          </td>
      </tr>
      <tr> 
        <td> MEP Link </td>
        <td>  
          <input type="text" name="mep_link">
          </td>
        <td> MEP Target </td>
        <td>  
          <input type="text" name="mep_target">
          </td>
        <td> PQ Link </td>
        <td>  
          <input type="text" name="pq_link">
          </td>
      </tr>
      <tr> 
        <td> PQ Target </td>
        <td>  
          <input type="text" name="pq_target">
          </td>
        <td> PLP Link </td>
        <td>  
          <input type="text" name="plp_link">
          </td>
        <td> PLP Target </td>
        <td>  
          <input type="text" name="plp_target">
          </td>
      </tr>
      <tr> 
        <td> MSI Link </td>
        <td>  
          <input type="text" name="msi_link">
          </td>
        <td> MSI Target </td>
        <td>  
          <input type="text" name="msi_target">
          </td>
        <td> LMP Link </td>
        <td>  
          <input type="text" name="lmp_link">
          </td>
      </tr>
      <tr> 
        <td> LMP Target </td>
        <td>  
          <input type="text" name="lmp_target">
          </td>
        <td> CA Link </td>
        <td>  
          <input type="text" name="ca_link">
          </td>
        <td> CA Target </td>
        <td>  
          <input type="text" name="ca_target">
          </td>
      </tr>
      <tr> 
        <td> PGI </td>
        <td>  
          <%
	  'response.write(rst1("pgi"))
	  'response.write(rst1("iri"))
	  'response.write(rst1("userid"))
	  %>
          <input type="checkbox" name="pgi" value="1">
          </td>
        <td> ERI </td>
        <td>  
          <input type="checkbox" name="eri" value="1">
          </td>
        <td> Meter </td>
        <td>  
          <input type="checkbox" name="meter" value="1">
          </td>
      </tr>
      <tr> 
        <td> OVTHVAC </td>
        <td>  
          <input type="checkbox" name="ovthvac" value="1">
          </td>
        <td> POW AVA </td>
        <td>  
          <input type="checkbox" name="pow_ava" value="1">
          </td>
        <td> PA Chart </td>
        <td>  
          <input type="checkbox" name="pa_chart" value="1">
          </td>
      </tr>
      <tr> 
        <td> REV Prof </td>
        <td>  
          <input type="checkbox" name="rev_prof" value="1">
          </td>
        <td> IRI </td>
        <td>  
          <input type="checkbox" name="iri" value="1">
          </td>
        <td> MEP </td>
        <td>  
          <input type="checkbox" name="mep" value="1">
          </td>
      </tr>
      <tr> 
        <td> PQ </td>
        <td>  
          <input type="checkbox" name="pq" value="1">
          </td>
        <td> PLP </td>
        <td>  
          <input type="checkbox" name="plp" value="1">
          </td>
        <td> MSI </td>
        <td>  
          <input type="checkbox" name="msi" value="1">
          </td>
      </tr>
      <tr> 
        <td> LMP </td>
        <td> 
          <input type="checkbox" name="lmp" value="1">
          </td>
        <td> CA </td>
        <td>  
          <input type="checkbox" name="ca" value="1">
          </td>
        <td></td>
      </tr>
      <tr> 
        <td colspan="2"> 
          <input type="submit" name="choice" value="Add" style="border:1px outset #ddffdd;background-color:ccf3cc;">
          </td>
        <td colspan="2"></td>
        <td colspan="2" align="right"><img src="/images/intranet/btn-back.gif" width="68" height="19" name="goback" onclick="history.back()" onmouseover="buttonOver(this);" onmousedown="buttonDn(this);" onmouseout="buttonOut(this);" style="border:1px solid #eeeeee;"></td>
      </tr>
    </table>
</form>
<%
end if
rst1.close
%>
   </div>
</body>

</html>
