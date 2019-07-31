<!-- #include file="./adovbs.inc" -->
<!-- #include file="secure.inc" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" content="Auscomp eNavigator Suite 2000">
<META NAME="DESCRIPTION" content="Download your FREE evaluation copy at www.auscomp.com">
<TITLE></TITLE>
<script> 
function logoff(){
		
	parent.opener.location.href="https://appserver1.genergy.com/eri_th/login.asp"
	parent.window.close()

}

</script>
</HEAD>
<BODY bgcolor="#FFFFFF">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td bgcolor="#0099FF"> 
      <div align="center"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0" width="89" height="19">
          <param name=movie value="g1.swf">
          <param name=quality value=high>
          <param name="BGCOLOR" value="#0099FF">
          <param name="SCALE" value="exactfit">
          <embed src="g1.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" scale="exactfit" width="89" height="19" bgcolor="#0099FF">
          </embed> 
        </object></div>
    </td>
  </tr>
  <tr> </tr>
  <tr> 
    <td> <APPLET CODEBASE="nav/./" NAME="JNaviTree" ARCHIVE="JAVATreeF.jar" CODE="JAVATreeF.class" WIDTH=220 HEIGHT=500>
        <PARAM NAME="Copyright" VALUE="(c)1997-2001 AUSCOMP (www.auscomp.com)">
        <PARAM NAME="CALENDAR" VALUE="Yes">
        <PARAM NAME="PROXY" VALUE="No">
        <PARAM NAME="IMAGEDIR" VALUE="nav/./">
        <PARAM NAME="FONT" VALUE="Arial; PLAIN;12">
        <PARAM NAME="BACKGROUND" VALUE="255,255,255">
        <PARAM NAME="FOREGROUND" VALUE="0,0,0">
        <PARAM NAME="SELBACKGROUND" VALUE="0,0,255">
        <PARAM NAME="SELFOREGROUND" VALUE="255,255,255">
        <PARAM NAME="TOOLTIPBACKCOLOR" VALUE="255,255,0">
        <PARAM NAME="TOOLTIPFORECOLOR" VALUE="0,0,0">
        <PARAM NAME="LINECOLOR" VALUE="0,0,0">
        <PARAM NAME="LEAFCOLOR" VALUE="0,0,0">
        <PARAM NAME="FOLDERCCOLOR" VALUE="0,0,0">
        <PARAM NAME="FOLDEROCOLOR" VALUE="0,0,0">
        <PARAM NAME="MOUSEOVERFRAMECOLOR" VALUE="0,0,0">
        <PARAM NAME="MOUSEOVER" VALUE="255,0,0">
        <PARAM NAME="BORDER" VALUE="No">
        <PARAM NAME="TOOLTIP" VALUE="Yes">
        <PARAM NAME="DOUBLECLICK" VALUE="No">
        <PARAM NAME="AUTOEXPAND" VALUE="No">
        <PARAM NAME="STYLE" VALUE="AXAX">
        <PARAM NAME="CAL_EXT" VALUE=".htm">
        <PARAM NAME="CAL_PRE" VALUE="../../">
        <PARAM NAME="CAL_LINK" VALUE="None">
        <PARAM NAME="CAL_BG" VALUE="192,192,192">
        <PARAM NAME="CAL_FG" VALUE="0,0,0">
        <PARAM NAME="CAL_HI_BG" VALUE="255,255,255">
        <PARAM NAME="CAL_SUN_FG" VALUE="0,0,255">
        <PARAM NAME="CAL_MONTH" VALUE="Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec">
        <PARAM NAME="CAL_DAY" VALUE="Sa,Mo,Tu,We,Th,Fr,Sa">
        <PARAM NAME="CAL_TYPE" VALUE="MDY">
        <% 	
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set userdetails = Server.CreateObject("ADODB.recordset")
Set userinfo  = Server.CreateObject("ADODB.recordset")
Set meters = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=security;"

strsql = "SELECT * FROM clients WHERE (username='" & Session("loginemail") & "') "  
 
userinfo.Open strsql, cnn1, adOpenStatic

strsql = "SELECT * FROM clientsites WHERE (userid='" & Session("loginemail") & "') order by regioncount,bldgid"  

userdetails.Open strsql, cnn1, adOpenStatic

x=0
regcnt=userinfo("regioncount")
r=0

treelevel=0



%>
        <!-- Nav Header - (CLIENT COMPANY NAME) -->
        <PARAM NAME=<%="I"&x%> VALUE="<%=userinfo("company")%>; None; None; <%=treelevel%> ;  ;  ; ;None; None; NO;CA0;">
        <% x=x+1
	 
Do until (r=regcnt)

if userdetails("regioncount") = r then
%>
        <PARAM NAME=<%="I"&x%> VALUE="<%=userdetails("region") %>; None; None; <%=treelevel%> ;  ;  ; ;None; None; NO;NO;;">
        <% 
	treelevel=treelevel+1
	x=x+1

    Do until userdetails.eof
	if userdetails("regioncount") = r then
		%>
        <PARAM NAME=<%="I"&x%> VALUE="<%=userdetails("bldg_name") %> ; None; None; <%=treelevel%> ;  ;  ; ;None; None; NO;NO;;">
        <%
		treelevel=treelevel+1
		x=x+1
		
		If userdetails("lmp") then
		%>
        <PARAM NAME=<%="I"&x%> VALUE="Load Management; <%=userdetails("lmp_link")%>;<%=userdetails("lmp_target")%>; <%=treelevel%> ;  ;  ; Load Management;None; None; NO;NO;;">
        <% 
		x=x+1
		end if 
		
		If userdetails("eri") then
		%>
        <PARAM NAME=<%="I"&x%> VALUE="Electric Rent Inclusion; <%=userdetails("eri_link") %>;<%=userdetails("eri_target")%> ; <%=treelevel%>;  ;  ; ERI Management Profile;None; None; NO;NO;;">
        <% 
		x=x+1
		end if 
		
		if userdetails("pgi") then
		%>
        <PARAM NAME=<%="I"&x%> VALUE="Power Grid Identification; <%=userdetails("pgi_link") %>; <%=userdetails("pgi_target")%> ; <%=treelevel%>;  ;  ; Power Grid Identification;None; None; NO;NO;;">
        <% 
		x=x+1
		end if 
		
		if userdetails("pow_ava") then
		%>
        <PARAM NAME=<%="I"&x%> VALUE="Power Availability; <%=userdetails("pow_ava_link") %>; <%=userdetails("pow_ava_target")%> ; <%=treelevel%>;  ;  ; Power Availability Profile;None; None; NO;NO;;">
        <% 
		x=x+1
		end if 
		
		if userdetails("pa_chart") then
		%>
        <PARAM NAME=<%="I"&x%> VALUE="Power Chart; <%=userdetails("pa_chart_link") %>; <%=userdetails("pa_chart_target")%>; <%=treelevel%>;  ;  ; Power Chart Profile;None; None; NO;NO;;">
        <% 
		x=x+1
		end if 
		
		if userdetails("rev_prof") then
		%>
        <PARAM NAME=<%="I"&x%> VALUE="Revenue Profile; <%=userdetails("rev_prof_link") %>; <%=userdetails("rev_prof_target")%>; <%=treelevel%>;  ;  ; Revenue Profile;None; None; NO;NO;;">
        <% 
		x=x+1
		end if 
		
		if userdetails("ovthvac") then
		%>
        <PARAM NAME=<%="I"&x%> VALUE="Over Time HVAC; <%=userdetails("ovthvac_link") %>; <%=userdetails("ovthvac_target")%>; <%=treelevel%>;  ;  ; Over Time HVAC Profile;None; None; NO;NO;;">
        <% 
		x=x+1
		end if 
		
		if userdetails("plp") then
		%>
        <PARAM NAME=<%="I"&x%> VALUE="PLP; <%=userdetails("plp_link") %>; <%=userdetails("plp_target") %>; <%=treelevel%>;  ;  ; PLP Profile;None; None; NO;NO;;">
        <% 
		x=x+1
		end if 
		
		if userdetails("iri") then
		%>
        <PARAM NAME=<%="I"&x%> VALUE="IRI; <%=userdetails("iri_link") %>; <%=userdetails("iri_target") %>; <%=treelevel%>;  ;  ; IRI;None; None; NO;NO;;">
        <% 
		x=x+1
		end if 
		if userdetails("meter") then
		%>
        <PARAM NAME=<%="I"&x%> VALUE="Meter Services; <%=userdetails("meter_link") %>; <%=userdetails("meter_target") %>; <%=treelevel%>;  ;  ; Meter Services;None; None; NO;NO;;">
        <% 
		x=x+1
		end if 
		if userdetails("ca") then
		%>
        <PARAM NAME=<%="I"&x%> VALUE="Cost Analysis; <%=userdetails("ca_link") %>; <%=userdetails("ca_target") %>; <%=treelevel%>;  ;  ; Cost Analysis;None; None; NO;NO;;">
        <% 
		x=x+1
		end if 	
	end if  
	
	treelevel=1
	userdetails.movenext
	loop 
	
	userdetails.movefirst
	r = r + 1
else
	userdetails.movenext
end if


treelevel=0
loop
%>
        <PARAM NAME=<%="I"&x %> VALUE="Utilities ; None; None;0 ;  ;  ;;None; None; NO;NO;;">
		<% x=x+1 %>
		<PARAM NAME=<%="I"&x %> VALUE="Manual ; ../manual/contents.htm; main; 1 ;  ;  ;Manual;None; None; NO;NO;;">
		<% x=x+1 %>
		<PARAM NAME=<%="I"&x %> VALUE="(c) Genergy 2001 ; None; None; <%=treelevel%> ;  ;  ; ;None; None; NO;NO;;">
        <!-- Enter code for non-java-enabled browsers here -->
        <p>You will need to activate java to view this menu.</p>
      </APPLET> 
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" height="43"> 
      <div align="center">
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td> 
              <div align="center"><font face="Arial, Helvetica, sans-serif">USER: 
                <%=Session("loginemail") %> </font></div>
            </td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF"> 
              <div align="center"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0" width="100" height="22">
                  <param name=movie value="reload.swf">
                  <param name=quality value=high>
                  <param name="BGCOLOR" value="#FFFFFF">
                  <embed src="reload.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="100" height="22" bgcolor="#FFFFFF">
                  </embed> 
                </object></div>
            </td>
          </tr>
        </table>
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0" width="100" height="22">
          <param name=movie value="logoff.swf">
          <param name=quality value=high>
          <param name="BGCOLOR" value="#FFFFFF">
          <embed src="logoff.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="100" height="22" bgcolor="#FFFFFF">
          </embed> 
        </object></div>
    </td>
  </tr>
</table>
<% 
userdetails.close 
userinfo.close
%>
<p>&nbsp;</p>
</BODY>
</HTML>
