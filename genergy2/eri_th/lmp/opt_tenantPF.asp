<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldg, profiletype, billingid, user, meterid, utility,roleid
bldg = request.querystring("bldg")
meterid = request.querystring("meterid")
billingid= Request("billingid")
profiletype=Request("profiletype")
user=session("loginemail")
utility = request("utility")
roleid = getkeyvalue("roleid")

dim tenantmeter
dim cnn1, rst1
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getLocalConnect(bldg)
%>

<html>
<head>
<title>Untitled Document</title>
<script>
function meterfill(billingid)
{ if(billingid=='0') billingid=''
  parent.document.forms[0].billingid.value = billingid
  parent.document.forms[0].meterid.value = ''
  if(parent.document.all['tab3'].style.backgroundColor=="#0099ff") parent.loadcalendar();else parent.loadchart();
  parent.openLoadBox('loadFrame2')
	document.location.href="opt_tenantPF.asp?bldg=<%=bldg%>&meterid=&utility=<%=utility%>&billingid="+billingid;
}

function loadmeter(meterid)
{ parent.document.forms[0].meterid.value = meterid
  if(parent.document.all['tab3'].style.backgroundColor=="#0099ff") parent.loadcalendar();else parent.loadchart();
  parent.openLoadBox('loadFrame1')
}
</script>

</head><style type="text/css"><!--

BODY {
	SCROLLBAR-FACE-COLOR: #0099FF;
	SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
	SCROLLBAR-SHADOW-COLOR: #333333;
	SCROLLBAR-3DLIGHT-COLOR: #333333;
	SCROLLBAR-ARROW-COLOR: #333333;
	SCROLLBAR-TRACK-COLOR: #333333;
	SCROLLBAR-DARKSHADOW-COLOR: #333333;
}

.grayout 
{
	color:gray;
	font-style : italic;
	font-decoration:none;
}
-->
</style>

<body bgcolor="#FFFFFF" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" onload="parent.closeLoadBox('loadFrame2');">
<form name="lmp" method="post" action="">
  <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000">
    <tr>
      <td width="48%"><font face="Arial, Helvetica, sans-serif" size="2"><b><font color="#FFFFFF">Other 
        Load Profiles</font></b></font></td>
      <td width="52%">
        <div align="right"><font face="Arial, Helvetica, sans-serif" size="2"><b><a href="javascript:document.location.href='<%="options.asp?bldg=" & bldg & "&meterid=" & meterid &"&billingid="& billingid&"&utility="& utility%>'" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Return To Options</a></b></font></div>
      </td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" height="100">
    <tr> 
      <td width="79%" valign="top"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="2">
          <tr valign="top"> 
            <td height="2" width="789"><font face="Arial, Helvetica, sans-serif" size="3" color="#000000">Select Profile Below:</font></td>
          </tr>
          <tr valign="top"> 
            <td height="14" width="789"> <font face="Arial, Helvetica, sans-serif" size="3" color="#000000"> 
              <input type="hidden" name="bldg" value="<%=bldg%>">
              <input type="hidden" name="profiletype" value="<%=profiletype%>">
              <input type="hidden" name="utility" value="<%=utility%>">
              <input type="hidden" name="prev" value="prev">
              <input type="hidden" name="next" value="next">
    <%if roleid="1" then
        response.write "<input value='"& billingid &"' name='billingid' type='hidden'>"
        response.write "<nobr>"
        dim rstrole
        Set rstrole = Server.CreateObject("ADODB.recordset")
        rstrole.open "select billingname from tblLeases where billingid='"&billingid&"'", cnn1
        response.write rstrole("billingname")
        rstrole.close()
        response.write "</nobr>"
    else 'create select box%>
        <select name="billingid" onChange="meterfill(this.value)">
        <option value="0">View Building Profile</option>
        <%
        dim strsql
		strsql = "SELECT TenantNum, tName, billingid FROM tblleases l WHERE l.billingid in (SELECT billingid FROM tblleasesutilityprices lup WHERE utility="&utility&") and leaseexpired=0 and bldgnum='"&bldg&"'"
		rst1.Open strsql, cnn1, adOpenStatic
        do until rst1.EOF
            %><option value="<%=rst1("billingid")%>"<%if trim(billingid)=trim(rst1("billingid")) then response.write " SELECTED"%>><%=rst1("tenantnum") &" - "& rst1("tname")%></option><%
            rst1.movenext
        loop
        rst1.close
        response.write "</select>"
    end if
    response.write "<BR>"
dim cnn2, rst2



Set cnn2 = Server.CreateObject("ADODB.Connection")
Set rst2 = Server.CreateObject("ADODB.recordset")
cnn2.Open getConnect(0,0,"dbCore")
rst2.open "SELECT Label, Link, Target, Active, tbladdonlinks.SID FROM tbladdonlinks JOIN tbladdons on tbladdons.SID=tbladdonlinks.SID WHERE userid='"&session("userid")&"' and CSID=4 ORDER BY listorder", cnn2
if rst2.eof then response.write "Client has no options"
do while not(rst2.eof)
    if rst2("SID")<>1 and rst2("SID")<>5 then response.write "<a href="""&rst2("Link")&"?meterid="& meterid &"&bldg="& bldg &"&billingid="& billingid &"&utility="&utility&""" style=""color:black"" onMouseOver=""this.style.color='gray'"" onMouseOut=""this.style.color='Black'"" onclick=""parent.openLoadBox('loadFrame2');""><font size=""2"" face=""arial"">"&rst2("Label")&"</font></a><br>"
    rst2.movenext
loop
rst2.close
%>
              </font></td>
          </tr>
          &nbsp;<br>
        </table>
      </td>
      <td width="21%" valign="top" bgcolor="#CCCCCC"> 
	  <table width="100%">
		  <%if billingid <> "" then %>
            <tr> 
              <td bgcolor="#0099FF" valign="top"> <div align="center"><font color="#000000" face="Arial, Helvetica, sans-serif" size="1"><b><font face="Arial, Helvetica, sans-serif" size="3"> 
                  </font><font color="#FFFFFF">Meter List</font><font face="Arial, Helvetica, sans-serif" size="3"> 
                  </font></b></font></div></td>
            </tr>
            <%end if%>
		</table>
        <div style="border-width:1; width:100%; height:150; overflow-y: auto; overflow-x: hidden;" name="meterlist"> 
          <table border="0" cellpadding="0" cellspacing="0" width="101%" bgcolor="#FFFFFF">
            <%
			dim offliners
				if billingid <> "" then
					if cint(utility) = 17 then
						strsql = "SELECT meterid, meterid AS hasdata, meternum, lmnum, datasource, online FROM meters WHERE leaseutilityid in (SELECT leaseutilityid FROM tblleasesutilityprices WHERE billingid="&billingid&" and utility="&utility&") ORDER BY meternum"
					else
						strsql = "SELECT meterid, meterid AS hasdata, meternum, lmnum, datasource, online FROM meters WHERE leaseutilityid in (SELECT leaseutilityid FROM tblleasesutilityprices WHERE billingid="&billingid&" and utility="&utility&") and meters.lmp = 0 ORDER BY meternum"
					end if
'		  response.write strsql
					rst1.Open strsql, cnn1, adOpenStatic
	    			do until rst1.eof
        				dim italic, italicx 'this is for italisizing the meter link if it has data
	    			    italic=""
	    			    italicx=""
	    			    if IsNumeric(trim(rst1("hasdata"))) then 'if hasdata has a meter number in it then pulse has a data set for that meter number
	    			        italic="<u>"
	    			        italicx="</u>"
	    			    end if%>
            <tr> 
              <td <%if cint(rst1("online"))=0 then%><%offliners = true%>class="grayout"<%end if%> onMouseOver="this.style.background='lightgreen'" onMouseOut="this.style.background='white'" height="19"><a <%if IsNumeric(trim(rst1("hasdata"))) then%>onClick="loadmeter(<%=rst1("meterid")%>);" href="javascript:parent.loadchart()"<%end if%> target="lmp" style="text-decoration:none; color:black; font-size:12px; font-family:arial"><%= italic &rst1("meternum")& italicx %></a></td>
            </tr>
            <%rst1.movenext
		    		loop
				End if%>
          </table>
          
        </div>
        <div align="center"><a href="javascript:meterfill(parent.document.forms[0].billingid.value)" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'"><font color="#000000" face="Arial, Helvetica, sans-serif" size="1"><b><u>View Aggregated Profile</u></b></font></a><br>
		<%if offliners then%><span class="grayout"><font color="#000000" face="Arial, Helvetica, sans-serif" size="1">Offline&nbsp;Meters&nbsp;Italicized</font></span><%end if%>
        </div></td>
    </tr>
	<tr><tr colspan="2"></tr></tr>
    <tr>
      <td height="20" width="789" valign="bottom" bgcolor="#000000"> 
        <div align="left">&nbsp;</div>
      </td>
      <td width="21%" valign="top" bgcolor="#000000">&nbsp;</td>
    </tr>
  </table>
</form>
</body>
</html>