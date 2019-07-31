<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!-- #include virtual="/genergy2_Intranet/itservices/ttracker/TTServices.inc" -->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

'N.Ambo 5/27/2009 amended page so that the values for 'Account Rate' are taken from the tbale 'ratetypes', field 'type' instead of field 'typecheck'


dim pid, bldg, tid, lid
pid = request("pid")
bldg = request("bldg")
tid = request("tid")
lid = request("lid")

dim cnn1, rst1, strsql
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

'dim DBmainmodIP
'DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."

dim AdminFee, TenantRate, AddonFee, ModifyRate, Coincident, CoinWithPeak, FullOnPeak, utility, procname, customsrc, calcIntPeak, acctid, use_acctid, billnote, shadow, ticketcount, masterticketid,opentickets, criticalopentickets, totaltickets
if trim(lid)<>"" then
	rst1.Open "SELECT * FROM tblleasesutilityprices WHERE leaseutilityid='"&lid&"'", cnn1
	if not rst1.EOF then
		AdminFee = rst1("AdminFee")
		TenantRate = rst1("rateTenant")
		AddonFee = rst1("AddonFee")
		ModifyRate = rst1("rateModify")
		Coincident = rst1("Coincident")
		CoinWithPeak = rst1("Coincident_peak")
		FullOnPeak = rst1("FullOnPeak")
		utility = rst1("utility")
		procname = rst1("procname")
		customsrc = rst1("customsrc")
		calcIntPeak = rst1("calcintpeak")
		acctid = rst1("acctid")
		use_acctid = rst1("use_acctid")
		billnote = rst1("bill_note")
		shadow = rst1("shadow")
	end if
	rst1.close

end if
if trim(utility)="" then utility=0
dim bldgname, portfolioname, rid
if trim(bldg)<>"" then
  rst1.open "SELECT bldgname, name, region FROM buildings b join portfolio p on b.portfolioid=p.id WHERE bldgnum='"&bldg&"'", cnn1
	if not rst1.EOF then
		bldgname = rst1("bldgname")
		portfolioname = rst1("name")
		rid = rst1("region")
	end if
	rst1.close
end if

dim billingname
if trim(bldg)<>"" then
  rst1.open "SELECT billingname FROM tblleases WHERE billingid='"&tid&"'", cnn1
	if not rst1.EOF then
		billingname = rst1("billingname")
	end if
	rst1.close
end if
if trim(lid) <> "" and bldg <> "" then 
	dim ticket
	set ticket = New tickets
	ticket.Label="Lease Utility"
	ticket.Note="Master Ticket for Lease ID "&split(getBuildingIP(bldg),"\")(1)&"-"&lid
	ticket.ccuid  = "rbdept"
	ticket.client = 1
	if lid<>"0" then ticket.findtickets "leaseid", split(getBuildingIP(bldg),"\")(1)&"-"&lid
end if 
%>
<html>
<head>
<title>Building View</title>
<script>
function openCustomWin(clink, cname, cspec)
{	customlink = open(clink, cname, cspec)
	customlink.focus()
}

function meterEdit(meterid)
{	
  document.location = "contentfrm.asp?action='meteredit'&pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&meterid="+meterid;
//  document.location.href = 'meteredit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&meterid='+meterid
}

function goToRates(){
  if ((parent.name=="main") && (parent.frames.length >= 2)) { 
    top.main.location="rateTypeView.asp?rid=<%=rid%>";
  } else {
    location="rateTypeView.asp?rid=<%=rid%>";
  }
}
</script>
<link rel="Stylesheet" href="setup.css" type="text/css">
</head>

<body>
<form name="form2" method="post" action="leaseutilitysave.asp">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr bgcolor="#3399cc">
	<td><span class="standardheader">
		<%if trim(lid)<>"" then%>
			Update Lease Utility | <span style="font-weight:normal;"><a href="portfolioeditG1.asp?pid=<%=pid%>" style="color:#ffffff;"><%=portfolioname%></a> &gt; <a href="buildingeditG1.asp?pid=<%=pid%>&bldg=<%=bldg%>" style="color:#ffffff;"><%=bldgname%></a> &gt; <a href="tenanteditG1.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>" style="color:#ffffff;"><%=billingname%></a></span>
		<%else%>
			Add New Lease Utility  | <span style="font-weight:normal;"><a href="portfolioeditG1.asp?pid=<%=pid%>" style="color:#ffffff;"><%=portfolioname%></a> &gt; <a href="buildingeditG1.asp?pid=<%=pid%>&bldg=<%=bldg%>" style="color:#ffffff;"><%=bldgname%></a> &gt; <a href="tenanteditG1.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>" style="color:#ffffff;"><%=billingname%></a></span>
		<%end if%>
	</span></td>
  </tr>
</table>

  <table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#eeeeee">
    <tr bgcolor="#eeeeee">
      <td align="left" colspan=2>&nbsp;<b>Detail for Lease ID SVR<%=split(getBuildingIP(bldg),"\")(1)%>-<%=lid%></b></td>
      <td>&nbsp;</td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Utility Type</span></td>
      <td><select name="utility">
          <%
			rst1.open "SELECT * FROM tblutility ORDER BY utilitydisplay", cnn1
			do until rst1.eof
				%>
          <option value="<%=rst1("utilityid")%>"<%if trim(utility)=trim(rst1("utilityid")) then response.write " SELECTED"%>><%=rst1("utilitydisplay")%></option>
          <%
				rst1.movenext
			loop
			rst1.close
			%>
        </select> </td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Admin Fee</span></td>
      <td><input type="text" name="AdminFee" value="<%=AdminFee%>"></td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right">
        <%if utility=0 then%>
        <input type="hidden" name="acctidref" value="0">
        <%else%>
        Billing Account
        <%end if%>
      </td>
      <td>
        <%if utility<>0 then%>
        <select name="acctidref">
          <option value="0">Default</option>
          <%rst1.open "SELECT * FROM tblacctsetup WHERE utility="&utility&" and esco=0 and bldgnum='"&bldg&"'", cnn1
						do until rst1.eof%>
          <option value="<%=rst1("acctid")%>"<%if trim(acctid)=trim(rst1("acctid")) then response.write " SELECTED"%>><%=rst1("acctid")%></option>
          <%
							rst1.movenext
						loop
						rst1.close
						%>
        </select>
        &nbsp;
        <%end if%>
        <%if allowGroups("IT Services,AllEnergyServicesEmp") then%>
        <input type="checkbox" name="use_acctid" value="1" <%if use_acctid then response.write "CHECKED"%>>
        &nbsp;Uses&nbsp;Account&nbsp;in&nbsp;Rate
        <%else%>
        <%if use_acctid then response.write "<font color=""red"">Uses&nbsp;Account&nbsp;in&nbsp;Rate</font>"%>
        <input type="hidden" name="use_acctid" value="<%if use_acctid then response.write "1" else response.write "0"%>">
        <%end if%>
      </td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Account Rate</span></td>
      <td><select name="TenantRate">
          <%
			rst1.open "SELECT * FROM ratetypes WHERE regionid in (SELECT region FROM buildings WHERE bldgnum='"& bldg &"') ORDER BY type", cnn1
			do until rst1.eof
				%>
          <option value="<%=rst1("id")%>"<%if trim(tenantrate)=trim(rst1("id")) then response.write " SELECTED"%>><%=rst1("type")%></option>
          <%
				rst1.movenext
			loop
			rst1.close
			%>
        </select> 
        
      </td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Rate Function</span></td>
      <td><select name="procname">
          <%
         
         rst1.open "SELECT * FROM functiontypes ORDER BY description", cnn1
			' rst1.open "SELECT * FROM functiontypes where description = 'ouc invoice'", cnn1 --OUC specific rates
			do until rst1.eof
				%>
          <option value="<%=rst1("id")%>"<%if trim(procname)=trim(rst1("id")) then response.write " SELECTED"%>><%=rst1("description")%></option>
          <%
				rst1.movenext
			loop
			rst1.close
			%>
        </select> </td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Add-on Fee</span></td>
      <td><%=AddonFee%></td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Modify Rate</span></td>
      <td><input type="text" name="ModifyRate" value="<%=ModifyRate%>"></td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Coincident</span></td>
      <td><input type="checkbox" value="1" name="Coincident" <%if Coincident="True" then Response.Write "CHECKED"%> onclick="this.form.CoinWithPeak.checked=false"> 
      </td>
      <td rowspan="4" valign="top"> Note on Bill:<br> <textarea cols="45" rows="5" name="billnote" onKeyUp="if(this.value.length>250){this.value=this.value.substr(0,250)}"><%=billnote%></textarea> 
      </td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Coincident w/Building Peak</span></td>
      <td><input type="checkbox" value="1" name="CoinWithPeak" <%if CoinWithPeak="True" then Response.Write "CHECKED"%> onclick="this.form.Coincident.checked=false"> 
      </td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Full On Peak</span></td>
      <td><input type="checkbox" value="1" name="FullOnPeak" <%if FullOnPeak="True" then Response.Write "CHECKED"%>></td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Intermediate Peak </span></td>
      <td><input type="checkbox" value="1" name="calcintpeak" <%if calcIntPeak="True" then Response.Write "CHECKED"%>></td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td align="right"><span class="standard">Shadow Bill</span></td>
      <td><input type="checkbox" value="1" name="shadow" <%if shadow="True" then Response.Write "CHECKED"%>></td>
    </tr>
    <tr bgcolor="#eeeeee"> 
      <td><span class="standard">&nbsp;</span></td>
      <td> 
        <%if trim(lid)<>"" then%>
      
        <!--[[input type="submit" name="action" value="Delete" class="standard"  style="border:1px outset #ddffdd;background-color:ccf3cc;"]]-->
      
        <%
        if trim(customsrc)<>"" then
        response.write "*Contains custom fields"
        end if
      %>
        <%else%>
        <input type="submit" name="action" value="Save" class="standard"  style="border:1px outset #ddffdd;background-color:ccf3cc;"> 
        <input type="button" value="Cancel" onclick="document.location='tenanteditG1.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>';" class="standard"  style="border:1px outset #ddffdd;background-color:ccf3cc;"> 
        <%end if%>
      </td>
    </tr>
    <% if trim(lid)<>"" then %>
    <tr bgcolor="#eeeeee"> 
      <td >&nbsp;</td>
      <td > <img src="images/aro-rt.gif" align="absmiddle"  hspace="2" border="0"><a href="javascript:openCustomWin('customsetup/customcredit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>','customlink', 'width=400,height=200, scrollbars=yes');"></a> 
        <%
  rst1.open "SELECT * FROM custom_links WHERE code=3 and unitid='"&pid&"'", cnn1
  do while not rst1.eof
    response.write "<img src=""images/aro-rt.gif"" align=""absmiddle""  hspace=""2"" border=""0""><a href=""javascript:openCustomWin('"&rst1("link")&"?pid="&pid&"&bldg="&bldg&"&tid="&tid&"&lid="&lid&"','customlink', 'width="&rst1("width")&",height="&rst1("height")&", scrollbars=yes');"">"&rst1("label")&"</a><br>"
    rst1.movenext
  loop
  rst1.close
  %>
        <br> </td>
    </tr>
    <% end if %>
  </table>
<% if trim(lid)<>"" then %>
<table width="100%" border=0 cellpadding="3" cellspacing="0">
    <tr bgcolor="#eeeeee">
      
	  </td>
    </tr>
</table>
	<%end if %>
<input type="hidden" name="pid" value="<%=pid%>">
<input type="hidden" name="bldg" value="<%=bldg%>">
<input type="hidden" name="tid" value="<%=tid%>">
<input type="hidden" name="lid" value="<%=lid%>">
</form>
</body>
</html>
