<%option explicit%>
<%' TT 5/22/2008 This is an altered version on entry.asp in order to eliminate full portfolio drop down option and jump to option.  %>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Utility Bill Entry</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">		
</head>
<%
Dim cnn1, rst1, rst2, str1, str
Set cnn1 = Server.CreateObject("ADODB.connection")

dim acctid, pid, building
acctid=Request.querystring("acctnum")
pid=Request.Querystring("pid")
building=Request.Querystring("building")
cnn1.Open getConnect(pid,building,"Billing")
'response.write pid
'response.end
%>
<script>
function acctselect(building,utility){
	var temp = "acctlist.asp?building=" + building+ "&utility="+utility
	window.open(temp,"selectaccount", "scrollbars=yes,width=1000, height=300, status=no" );
//	window.selectaccount.focus();
}
function ypid(id1,building,utility){
	var temp = "ypid.asp?acctid=" + id1 + "&building=" + building+ "&utility=" + utility
	window.open(temp,"ypid", "scrollbars=yes,width=500, height=300, status=no" );
}
function setup(building,utility){
	clearselections();
	var temp = "editacct.asp?building="+building+ "&utility=" + utility
	document.all.entryframe.style.visibility = "visible"
	document.frames.entry.location=temp;
	}

function editacct(acctid)
{	if(acctid.length>0)
	{	var temp = "editacct.asp?acctid=" +acctid
		document.frames.entry.location=temp;
	}else
	{	alert('Select an Account');
	}
}
function clearselections()
{	document.frames.entry.location='about:blank';
	document.form1.acctid.value=''
	document.all['accountdisplay'].innerText='No Account Selected'
	document.all['enterbillbutton'].style.visibility='hidden'
}


function loadportfolio()
{	var frm = document.forms['form1'];
	var newhref = "entryG1.asp?pid="+frm.pid.value;
	document.location.href=newhref;
}

function loadbuilding()
{	var frm = document.forms['form1'];
	var newhref = "entryG1.asp?pid="+frm.pid.value+"&building="+frm.building.value;
	document.location.href=newhref;
}



</script>
<body bgcolor="#eeeeee" text="#000000">
<table width="100%" border="0">
	<form name="form1" method="post" action="">      
		<tr>
			<td>
				<table width="100%" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="91%" bgcolor="#6699cc"><span class="standardheader">Utility Bill Entry</span></td>
          
             
             
                <table width="100%" border="0">
					<%Set rst1 = Server.CreateObject("ADODB.recordset")
					if not(trim(pid)="" and trim(building)<>"") then%>
						<tr>
							<td width="25%" height="27">
								Portfolio:
							</td>
							<td width="75%" height="27">								
								<%if allowGroups("Genergy Users") then%>
									<select name="pid" onChange="loadportfolio()">
									
										<option value="">Select Portfolio</option>
										
								<%if trim(pid)<>"" then 
											 rst1.open "SELECT distinct id, name FROM portfolio p where id="&pid&"", getConnect(0,0,"dbCore")
								            do until rst1.eof
						
									%><option value="<%=trim(rst1("id"))%>"<%if trim(rst1("id"))=trim(pid) then %>SELECTED<%end if%>>
												<%=rst1("name")%>
											</option><%
											
											rst1.movenext
									
										loop
										
										rst1.close%>
									</select>
									<%end if %>
											
									<% 	 if trim(pid)="" then 
											 rst1.open "SELECT distinct id, name FROM portfolio p ORDER BY name", getConnect(0,0,"dbCore")
									          do until rst1.eof
						%><option value="<%=trim(rst1("id"))%>"<%if trim(rst1("id"))=trim(pid) then %>SELECTED<%end if%>>
												<%=rst1("name")%>
											</option><%
											rst1.movenext
											
										loop
										
										rst1.close%>
									</select><%end if %>
								
									
									
									
								<%elseif isnumeric(pid) then
									rst1.open "SELECT name FROM portfolio WHERE id="&pid&" ORDER BY name", cnn1
									if not rst1.eof then response.write rst1("name") end if
									rst1.close%>
									<
									input type="hidden" name="pid" value="<%=pid%>">
								<%end if%>								
							</td>
						</tr>
						<tr>
							<td width="25%" height="27">Building:</td>
							<td width="75%" height="27">
				
								
								
							<%if trim(pid)<>"" then %>
									<select name="building" onChange="loadbuilding();">
									<option selected>Select Building</option>
									
									
								
								<%rst1.open "SELECT BldgNum, strt FROM buildings b WHERE portfolioid='"&pid&"' ORDER BY strt", cnn1
										do until rst1.eof %>
											%><option  <%if isBuildingOff(rst1("bldgnum")) then%>class="grayout"<%end if%> value="<%=trim(rst1("Bldgnum"))%>"<%if trim(rst1("Bldgnum"))=trim(building) then%>selected<%end if%>>
												<%=rst1("strt")%> (<%=rst1("Bldgnum")%>)
											</option><%
											rst1.movenext
										loop
										rst1.close%>
									</select>								
								<%else%>
									<input type="hidden" name="building" value="">No Building Selected
								<%end if%>
								</td>
						</tr>
					<%else%>
						<tr>
							<td width="25%" height="27">Building:</td>
							<td width="75%"height="27">
								<%
								rst1.open "SELECT BldgNum, strt FROM buildings WHERE bldgnum='"&building&"' ORDER BY strt", getLocalConnect(building)
								if not rst1.eof then %>
									<input type="hidden" name="building" value="<%=building%>"><%=rst1("strt")%>
								<%end if%>
							</td>
						</tr>
					<%end if%>
					<tr>
						<td width="25%" height="26">Utility :</td>
				<td width="75%">
							<select name="utility" onChange="clearselections();"><% 
								Set rst2 = Server.CreateObject("ADODB.recordset")
								str="select * from tblutility order by utility"
								rst2.Open str, getConnect(0,0,"dbCore"), 0, 1, 1
								do until rst2.eof
									%><option value="<%=rst2("utilityid")%>" <%if lcase(rst2("utility"))="electricity" then%>SELECTED<%end if%>>
										<%=rst2("utilitydisplay")%>
									</option><%
									rst2.movenext
								loop
								rst2.close%>
							</select>
						</td>
					</tr> 
					<tr>
						<td width="25%" height="26" valign="top">Account Number :</td>
				<td width="75%">				
							<span id="accountdisplay">No Account Selected</span><br>
							<input type="button" name="Submit" value="Select Account" onClick="acctselect(building.value,utility.value)">
							<%if not(isBuildingOff(building)) then%><input type="button" name="this" value="Setup New Account" onClick="setup(building.value,utility.value)"><%end if%>
							<span id="enterbillbutton" style="visibility:hidden">
								<input type="button" name="acctinfo" value="<%if isBuildingOff(building) then%>View<%else%>Enter<%end if%> Bill" onClick="ypid(acctid.value,building.value,utility.value)">
							</span>				
						</td>
					</tr>
					<tr>    
						<td width="25%">
							<input type="hidden" name="bldg"><input type="hidden" name="id1"><input type="hidden" name="acctid">
						</td>
						<td width="75%">&nbsp;</td>
					</tr>
				</table>
			</td>
		</tr>		
	</form>
</table>
<%
'TK: 04/28/2006
on error resume next
If rst1.State = 1 Then
	rst1.Close 
End If
If rst2.State = 1 Then
	rst2.Close 
End If
set rst1 = nothing
set rst2 = nothing

set cnn1=nothing%>
<div id="entryframe" style="visibility:hidden">
<IFRAME name="entry" width="100%" height="550" src="" scrolling="no" marginwidth="0" marginheight="0" frameborder="0"></IFRAME>
</div></body>
</html>