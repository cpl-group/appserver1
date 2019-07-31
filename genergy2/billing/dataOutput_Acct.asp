<%option explicit
server.ScriptTimeout = 580
%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>gEnergyOne</title>
<style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}


-->
</style>
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<script src = "/genergy2/sorttable.js" type="text/javascript"></script>
</head>
<body BGCOLOR="#eeeeee" LINK="#0000CC" VLINK="#0000CC" TEXT="#000000">
<form method="POST" name="form1" action="dataOutput_Acct.asp">
<table border="0" cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#6699cc"><td colspan="2"><span class="standardheader"><%=buildingname%> Data  Downloading</span></td></tr>
<%
dim pid, byear, bperiod, building, utilityid, buildingname, procedure, downloadlink, recCount, coreCmd, i

procedure =""

pid = request("pid")
byear = request("byear")
bperiod = request("bperiod")
building = request("building")
utilityid = request("utilityid")
procedure = request("procedure")

dim sql, rst1, rst2, cnn1, cmd, prm
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.createobject("ADODB.recordset")
set cmd = server.createobject("ADODB.command")
set coreCmd = server.createobject("ADODB.command")
set cnn1 = server.createobject("ADODB.connection")
cnn1.open getConnect(pid,building,"billing")
cnn1.CommandTimeout = 580
cmd.ActiveConnection = cnn1

rst1.open "SELECT * FROM buildings WHERE bldgnum='"&building&"'", cnn1
if not rst1.eof then buildingname = rst1("strt")
rst1.close

if trim(procedure)<>"" then
Err.clear
on Error resume next
	if (pid <> "108") then
	
	cnn1.CursorLocation = adUseClient
	cmd.CommandText = procedure
	cmd.CommandType = adCmdStoredProc
	Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("util", adInteger, adParamInput)
	cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("file", adVarChar, adParamOutput, 80)
	cmd.Parameters.Append prm

	Set cmd.ActiveConnection = cnn1
	cmd.Parameters("bldg")		= building
	cmd.Parameters("by")		= byear
	cmd.Parameters("bp")		= bperiod
	cmd.Parameters("util")		= utilityid

	'response.Write(procedure + " " + building + "," + byear + "," + bperiod + "," + utilityid)
    'response.End()	
	
	else
		
	cnn1.CursorLocation = adUseClient
	cmd.CommandText = "sp_PaDtsStart"
	cmd.CommandType = adCmdStoredProc
	
	Set prm = cmd.CreateParameter("by", adVarChar, adParamInput,10)
	cmd.Parameters.Append prm
	
	Set prm = cmd.CreateParameter("BP", adVarChar, adParamInput, 10)
	cmd.Parameters.Append prm
	
	Set prm = cmd.CreateParameter("BLDG", adVarChar, adParamInput, 20)
	cmd.Parameters.Append prm
	
	Set prm = cmd.CreateParameter("FLAG", adVarChar, adParamInput, 20)
	cmd.Parameters.Append prm

	Set prm = cmd.CreateParameter("util", adInteger, adParamInput)
	cmd.Parameters.Append prm
		
	
	Set cmd.ActiveConnection = cnn1
	cmd.Parameters("BP") = bperiod
	cmd.Parameters("by") = byear
	cmd.Parameters("BLDG") = building
	cmd.Parameters("FLAG") ="0"
	cmd.Parameters("util")		= utilityid
	
	'response.Write("sp_PaDtsStart" + " " +  bperiod +"," + building)
   ' response.End()	
	
	end if

	cmd.execute
	if err.number = 0 then 
	%>
	<tr><td colspan="2" align="center">&nbsp;<br>
			Data Files for <b><%=buildingname%></b> 
			period <%=bperiod%>, <%=byear%><br>have been created.<br>They can be accessed via your<br>gEnergyOne Data File Access Module.
		   </td></tr>
	<%
	else
	%>
	<tr><td colspan="2" align="center">&nbsp;<%=err.description%><br>
			Data Files for <b><%=buildingname%></b> 
			period <%=bperiod%>, <%=byear%><br>failed to be created.<br>Please try again or contact support@genergy.com
		   </td></tr>
	<%
	end if
else%>
<tr><td colspan="2">Data File Type For period <%=bperiod%>, <%=byear%>:</td></tr>
<tr valign="top"> 
    <% if pid <> "108" then%>
	<td width="5%">
      <select name="procedure">
        <%
          rst1.open "SELECT * FROM ADF WHERE (bldgnum='"&building&"' or bldgnum is null) and (pid="&pid&" or pid is null) ORDER BY label", getConnect(pid,building,"Billing")
          do until rst1.eof
            %><option value="<%=rst1("procname")%>"><%=rst1("label")%></option><%
            rst1.movenext
          loop
          rst1.close
        %>
      </select>
    </td>
    <%end if%>
	
	  <%
	     rst1.open "SELECT * FROM bldg_acct_trans WHERE bldgNum = '" + building + "' AND Billyear = '" + byear + "' AND billperiod = " + bperiod, cnn1
	     	  
	  if pid <> "108" then%>	
	<td width="95%"><input type="button" value="Create" onclick="document.location.href='loading.asp?url=<%=server.urlencode("/genergy2/billing/dataOutput.asp?pid="&pid&"&byear="&byear&"&bperiod="&bperiod&"&building="&building&"&utilityid="&utilityid)%>%26procedure%3D'+this.form.procedure.value"></td>
		<%
		else if not rst1.eof then 
		
		    coreCmd.ActiveConnection = cnn1
            coreCmd.CommandText = "sp_getAcctTransForBuilding"
            coreCmd.CommandType = adCmdStoredProc
            
            Set prm = coreCmd.CreateParameter("bldgNum", adVarChar, adParamInput, 50)
            coreCmd.Parameters.Append prm
            'Set prm = coreCmd.CreateParameter("dateFrom", adVarChar, adParamInput, 20)
            'coreCmd.Parameters.Append prm
            'Set prm = coreCmd.CreateParameter("dateTo", adVarChar, adParamInput, 20)
            'coreCmd.Parameters.Append prm
            Set prm = coreCmd.CreateParameter("byear", adVarChar, adParamInput, 10)
            coreCmd.Parameters.Append prm
            Set prm = coreCmd.CreateParameter("bperiod", adTinyInt, adParamInput)
            coreCmd.Parameters.Append prm
            'Set prm = coreCmd.CreateParameter("lid", adInteger, adParamInput)
            'coreCmd.Parameters.Append prm
            'Set prm = coreCmd.CreateParameter("UTIL", adInteger, adParamInput)
            'coreCmd.Parameters.Append prm
            
            coreCmd.Parameters("bldgNum") = trim(building)
            'coreCmd.Parameters("dateFrom") = "1/1/2000"
            'coreCmd.Parameters("dateTo") = "1/1/2000"
            coreCmd.Parameters("byear") = trim(byear)
            coreCmd.Parameters("bperiod") = CInt(bperiod)
            'coreCmd.Parameters("lid") = 0
            'coreCmd.Parameters("UTIL") = CInt(utilityId)
        
            rst2.open coreCmd
		%>
		<td>
            <div style="margin: 10px; overflow: auto">
        		<table id="sortTable" class="sortable" style="font-size: 11px; font-family: Arial, Helvetica, sans-serif;" cellspacing="1" cellpadding="3" border="1" width="99%">
                	<thead align="center">
                    	<% for i = 0 to rst2.fields.Count - 1 %>
                        <th><a href="#"><%=rst2.fields(i).Name%></a></th>
                        <%next%>
                    </thead>
                    <tbody align="center">
                    	<%do while not rst2.eof%>
                    	<tr onMouseOver="this.style.backgroundColor = 'lightgreen'" onMouseOut="this.style.backgroundColor = '#eeeeee'" onclick="popUp('/genergy2/accounting_files/accountingFile_output.asp?bldg=<%=building%>&util=<%=utilityid%>&lid=0&billyear=<%=byear%>&billperiod=<%=bperiod%>&dateFrom=<%=rst2("transDate")%>', 'CreateFiles')" > 
                    	<%for i = 0 to rst2.fields.Count - 1%>
                    		<td style="border-bottom: 1px solid #CCCCCC"><%=UCase(rst2(i))%></td>
                    	<%next%>
                    	</tr><%
                    rst2.movenext
                    loop%>
                	</tbody>
                	<tr><td>Create from Current data >>> <input type="button" value="Create" onclick="popUp('/genergy2/accounting_files/accountingFile_output.asp?bldg=<%=building%>&util=<%=utilityid%>&lid=0&billyear=<%=byear%>&billperiod=<%=bperiod%>&dateFrom=<%="mm/dd/yy"%>', 'CreateFiles')"></td></tr>
            	</table> 
            </div>
            </td>
             <%else %>
            <td width="95%"><input type="button" value="Create" onclick="document.location.href='loading.asp?url=<%=server.urlencode("/genergy2/accounting_files/accountingFile_output.asp?pid="&pid&"&byear="&byear&"&bperiod="&bperiod&"&building="&building&"&utilityid="&utilityid)%>%26procedure%3D'+'PA'"></td>
          
          <%  end if 
              'rst2.close
              rst1.close
          end if%>

</tr>
<%end if%>
</table>
<input name="pid" value="<%=pid%>" type="hidden">
<input name="byear" value="<%=byear%>" type="hidden">
<input name="bperiod" value="<%=bperiod%>" type="hidden">
<input name="building" value="<%=building%>" type="hidden">
<input name="utilityid" value="<%=utilityid%>" type="hidden">
</form>
<script type="text/javascript">

function popUp(newPage, title){
  popper = window.open(newPage, title,"width=400,height=300,scrollbars=1,status=0,resizeable=1")
}

</script>
</body>
</html>
