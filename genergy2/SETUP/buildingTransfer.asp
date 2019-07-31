<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, newPid, note, coreCmd, prm, result, buildingStr, bldgNum, message

pid = request("pid")
bldg = request("bldg")
newPid = request("newPid")
buildingStr = request("buildingStr")

bldgNum = split(buildingStr, "+")
  
dim cnn1, rst1, strsql, rst2
set cnn1 = server.createobject("ADODB.connection")
set coreCmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.createobject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

cnn1.CursorLocation = adUseClient
coreCmd.activeConnection = cnn1
    
dim bldgname, portfolioname
if trim(bldg)<>"" then
  rst1.open "SELECT bldgname, name FROM buildings b INNER JOIN portfolio p on b.portfolioid=p.id WHERE bldgnum='"&bldg&"'", cnn1
	if not rst1.EOF then
		bldgname = rst1("bldgname")
		portfolioname = rst1("name")
	end if
	rst1.close
end if

' execute the following if submit is clicked
if (request("submit") = "Submit") then
    
    coreCmd.CommandText = "sp_transfer_bldg " 'stored Procedure name goes here
    coreCmd.CommandType = adCmdStoredProc
    
    set prm = coreCmd.CreateParameter("BldgNum", adVarChar, adParamInput, 10)
    coreCmd.Parameters.Append prm
    set prm = coreCmd.CreateParameter("pidOld", adInteger, adParamInput)
    coreCmd.Parameters.Append prm
    set prm = coreCmd.CreateParameter("pidNew", adInteger, adParamInput)
    coreCmd.Parameters.Append prm
    set prm = coreCmd.CreateParameter("user", adVarChar, adParamInput, 50)
    coreCmd.Parameters.Append prm
    set prm = coreCmd.CreateParameter("date", adVarChar, adParamInput, 12)
    coreCmd.Parameters.Append prm
    set prm = coreCmd.CreateParameter("note", adVarChar, adParamInput, 1000)
    coreCmd.Parameters.Append prm
    set prm = coreCmd.CreateParameter("result", adTinyInt, adParamOutput)
    coreCmd.Parameters.Append prm
    
    coreCmd.Parameters("pidOld") = pid
    coreCmd.Parameters("pidNew") = newPid
    coreCmd.Parameters("user") = trim(Session("login"))
    coreCmd.Parameters("date") = trim(secureRequest("xferDate"))
    coreCmd.Parameters("note") = trim(secureRequest("note"))
    coreCmd.Parameters("result") = result
    
    message = "Building Transfered successfully!"
        
    for each x in bldgNum
        response.write(CStr(x))
        coreCmd.Parameters("BldgNum") = trim(x)
        coreCmd.execute()    
        result = coreCmd.Parameters("result")
        if (result <> 0) then
            message = "Building Number " + x + " transfered failed!"
            exit for
       end if
    next    
    
    response.write("<h4 style='color:red;'> " + message + " </h4>")
    end if    

%>
<html>
<head>
<title>Building Transfer</title>
<script src="buildingXfer.js" type="text/javascript"></script>
<script src="/genergy2/calendar.js" type="text/javascript"></script>
<link rel="Stylesheet" href="setup.css" type="text/css">
    <style type="text/css">
        .style1
        {
            height: 70px;
        }
    </style>
</head>
<body bgcolor="#ffffff">
<form name="buildingXferFrm" method="post" action="buildingTransfer.asp" >

<table width="100%" >
<tr bgcolor="#3399cc">
	<td colspan="2">
    <span class="standardheader">
   &nbsp; Building Transfer </span>
	</td>
</tr>
<tr bgcolor="#fffff" valign="top">
  <td bgcolor="#ffffff" width="50%">
  <table border=0 cellpadding="3" cellspacing="0" width="100%">
  
<tr>
    <% 
    
    strsql = "SELECT id, name FROM Portfolio ORDER BY name"
    rst1.open strsql, getConnect(0,0, "dbCore")
    if not rst1.eof then
    
    %>
    
    <td align="right"> Old Protfolio Name: </td>
    <td><select name="pid" onchange="pidChange(this.value)">
        <option value="">Select a Protfolio...</option>    
        <% do until rst1.eof %>
        <option value="<%=rst1("id")%>" <%if (trim(pid) = CStr(rst1("id"))) then %> selected <% end if %>><%=rst1("name")%></option>
        <% rst1.movenext
            loop
            rst1.close %>
    </select>
    </td>
       <% end if %>
    </tr>
    <tr>
        <% rst1.open strsql, getConnect(0,0, "dbCore")
            if not rst1.eof then
        %>
       <td align="right"> New Protfolio Name: </td>
       <td><select name="newPid">
            <option value="" selected >Select a Protfolio...</option>
            <% do until rst1.eof %>
            <option value="<%=rst1("id")%>" <% if (trim(newPid) = CStr(rst1("id"))) then %> selected <% end if %>><%=rst1("name")%></option>
            <%
              rst1.movenext
              loop
              rst1.close
             %>
            </select></td>
            <% end if  %>
    </tr>
    <tr>
        <td align="right">Transfer Date:</td>
        <td><input type="text" name="xferDate" value="dd/mm/yy" onfocus="this.select();lcs(this)" onclick="event.cancelBubble=true;this.select();lcs(this)" /></td>
    </tr>
    <tr>
        <td align="right">Note: </td>
        <td><textarea name="note" cols="30" rows="5"></textarea></td>
    </tr>
  </table>  
  </td>
  <td width="50%" style="background-color:White;">
       <h4> Select building(s) to be Transfered : </h4>
        
        <%  
            if (pid = "") then
                strsql = "Select bldgNum, bldgName FROM buildings WHERE portfolioid = 0"
            else
                strsql = "Select bldgNum, bldgName FROM buildings WHERE portfolioid = " + pid
            end if
            rst1.open strsql, getConnect(pid, 0, "Billing")
            if not rst1.eof then
         %>
         <br />
         <% do until rst1.eof %>
         <input type="checkbox" name="<%=rst1("bldgNum")%>" /> &nbsp; <%=rst1("bldgName") %> <br />
         <% rst1.movenext
            loop
            rst1.close 
            
            end if%>
         
  </td>
</tr>
<tr>
    <td align="right">
        <input type="submit" name="submit" value="Submit" onclick="return checkForm();" /> &nbsp;
    </td>
    <td>
    &nbsp;<input type="reset" name="reset" value="Reset" />
    </td>
</tr>
</table>
<input type="hidden" name="buildingStr" value="<%=buildingStr%>"
</form>
</body>
</html>