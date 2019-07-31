<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%

dim bldgNum, pid, bldgName, accTemplate, currTemplate, currTemplateId, isPostBack, useFile, acctCode, codeDesc
dim cnn, rst, strSql, strSql2, rst2, message, opType
dim useFileBit, isDirty

bldgNum = request("bldgNum")
pid = request("pid")
bldgName = request("bldgname")
accTemplate = request("accTemplate")
isDirty = request("dirty")
if (isDirty = "") then isDirty = false

if(request("useAcctFile") = "on") then
    useFileBit = 1
else
    useFileBit = 0
end if

set cnn = server.createobject("ADODB.Connection")
set rst	= server.createobject("ADODB.Recordset")
set rst2 = server.createobject("ADODB.Recordset")
cnn.open getConnect(0,0,"billing")

'do the following if user clicked on "Add Code" button
if (request("submit") = "Add Code") then
    acctCode = request("acctCodeType")
    codeDesc = request("codeDesc")
    
    strSql = "SELECT * FROM Bldg_AcctCodes WHERE TypeId = " + acctCode + " AND BldgNum = '" + bldgNum + "'"
    rst.open strSql, cnn
    
    if rst.eof then
        strSql = "INSERT INTO Bldg_AcctCodes (TypeId, BldgNum, Code) VALUES (" + acctCode + ", '" + bldgNum + "', '" + codeDesc + "')"
        opType = "Inserted accounting code"
    else
        strSql = "UPDATE Bldg_AcctCodes SET code = '" + codeDesc + "' WHERE TypeId = " + acctCode + " AND BldgNum = '" + bldgNum + "'"
        opType = "Updated accounting code"
    end if
    
    on error resume next
    cnn.Execute strSql
    
    if err<>0 then
        message = opType + " failed, please check your data and try again"
      else 
        message = opType + " Successful"
      end if 
    
else

    'do the following if user click on submit button

    strSql = "Select templateName, useFile, a.templateID, ProcedureName FROM acctFile_Template a INNER JOIN Bldg_AcctFile_Setup b ON a.templateID = b.templateID WHERE bldgNum = '" + bldgNum + "'"
    rst.open strSql, cnn

    if not rst.eof then 
        currTemplate = rst("templateName")
        useFile = rst("useFile")
        currTemplateId = rst("templateId")
    end if
    rst.close
            
    'if the form is dirty, do insert/update accounting file
    'if current template is null, do insert
    if (isDirty) then
        if (currTemplateId <> accTemplate ) then
            if (isNull(currTemplateId) OR currTemplateId = "") then
                strSql2 = "INSERT INTO Bldg_AcctFile_Setup([TemplateId], [UseFile], [BldgNum]) VALUES("+ accTemplate +  ", " + CStr(useFileBit) +  ", '" + bldgNum + "')"
                opType = "Inserted"
            else
                strSql2 = "Update Bldg_AcctFile_Setup SET [TemplateId] = " + accTemplate + ", " + "[UseFile] = " + CStr(useFileBit) + ", " + "[bldgNum] = '" + bldgNum + "' WHERE bldgNum = '" + bldgNum + "'"
                opType = "Updated"
            end if 
        end if

    'response.write("<h2> Accounting File Setup Page </h2>")
    'response.write("Pid : " + pid + "   " + " bldg : " + bldgNum + "   bldgName: " + bldgName + "<br />")
    'response.write(CStr(isDirty) + "<br />")
    'response.write(CStr(currTemplateId) + "_ accTemplate = " + acctemplate + "<br />")
    'response.write("useFile :" + CStr(useFileBit)) 
    'response.write("<br />")
    'response.write(strSql2)    
    'response.end    

        on error resume next
        cnn.Execute strSql2
      
          if err<>0 then
            message = opType + " failed, please check your data and try again"
          else 
            message = opType + " Successful"
          end if 
   
        rst2.close
   end if
   
end if

if (not isNull(message)) then
        response.write(message)
    end if        
%>
<script language="javascript" src="SetupPages.js" type="text/javascript"></script>
<html>
<head>
<title>Accounting File Setup</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<body bgcolor="#eeeeee">
<form name="acctSetupForm" action="accountingFileSetup.asp" onsubmit="return(checkform(this))">

<table>
<tr>

<td colspan ="2" width="100%" bgcolor="#6699cc" nowrap><span class="standardheader">&nbsp;&nbsp;Accounting File Setup for building # <%=bldgName%> </span></td>

</tr>
<%
    ' get current Accounting file for this building
strSql = "Select templateName, useFile, a.templateID, ProcedureName FROM acctFile_Template a INNER JOIN Bldg_AcctFile_Setup b ON a.templateID = b.templateID WHERE bldgNum = '" + bldgNum + "'"
rst.open strSql, cnn

if not rst.eof then 
    currTemplate = rst("templateName")
    useFile = rst("useFile")
    currTemplateId = rst("templateId")
end if
rst.close
 %>
<tr>
   
    <td align="right">
        Current accounting template is : 
    </td>
    <td> <input name="currTemplate" disabled="disabled" type="text" value='<%=currTemplate%>' /></td>
</tr>
<tr>
<td align="right" width="50%">Select a new accounting template: </td>
<td width="50%">
    <%
        'get the list of template name and stored procedure name to generate the drop down list
        
        strSql = "Select templateName, templateId FROM AcctFile_Template ORDER BY templateName"
        rst2.open strSql, cnn
     %>
				<select name="accTemplate" onchange="templateChanged()" >
					<%
					if not rst2.eof then
						do until rst2.eof
							%>
							<option value="<%=rst2("templateId")%>" <% if currTemplateId = rst2("templateId") then %> Selected <% end if %> > <%=rst2("templateName")%> </option>
							<%
							rst2.movenext
						loop
					end if
					rst2.close
					%>
				</select><!--&nbsp;&nbsp;<a href="Gatewaysetup.asp?New=Yes">New</a>-->
	    </td>
    </tr>
    <tr>
        <td align="right">Template Description : </td>
        <td ><textarea name="templateDesc" onkeydown="templateChanged()" ></textarea></td>
    </tr>
    <tr>
        <td>Use Accounting file for this building? </td>
        <td><input type="checkbox" onclick="templateChanged()" name="useAcctFile" <% if (useFile) then %> Checked <% end if %> /> </td>
    </tr>
    <tr>
		<td align="right"> <input type="submit" name="submit" value="Submit" ></td>
		<td><input type="button" value="Close" onclick="closewin();" /></td>
    </tr>
    </table>

	<input type="hidden" name="bldgNum" value='<%=bldgNum%>' />
    <input type="hidden" name="pid" value='<%=pid%>' />
    <input type="hidden" name="bldgName" value='<%=bldgName %>' />
    <input type="hidden" name="dirty" />


</td></tr>
</table>
<hr />
<div>
<br />
<span class="standardheader" style="background-color:#6699cc; width: 80%"> &nbsp;Accounting Codes for Building # <%=bldgNum %></span> <br />
<br />
<%
    strSql = "SELECT CodeId, CodeType FROM AcctFile_CodeType ORDER BY CodeType" 
    rst2.open strSql, cnn  
 %>
&nbsp; <select name="acctCodeType">
   <%
	if not rst2.eof then
		do until rst2.eof
			%>
			<option value="<%=rst2("CodeId")%>" > <%=rst2("CodeType")%> </option>
			<%
			rst2.movenext
		loop
	end if
	rst2.close
	%>
</select>&nbsp;
<input type="text" name="codeDesc" />&nbsp;
<input type="submit" name="submit" value="Add Code" onclick="return AcctCode(); acctSetupForm.submit(); " /> <br /><br />
    <%
        strSql = "SELECT CodeType, Code FROM Bldg_AcctCodes b INNER JOIN AcctFile_CodeType a ON b.TypeId = a.CodeId WHERE bldgNum = '" + bldgNum + "'"
        rst2.open strSql, cnn
        
        if not rst2.eof then 
     %>
        <table style="font-size: 11px; font-family: Arial, Helvetica, sans-serif;" cellspacing="1" cellpadding="3" border="1" width="99%">
          <tr>
            <th><%=rst2.fields(0).Name%></th><th><%=rst2.fields(1).Name%></th>
          </tr>
          <% while not rst2.eof  %>
          <tr>
            <td align="center"><%=rst2("CodeType")%></td><td align="center"><%=rst2("Code")%></td>
          </tr>
          <% rst2.movenext
		     wend 
        end if 
        rst2.close
        %>
        </table>
</div>
</form>
<%
    set cnn = nothing

 %>
</body>
</html>
