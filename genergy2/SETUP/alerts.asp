<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim bldgnum, rst1, rstE, rstNo, cnn1, strsqlE, strsqlNo, strsql1, cmd, prm
dim eAlertId, eScheduleId, Eemail, eActive, estimated, action, hasEAlert, hasNoAlert
dim noAlertId, noScheduleId, nEmail, noActive, noUsage, errorcode

bldgnum = request("bldgNum")
action = request("save")

set cnn1 = server.CreateObject("ADODB.connection")
set rstE = server.CreateObject("ADODB.recordset")
set rstNo = server.CreateObject("ADODB.recordset")
set rst1 = server.CreateObject("ADODB.recordset")
set cmd = server.CreateObject("ADODB.command")

cnn1.open getConnect(0,bldgnum,"billing")

if (action = "Save") then

    eScheduleId = request("estimated_schedule")
    eAlertId = request("eAlertId")
    Eemail = request("estimated_contacts")
    estimated = request("estimated")
    
    if (estimated = "on") then
        estimated = true
    else
        estimated = false
    end if
    
    if (eScheduleId <> "" OR eScheduleId <> null) then
        eScheduleId = cint(eScheduleId)
    end if

    noScheduleId = request("noUsage_schedule")
    noAlertId = request("noAlertId")
    nEmail = request("noUsage_contacts")
    noUsage = request("noUsage")
    
    if (noUsage = "on") then
        noUsage = true
    else
        noUsage = false
    end if
    
    if (noScheduleId <> "" OR noScheduleId <> null) then
        noScheduleId = cint(noScheduleId)
    end if

else
   
    strsqlE = "Select * from BuildingAlert where BldgNum='"&bldgnum&"' AND typeId = 1"
    strsqlNo = "Select * from BuildingAlert where BldgNum='"&bldgnum&"' AND typeId = 2"
    
    rstE.open strsqlE, cnn1
    
    if (NOT rstE.eof) then
        eScheduleId = rstE("ScheduleId")
        eAlertId = rstE("AlertId")
        Eemail = rstE("email")
        estimated = rstE("Active")
    else
        eAlertId = 0
        estimated = false
    end if
    
    rstE.close
    
    rstNo.open strsqlNo, cnn1
    
    if (NOT rstNO.eof) then
        noScheduleId = rstNo("ScheduleId")
        noAlertId = rstNo("AlertId")
        nEmail = rstNo("email")
        noUsage = rstNo("Active")
    else
        noAlertId = 0
        noUsage = false
    end if
    
    rstNo.close
   
end if

'response.Write("eScheduleId : " + cstr(eScheduleId) + "<br />")
'response.Write("eAlertId : " + cstr(eAlertId) + "<br />")
'response.Write("Eemail : " + Eemail + "<br />")
'response.Write("estimated : " + cstr(estimated) + "<br />")
'response.Write("noScheduleId : " + cstr(noScheduleId) + "<br />")
'response.Write("noAlertId : " + cstr(noAlertId) + "<br />")
'response.Write("nEmail : " + nEmail + "<br />")
'response.Write("noUsage : " + cstr(noUsage) + "<br />")

if (action = "Save") then
    
    cnn1.CursorLocation = adUseClient
    cmd.activeConnection = getConnect(0,bldgnum,"billing")
    cmd.CommandText =  "sp_Save_BuildingAlert"
    cmd.CommandType = adCmdStoredProc
    
    Set prm = cmd.CreateParameter("AlertId", adInteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("BldgNum", adVarChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("TypeId", adInteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("ScheduledId", adInteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("Email", adVarChar, adParamInput, 75)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("Active", adBoolean, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("out", adBoolean, adParamOutput)
    cmd.Parameters.Append prm
      
    cmd.Parameters("BldgNum") = bldgnum
    'Estimated
    if (lcase(estimated) = "true" OR eAlertId <> 0) then
        response.Write("in esimated <br />")
        cmd.Parameters("AlertId") = eAlertId
        cmd.Parameters("TypeId") = 1
        cmd.Parameters("ScheduledId") = eScheduleId
        cmd.Parameters("Email") = Eemail
        cmd.Parameters("Active") = estimated
        cmd.Parameters("out") = errorcode
        cmd.execute()
        
        if (errorcode = 0) then
            response.Write("Estimated Alert saved <br />")
        else
            response.Write("Error saving Estimated alert <br />")
        end if
    end if
    
    if (lcase(noUsage) = "true" OR noAlertId <> 0) then
        'NoUsage
        response.Write("in noUsage <br />")
        cmd.Parameters("AlertId") = noAlertId
        cmd.Parameters("TypeId") = 2
        cmd.Parameters("ScheduledId") = noScheduleId
        cmd.Parameters("Email") = nEmail
        cmd.Parameters("Active") = noUsage
        cmd.Parameters("out") = errorcode
        
        cmd.execute()
        
        if (errorcode = 0) then
            response.Write("NoUsage Alert saved <br />")
        else
            response.Write("Error saving NoUsage alert <br />")
        end if
    
    end if
    
end if 

 %>
<html>
<head>
    <title>Untitled Page</title>

<script type="text/javascript" language="javascript">
    function checkForm() {

        var returnval = false;
        var message = "";

        //if (document.alertForm.estimated.checked) {
         //   alert("in Estimated checked");
         //   returnval = false;
        //}
        
      //  if (document.form1.noUsage.checked) {
      //      alert("in noUsage checked");

        //  }

        return returnval;
    }
</script>
</head>
<body>
    <form id="alertForm" action="alerts.asp" method="post" >
    <input type="hidden" name="bldgNum" value="<%=bldgnum%>" />
    <div style="text-align:center">
        <p>Alert Status for <%=bldgnum%></p>
 
        <p>Notify when: </p>
        <table>
            <thead>
                <tr>
                    <th>Type of Alert</th><th>Schedule</th><th>Contacts (Emails)</th>
                </tr>
            </thead>
            <tr>
                <td><input type="checkbox" value="estimated" name="estimated" <% if (estimated = true) then %> checked <% end if %> /> Data is estimated </td>
                <td>
                    <%
                        strsql1 = "select * from BldgAlertSchedule"
                        rst1.open strsql1, cnn1
                     %>
                <select name="estimated_schedule" >
                    <option value="" selected="selected" >Select a Interval</option>
                    <% do until rst1.eof
						%>
						<option value="<%=rst1("ScheduleId")%>" <% if(eScheduleId = rst1("ScheduleId")) then %> selected <% end if %> > <%=rst1("ScheduleType")%></option>
						<%
						rst1.movenext
					    loop
					%>
                </select></td>
                <% rst1.close %>
                <td>
                    <input type="text" name="estimated_contacts" value="<%=Eemail%>" style="width: 300px" />
                    <input type="hidden" name="eAlertId" value="<%=eAlertId%>" />
                </td>
            </tr>
            <tr>
                    <%
                        strsql1 = "select * from BldgAlertSchedule"
                        rst1.open strsql1, cnn1

                     %>
                <td><input type="checkbox" value="noUsage" name="noUsage" <% if (noUsage = true) then %> checked <% end if %> /> No Usage reported </td>
                <td><select name="noUsage_schedule" >
                    <option value="" selected="selected">Select a Interval</option>
                    <% do until rst1.eof %>
                        <option value="<%=rst1("ScheduleId")%>" <% if (noScheduleId = rst1("ScheduleId")) then %> selected <% end if %> ><%=rst1("ScheduleType")%></option>
                    <% rst1.movenext
                       loop 
                    %>
                </select></td>
                <% rst1.close %>
                <td>
                    <input type="text" name="noUsage_contacts" value="<%=nEmail%>" style="width: 300px" />
                    <input type="hidden" name="noAlertId" value="<%=noAlertId%>" />         
                </td>
            </tr>
            <tr>
                <td colspan="3" align="right">*for multiple emails, use comma (,) to separate each email</td>
            </tr>
            <tr>
            <td colspan="3" align="right">
                <input type="submit" value="Save" name="save" />&nbsp;&nbsp;
                <input type="reset" value="Reset" name="reset" />
            </td>
            </tr>
        </table>
    </div>
    </form>
</body>
</html>
