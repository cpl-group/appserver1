<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%

dim cnnSuper,sql,rst, rst2,strBldgnum,strUser,strip,strPw,strGateway,btnsave,btnUpdate,load,deviceName
dim gateway,login,ip',password
dim i, gatewayId

strGateway=request("Gateway")
strBldgnum = request("bldgnum")
strUser = request("Username")
strPw =request("Password")
strip=request("IP")
btnsave=request("submit")
'btnUpdate=request("update")'for other one
'load=request("select")
deviceName=request("deviceName")
gatewayId = request("id")


set cnnSuper = server.createobject("ADODB.Connection")
set rst	= server.createobject("ADODB.Recordset")
set rst2 = server.CreateObject("ADODB.Recordset")
cnnSuper.open getConnect(0,0,"dbCore")
%>

<script>
    function closewin() {
        window.close()
    }
    function checkform(frm) {
        var err = "";
        if (frm.Gateway.value == '') err += "Select Gateway name\n";
        //if (frm.Username.value == '') err += "No Username entered\n";
        //if (frm.Password.value == '') err += "No Password entered\n";
        if (frm.IP.value == '') err += "No IP entered\n";
        if (frm.deviceName == '') err += "No Device Name entered\n";

        if (err == "")
            return true;
        else alert(err);
        return false;
    }

    function New() {
        document.location.href = "Gatewaysetup.asp?select=Yes"
    }

    function rowClick(gateway, login, password, ip, devicename, bldgnum, id) {

        document.gatewaySetup.submit.value = "Update";

        //document.gatewaySetup.devicename.value = gateway;
        document.gatewaySetup.ip.value = ip;
        document.gatewaySetup.deviceName.value = devicename;
        document.gatewaySetup.Username.value = login;
        document.gatewaySetup.Password.value = password;
        document.gatewaySetup.Username.value = login;
        document.gatewaySetup.bldgnum.value = bldgnum;
        document.gatewaySetup.id.value = id;

        var i;

        for (i = 0; i < document.employeeForm.Gateway.length; i++) {
            if (document.employeeForm.Gateway[i].text == gateway) {
                document.employeeForm.Gateway.selectedIndex = i;
            }
        }
        
        //alert(gateway + login + password + ip + devicename);
        
        window.scrollTo(0, 0);
    }
</script>

<%
if btnsave = "Insert" then 
'dim sqal
sql="insert into rm (DeviceName,Gateway,login,password,bldgnum,ip,enable,lm,hostip) values('"&deviceName&"','"&strGateway&"','"&strUser&"','"&strPw&"','"&strBldgnum&"','"&strip&"','1','1','10.0.7.149'"&")"
'response.write sql
rst.open sql, cnnSuper
'rst.close
response.write "<script>"
response.write "alert('New Entry Entered');"
response.write "closewin();"
response.write "</script>"
'response.end
end if

if btnsave = "Update" then 
sql="update rm set gateway='"&strGateway&"',login='"&strUser&"',password='"&strPw&"',ip='"&strip&"',DeviceName='"&deviceName&"' where id="&gatewayId
'response.Write sql

rst.open sql, cnnSuper
'rst.close
response.write "<script>"
response.write "alert('Updated');"
response.write "closewin();"
response.write "</script>"
end if

'if load = "Yes" then 

'sql ="select Gateway,login,password,bldgnum,ip,DeviceName from rm where bldgnum='"&strBldgnum&"'"

'rst.open sql, cnnSuper
'if not rst.EOF then

'dim gateway,login,password,ip
'gateway=rst("gateway")
'login=rst("login")
'password=rst("password")
'strip=rst("ip")
'deviceName=rst("DeviceName")
'rst.close
'else


'load="no"
'gateway=""
'login=""
'password=""
'ip=""

'end if


'rst.close
'end if


sql ="select gateway from GatewayIndex order by gateway"
rst.open sql, cnnSuper
 
 
%>
<html>
<head>
    <title>Gateway Setup</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

    <script language="javascript" src="/genergy2/sorttable.js" type="text/javascript"></script>

</head>
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
<body bgcolor="#eeeeee">
    <form name="gatewaySetup" action="GatewaySetup.asp" onsubmit="return(checkform(this))">
    <table>
        <tr>
            <td colspan="2" width="49%" bgcolor="#6699cc" nowrap>
                <span class="standardheader">&nbsp;&nbsp;Utility Manager Gateway Setup </span>
            </td>
        </tr>
        <td align="center">
            Gateway
        </td>
        <td>
            <select name="Gateway">
                <%
					if not rst.eof then
						do until rst.eof
                %>
                <option value="<%=rst("gateway")%>" <%if gateway=rst("gateway") then%> selected <%end if%>>
                    <%=rst("gateway")%>
                </option>
                <%
							rst.movenext
						loop
					end if
					rst.close
					set rst= nothing
                %>
            </select><!--&nbsp;&nbsp;<a href="Gatewaysetup.asp?New=Yes">New</a>-->
        </td>
        <tr>
            <td align="center">
                Device Name :
            </td>
            <td>
                <input type="text" name="deviceName" value="<%=deviceName%>" />
            </td>
        </tr>
        <tr>
            <td align="center">
                Username:
            </td>
            <td>
                <input type="text" name="Username" value="<%=login%>">
            </td>
        </tr>
        <tr>
            <td align="center">
                Password:
            </td>
            <td>
                <input type="text" name="Password" value="<%=password%>">
            </td>
        </tr>
        <tr>
            <td align="center">
                IP:
            </td>
            <td>
                <input type="text" name="ip" value="<%=strip%>">
            </td>
        </tr>
        <%if load="Yes" then %>
        <td align="center">
            <input type="submit" name="update" value="Update">
        </td>
        <%else%>
        <td align="center">
            <input type="submit" name="submit" value="Insert">
        </td>
        <%end if%>
        <td>
            <input type="button" value="Cancel" onclick="closewin();">
        </td>
        <input type="hidden" name="bldgnum" value="<%=strBldgnum%>" >
        <input type="hidden" name="id" value="<%=gatewayId%>" />
    </table>
    </form>
    <%
        sql ="select id, Gateway,login,password,bldgnum,ip,DeviceName from rm where bldgnum='"&strBldgnum&"'"
        rst2.open sql, cnnSuper
    
     %>
    
    <div style="margin: 10px; overflow: auto">
        <table id="sortTable" class="sortable" style="font-size: 11px; font-family: Arial, Helvetica, sans-serif;"
            cellspacing="1" cellpadding="3" border="1" width="99%">
            <thead align="center">
                <% for i = 0 to rst2.fields.Count - 1 %>
                <th>
                    <a href="#">
                        <%=rst2.fields(i).Name%></a>
                </th>
                <%next%>
            </thead>
            <tbody align="center">
                <%do while not rst2.eof%>
                <tr onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = '#eeeeee'"
                    onclick="rowClick('<%=rst2.fields("gateway")%>','<%=rst2.fields("login")%>', '<%=rst2.fields("password")%>', '<%=rst2.fields("ip")%>', '<%=rst2.fields("devicename")%>', '<%=rst2.fields("bldgnum")%>', '<%=rst2.fields("id")%>') ">
                    <%for i = 0 to rst2.fields.Count - 1%>
                    <td style="border-bottom: 1px solid #CCCCCC">
                        <%=UCase(rst2(i))%>
                    </td>
                    <%next%>
                </tr>
                <%
                    rst2.movenext
                    loop%>
            </tbody>
        </table>
    </div>
</body>
<%
            rst2.close
            set rst2 = nothing
			set cnnSuper = nothing
 %>
</html>
