<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
			Response.Redirect "http://www.genergyonline.com"
		end if		
%>
<html>
<!-- #include file ="adovbs.inc" -->
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>New Page 1</title>
</head>

<body>

<%

bill = request("B")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set cmd = Server.CreateObject("ADODB.command")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.8;uid=genergy1;pwd=g1appg1;database=genergy1;"

cmd.ActiveConnection = cnn1
cmd.CommandText = "sp_lease_display"
cmd.CommandType = 4
Set prm1 = cmd.CreateParameter("billid", adChar, adParamInput, 10, bill)
cmd.Parameters.Append prm1

Set rst1 = cmd.Execute

%>

<form method="POST" action="lease_update.asp?b=<%=rst1("billingid")%>&bldg=<%=rst1("bldgnum")%>&ten=<%=rst1("tenantnum")%>" target="_self">
  <table border="0" width="100%" height="46">
    <tr>
      <td width="40%" bgcolor="#C0C0C0" align="center" height="13"><font size="2">Billing 
        Name</font></td>
      <td width="9%" bgcolor="#C0C0C0" align="center" height="13"> 
        <p align="center"><font size="2">Floor</font></td>
      <td width="5%" bgcolor="#C0C0C0" align="center" height="13"> 
        <p align="center"><font size="2">Sqft</font></td>
      <td width="5%" bgcolor="#C0C0C0" align="center" height="13"> 
        <p align="center"><font size="2">Tax Exempt</font></td>
      <td width="0%" bgcolor="#C0C0C0" align="center" height="13"> 
        <p align="center"><font size="2">Off Line</font></td>
    </tr>
    
    
    <tr>
      <td width="40%" height="25"><font size="1"> 
        <input type="text" name="billingname" size="50" value="<%=rst1("billingname")%>" style="font-size: 10px">
        </font></td>
      <td width="9%" height="25"><font size="2"> 
        <input type="text" name="flr" size="10" value="<%=rst1("flr")%>" style="font-size: 10px">
        </font></td>
      <td width="5%" height="25"><font size="2"> 
        <input type="text" name="sqft" size="5" value="<%=rst1("sqft")%>" style="font-size: 10px">
        </font></td>
        <% if rst1("taxexempt") then %>  
        
      <td width="5%" align="center" height="25"><font size="2"> 
        <input type="checkbox" name="taxexempt" value ="on" checked></font></td>
         <%else%>
       
      <td width="0%" align="center" height="25"><font size="2"> 
        <input type="checkbox" name="taxexempt" value="off" unchecked></font></td>
        <%end if%> 
        
        <% if rst1("leaseexpired") then %>
        
      <td width="29%" align="center" height="25"><font size="2"> 
        <input type="checkbox" name="leaseexpired" value ="on" checked></font></td>
       <%else%>
       
      <td width="12%" align="center" height="25"><font size="2"> 
        <input type="checkbox" name="leaseexpired" value="off" unchecked></font></td>
     <%end if%>  
    </tr>
    
  </table>
  <table border="0" width="100%">
    <tr>
      <td width="8%"><font size="2">Tenant #</font></td>
      <td width="19%"><font size="2"><input type="text" name="tenantNum" size="25" value="<%=rst1("tenantnum")%>" style="font-size: 10px"></font></td>
      <td width="42%"><font size="2">
        <input type="text" name="tname" size="50" value="<%=rst1("tname")%>" style="font-size: 10px">
        </font></td>
      <td width="34%"></td>
    </tr>
  </table>
  <table border="0" width="100%">
    <tr>
      <td width="6%"><font size="2">Address</font></td>
      <td width="13%"><font size="2"><input type="text" name="tstrt" size="25" value="<%=rst1("tstrt")%>" style="font-size: 10px"></font></td>
      <td width="2%"><font size="2"><input type="text" name="tcity" size="25" value="<%=rst1("tcity")%>" style="font-size: 10px"></font></td>
      <td width="3%"><font size="2"><input type="text" name="tstate" size="3" value="<%=rst1("tstate")%>" style="font-size: 10px"></font></td>
      <td width="9%"><font size="2"><input type="text" name="tzip" size="10" value="<%=rst1("tzip")%>" style="font-size: 10px"></font></td>
      <td width="79%"></td>
    </tr>
  </table>
 <table border="1" width="100%">
  <tr>
    <td width="10%" bgcolor="#C0C0C0"><p align="center"><font size="2">Admin Fee</font></td>
    <td width="10%" bgcolor="#C0C0C0"><p align="center"><font size="2">Tenant Rate</font></td>
    <td width="10%" bgcolor="#C0C0C0"><p align="center"><font size="2">Add On Fee</font></td>
    <td width="10%" bgcolor="#C0C0C0"><p align="center"><font size="2">Mod. Rate</font></td>
    <td width="10%" bgcolor="#C0C0C0"><p align="center"><font size="2">Coincident</font></td>
    <td width="10%" bgcolor="#C0C0C0"><p align="center"><font size="2">Profile</font></td>
    <td width="10%" bgcolor="#C0C0C0"><p align="center"><font size="2">Graph</font></td>
  </tr>
 
  <tr>
    <td width="10%"><p align="left"><font size="2"><input type="text" name="adminfee" size="5" value="<%=rst1("adminfee")%>" style="font-size: 10px"></font></td>
    <td width="10%"><p align="left"><font size="2"><input type="text" name="ratetenant" size="5" value="<%=rst1("ratetenant")%>" style="font-size: 10px"></font></td>
    <td width="10%"><p align="left"><font size="2"><input type="text" name="addonfee" size="5" value="<%=rst1("addonfee")%>" style="font-size: 10px"></font></td>
    <td width="10%"><p align="left"><font size="2"><input type="text" name="ratemodify" size="5" value="<%=rst1("ratemodify")%>" style="font-size: 10px"></font></td>
    
    <% if rst1("coincident") then %>   
    <td width="10%" align="center"><p><font size="2"><input type="checkbox" name="coincident" value="ON"></font></p></td>
    <% else %>
    <td width="10%" align="center"><p><font size="2"><input type="checkbox" name="loadprofile" value="Off" ></font></p></td>
    <% end if %>   
    
    <% if rst1("loadprofile") then %>   
    <td width="10%" align="center"><p><font size="2"><input type="checkbox" name="prtgraph" value="ON" checked></font></p></td>
    <% else %>
    <td width="10%" align="center"><p><font size="2"><input type="checkbox" name="C1" value="Off" ></font></p></td>
    <% end if %>
    
    <% if rst1("prtgraph") then %>   
    <td width="10%" align="center"><p><font size="2"><input type="checkbox" name="C1" value="ON" checked></font></p></td>
    <% else %>
    <td width="10%" align="center"><p><font size="2"><input type="checkbox" name="C1" value="Off" ></font></p></td>
    <% end if %>
  </tr>

</table>
  <p style="margin-top: 0; margin-bottom: 0">
    <input type="submit" value="Done" name="B1">
    <input type="reset" value="Reset" name="B2"></p>
</form>

</body>
<%
rst1.close
set cnn1 = nothing
%>
</html>
