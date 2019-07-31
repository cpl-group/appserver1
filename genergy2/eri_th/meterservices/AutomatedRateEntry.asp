<%@ Language=VBScript %>
<%option explicit%>

<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!--  
    METADATA  
    TYPE="typelib"  
    UUID="CD000000-8B95-11D1-82DB-00C04FB1625D"  
    NAME="CDO for Windows 2000 Library"  
-->

<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%>
<!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"-->
<%end if

	function getNumber(number)
	'	response.write "|"&number&"|"
		if not(isNumeric(number)) then number = 0
		getNumber = number
	end function

	Dim  Billperiod, building, Billyear, PortFolioId, UtilityId, utilitydisplay, rpt, pdf, Genergy_Users, demo, sql,monthdescription
	' Set Parameters
	building = request("bldgNum")	
	
	Dim rst1, rst2, cnn1
	set rst1 = server.createobject("ADODB.Recordset") 	
%>
<html>
<head>
<title>Automated Rate Entry</title>

<style type="text/css">
INPUT#f9 {
	font-size:9
}
</style>
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">
</head>
<body>
<table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
   <form name="form1" action="AutomatedRateEntry.asp">
    <tr bgcolor="#eeeeee"> 
      <td colspan="2" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;"> 
        <table border=0 cellpadding="3" cellspacing="0">
          <tr> 
        
								           
			
            <td> <select name="billyear" onclick="loadPeriod()">
                <option value="">Select Year</option>
                <%
                	sql = "SELECT distinct billyear " & _
						" FROM billyrperiod where billyear > 2014 and billyear < 2020" & _
				        " order by billyear desc "
				        
					rst1.open sql, getLocalConnect(building)
					do until rst1.eof%>
					<option value="<%=rst1("billyear")%>"<%if trim(rst1("billyear"))=trim(billyear) then response.write " SELECTED"%>><%=rst1("billyear")%></option>
					<%
						
							rst1.movenext
					loop
					rst1.close
					%>
					</select> </td>
					
	  			
					<td> <select name="monthdescription">
					 <option value="">Select Month</option>
                <%
                
				sql = "SELECT monthdescription" & _
						" FROM tblmonths" & _
				        " order by monthnumber "
					
				rst1.open sql, getLocalConnect(building)
				do until rst1.eof
				%>
					<option value="<%=rst1("monthdescription")%>" <%if trim(rst1("monthdescription"))=monthdescription then response.write " SELECTED"%>><%=rst1("monthdescription")%></option>
                <%
				  rst1.movenext
				loop
				rst1.close
				%>
              </select> </td>
              
				<td>
				  <input type="Submit" name="Generate Report" value="Generate Report"> 
            </td>
          </tr>
        </table></td>
        </form>
	</tr>
</table>
