<%option explicit%>
<!--#include file="adovbs.inc"-->
<%
'http params
dim crdate,j,c,avg,wip,tcost,OPENPO,link

j = request ("j")
c = request("c")

'adodb vars
dim cnn, cmd, rs, prm
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")


' open connection
cnn.Open "driver={SQL Server};server=10.0.7.20;uid=sa;pwd=!general!;database=main;"
cnn.CursorLocation = adUseClient

rs.open "SELECT crdate FROM sysobjects WHERE name = 'GY_activity_ara_status'",cnn
crdate=rs(0)
rs.close

rs.open "SELECT isnull(SUM(AMOUNT),0) FROM " + c + "_master_po WHERE closed=0 and job = '" + j + "'",cnn
openpo=rs(0)
'response.write "SELECT SUM(AMOUNT) FROM " + c + "_master_po WHERE closed=0 and job = '" + j + "'"
rs.close

' specify stored procedure 

cmd.CommandText = "sp_job_cost"
cmd.CommandType = adCmdStoredProc

Set prm = cmd.CreateParameter("c", adchar, adParamInput,2)
cmd.Parameters.Append prm

Set prm = cmd.CreateParameter("j", advarchar, adParamInput,9)
cmd.Parameters.Append prm

' assign internal name to stored procedure
cmd.Name = "test"
Set cmd.ActiveConnection = cnn

'return set to recordset rs
cnn.test c,j, rs

if rs.eof then
response.write "no record found"
else
if rs("jtd_labor_units") = 0 then
response.write "no hours/times posted within the account system"
else
avg = (rs("jtd_labor_cost") + rs("jtd_overhead_cost") + rs("jtd_other_cost") )/ rs("jtd_labor_units")
wip =  (rs("Revised_Contract_Amount") * rs("percent_complete")/100)-rs("jtd_work_billed")
tcost = ( cdbl(rs("hours")) * cdbl(avg)) + cdbl(rs("jtd_subcontract_cost"))+ cdbl(rs("jtd_material_cost"))
end if

%>
<html>
<head>
<script language="JavaScript1.2">
function po(c,j) {
	theURL="https://appserver1.genergy.com/um/war/po/po.asp?c="+c+"&j=" +j
	openwin(theURL,800,400)
}
function invoice(c,j) {
	theURL="https://appserver1.genergy.com/um/war/ara/invoice.asp?c="+c+"&j=" +j
	openwin(theURL,800,400)
}
function time(c,j) {
	theURL="https://appserver1.genergy.com/um/war/ts/ts.asp?c="+c+"&j=" +j
	openwin(theURL,800,400)
}

function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, scrollbars=yes, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}
</script>

<title>Genergy War Room - Job Cost Report</title>
</head>
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

td.red {color: red}
-->
</style>

<body text="#333333" link="#000000" vlink="#000000" alink="#000000" bgcolor="#FFFFFF">

<tr style="font-family:arial; font-size:12; color:black" bgcolor="#0099FF">
    
  <td style="font-size:10" rowspan="2" align="center"> 
 
    <table width="95%" border="1">
      <tr> 
        <td width="14%" bgcolor="#00FF00" bordercolor="#000000">Job #</td>
        <td width="28%" bgcolor="#FFFFFF"><%=rs("job")%></td>
        <td width="16%" bgcolor="#FFCC00" bordercolor="#000000">Original Contract</td>
 
        <td width="15%" bgcolor="#FFFFFF"> 
          <div align="right">
          <%=formatcurrency(rs("Original_contract_amount"),2)%></div></td>
        <td width="17%" bgcolor="#FFCC00">Amt. paid</td>
        <td width="10%" bgcolor="#FFFFFF"> 
          <div align="right">
          <%=formatcurrency(rs("jtd_payments"),2)%></div></td>
      </tr>
      <tr> 
        <td width="14%" bgcolor="#00FF00" bordercolor="#000000">Description</td>
        <td width="28%" bgcolor="#FFFFFF"><%=rs("description")%></td>
        <td width="16%" bgcolor="#FFCC00" bordercolor="#000000">Change Orders</td>
   
        <td width="15%" bgcolor="#FFFFFF"> 
          <div align="right"><%=formatcurrency(rs("JTD_Aprvd_Contract_Chgs"),2)%> 
          </div>
        </td>
        <td width="17%" bgcolor="#FFCC00">&nbsp;</td>
        <td width="10%" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>
      <tr> 
        <td width="14%" bgcolor="#00FF00" bordercolor="#000000">Customer</td>
        <td width="28%" bgcolor="#FFFFFF"><%=rs("address_1")%></td>
        <td width="16%" bgcolor="#FFCC00" bordercolor="#000000">Revised Contract</td>
        
        <td width="15%" bgcolor="#FFFFFF"> 
          <div align="right"><%=formatcurrency(rs("Revised_Contract_Amount"),2)%> 
          </div>
        </td>
        <td width="17%" bgcolor="#FFCC00">&nbsp;</td>
        <td width="10%" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>
      <tr> 
        <td width="14%" bgcolor="#00FF00" bordercolor="#000000">&nbsp;</td>
        <td width="28%" bgcolor="#FFFFFF"><%=rs("address_2")%></td>
        <td width="16%" bgcolor="#FFCC00" bordercolor="#000000">% Complete</td>
        
        <td width="15%" bgcolor="#FFFFFF"> 
          <div align="right"><%=formatnumber(rs("percent_complete"),0)%> 
          </div>
        </td>
        <td width="17%" bgcolor="#FFCC00">&nbsp;</td>
        <td width="10%" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>
      <tr> 
        <td width="14%" bgcolor="#00FF00" bordercolor="#000000">Project Manager</td>
        <td width="28%" bgcolor="#FFFFFF"><%=rs("project_manager")%></td>
        <td width="16%" bgcolor="#FFCC00" bordercolor="#000000">Job Value</td>
        
        <td width="15%" bgcolor="#FFFFFF"> 
          <div align="right"><%=formatcurrency(rs("Revised_Contract_Amount") * rs("percent_complete")/100,2) %> 
          </div>
        </td>
        <td width="17%" bgcolor="#FFCC00">&nbsp;</td>
        <td width="10%" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>
      <tr> 
        <td width="14%" bgcolor="#00FF00" bordercolor="#000000">Type</td>
        <td width="28%" bgcolor="#FFFFFF"><%=rs("type")%></td>
        <td width="16%" bgcolor="#FFCC00" bordercolor="#000000"><a href=<%="javascript:invoice('" & c & "','" & rs("job") & "')"%>>Amount Billed</a></td>
        <td width="15%" bgcolor="#FFFFFF"><div align="right"><%=formatcurrency(rs("JTD_work_billed"),2) %> 
          </div> </td>
        <td width="17%" bgcolor="#FFCC00">&nbsp;</td>
        <td width="10%" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>
      <tr> 
        <td width="14%" bgcolor="#00FF00" bordercolor="#000000">Status</td>
       <%
	     'If status is closed then mark it as red  KD 8/1/02
	     if rs("status") = "Closed" then 
		 Response.Write("<td width=28% bgcolor=#FFFFFF class=red>" & rs("status") & "</td>")
		 else
		 Response.Write("<td width=28% bgcolor=#FFFFFF>" & rs("status") & "</td>") 
		 end if 'End of if statement
		%>
		<td width="16%" bgcolor="#FFCC00" bordercolor="#000000">WIP</td>
        <td width="15%" bgcolor="#FFFFFF"><div align="right"><%=formatcurrency(wip,2)%> 
          </div> </td>
        <td width="17%" bgcolor="#FFCC00">&nbsp;</td>
        <td width="10%" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>
      <tr> 
        <td width="14%">&nbsp;</td>
        <td width="28%">&nbsp;</td>
        <td width="16%">&nbsp;</td>
        <td width="15%"> </td>
        <td width="17%">&nbsp;</td>
        <td width="10%">&nbsp;</td>
      </tr>
      <tr> 
        <td width="14%">&nbsp;</td>
        <td width="28%">&nbsp;</td>
        <td width="16%" bgcolor="#33FFCC" bordercolor="#000000"> 
          <div align="center"><b>JTD Costs</b></div>
        </td>
        <td width="15%" bgcolor="#FFFFFF"> </td>
        <td width="17%" bgcolor="#FFFFFF">&nbsp;</td>
        <td width="10%" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>
      <tr> 
        <td width="14%" bgcolor="#FF66FF" height="17"><a href=<%="javascript:time('" & c & "','" & rs("job") & "')"%>>Current Hours (time sheet)</a></td>
        <td width="28%" height="17"> 
          <div align="right">
          <%=formatnumber(rs("hours"),2)%></div></td>
        <td width="16%" bgcolor="#33FFCC" height="17">Labor Hours (timberline)</td>
		
        <td width="15%" bgcolor="#FFFFFF" height="17"> 
          <div align="right"><%=formatnumber(rs("jtd_labor_units"),2)%> 
          </div>
        </td>
        <td width="17%" bgcolor="#33FFCC" height="17">&nbsp;</td>
        <td width="10%" bgcolor="#FFFFFF" height="17"> 
          <div align="right"> </div>
        </td>
      </tr>
      <tr> 
        <td width="14%" bgcolor="#FF66FF">&nbsp;</td>
        <td width="28%">&nbsp;</td>
        <td width="16%" bgcolor="#33FFCC">&nbsp;</td>
		
        <td width="15%" bgcolor="#FFFFFF"> 
          <div align="right"> </div>
        </td>
        <td width="17%" bgcolor="#33FFCC">&nbsp;</td>
        <td width="10%" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>
      <tr> 
        <td width="14%" bgcolor="#FF66FF">&nbsp;</td>
        <td width="28%">&nbsp;</td>
        <td width="16%" bgcolor="#33FFCC">&nbsp;</td>
		
        <td width="15%" bgcolor="#FFFFFF"> 
          <div align="right"></div>
        </td>
        <td width="17%" bgcolor="#33FFCC">&nbsp;</td>
        <td width="10%" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>
      <tr> 
        <td width="14%" bgcolor="#FF66FF">&nbsp;</td>
        <td width="28%">&nbsp;</td>
        <td width="16%" bgcolor="#33FFCC">&nbsp;</td>
		
        <td width="15%" bgcolor="#FFFFFF"> 
          <div align="right"></div>
        </td>
        <td width="17%" bgcolor="#33FFCC">&nbsp;</td>
        <td width="10%" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>
      <tr> 
        <td width="14%" bgcolor="#FF66FF">&nbsp;</td>
        <td width="28%">&nbsp;</td>
			
        <td width="16%" bgcolor="#33FFCC"><a href=<%="javascript:po('" & c & "','" & rs("job") & "')"%>>Material Cost</a></td>
		
        <td width="15%" bgcolor="#FFFFFF"> 
          <div align="right"><%=formatcurrency(rs("jtd_material_cost"),2)%> 
          </div>
        </td>
        <td width="17%" bgcolor="#33FFCC">&nbsp;</td>
        <td width="10%" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>
      <tr> 
        <td width="14%" bgcolor="#FF66FF">&nbsp;</td>
        <td width="28%">&nbsp;</td>
        <td width="16%" bgcolor="#33FFCC">Open PO's</td>
        <td width="15%" bgcolor="#FFFFFF"><div align="right"><%=formatcurrency(openpo,2)%> 
          </div> </td>
        <td width="17%" bgcolor="#33FFCC">&nbsp;</td>
        <td width="10%" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>
      <tr> 
        <td width="14%" bgcolor="#FF66FF">&nbsp;</td>
        <td width="28%">&nbsp;</td>
        <td width="16%" bgcolor="#33FFCC">Subcontractor Cost</td>
		
        <td width="15%" bgcolor="#FFFFFF"> 
          <div align="right"><%=formatcurrency(rs("jtd_subcontract_cost"),2)%> 
          </div>
        </td>
        <td width="17%" bgcolor="#33FFCC">&nbsp;</td>
        <td width="10%" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>
      <tr> 
        <td width="14%" bgcolor="#FF66FF">&nbsp;</td>
        <td width="28%">
          <div align="right"></div>
        </td>
        <td width="16%" bgcolor="#33FFCC">&nbsp;</td>
		
        <td width="15%" bgcolor="#FFFFFF"> 
          <div align="right"></div>
        </td>
        <td width="17%" bgcolor="#33FFCC">&nbsp;</td>
        <td width="10%" bgcolor="#FFFFFF">&nbsp;</td>
      </tr>
      <tr> 
        <td width="14%">&nbsp;</td>
        <td width="28%">&nbsp;</td>
        <td width="16%">&nbsp;</td>
        <td width="15%"> </td>
        <td width="17%">&nbsp;</td>
        <td width="10%">&nbsp;</td>
      </tr>
      <tr> 
        <td width="14%">&nbsp;</td>
        <td width="28%">&nbsp;</td>
        <td width="16%">&nbsp;</td>
        <td width="15%"> </td>
        <td width="17%">&nbsp;</td>
        <td width="10%">&nbsp;</td>
      </tr>
      <tr> 
        <td width="14%">&nbsp;</td>
        <td width="28%">&nbsp;</td>
        <td width="16%">&nbsp;</td>
        <td width="15%"> </td>
        <td width="17%">&nbsp;</td>
        <td width="10%">&nbsp;</td>
      </tr>
    </table>
      
       
	<table width="99%" border="0" dwcopytype="CopyTableRow">
      <tr> 
        <td height="23" width="35%"> 
          <div align="left"><font size="2">update as of <%=formatdatetime(crdate,0)%></font></div>
        </td>
        <td height="23" width="35%">&nbsp;</td>
        <td height="23" width="6%"> 
          <div align="right"></div>
        </td>
        <td height="23" width="24%">&nbsp;</td>
      </tr>
    </table>
	
	
      <tr> 
        <td>
          <div align="right"></div>
  </td>
      </tr> 
	  <%
	  set cnn = nothing 
	  end if %>
	 
</body>
</html>