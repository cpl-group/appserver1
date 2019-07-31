<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
server.ScriptTimeout=300
billyear 	= request("billyear")
billperiod 	= request("billperiod")
bldgnum 	= request("bldgnum")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rs	 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"Billing")

strsql = "exec sp_select_trip_report_bldg "&billyear&", "&billperiod&", '"&bldgnum&"'"

rst1.Open strsql, cnn1, 0, 1, 1
%>
<html>
<head>

<title>Trip Sheet Report</title>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<style type="text/css" media="print">
.tblunderline { border-bottom:1px solid #cccccc; }
</style>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"><head>
<script>

</script>
<%if not rst1.eof then%>
<body bgcolor="#eeeeee" leftmargin="0" topmargin="0" onload="window.print()">
<table border=0 cellpadding="3" cellspacing="0" bgcolor="#ffffff" width="100%">
<tr align="center" valign="top" bgcolor="#ffffff"> 
		    <td align="left" nowrap bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Bill Period</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Portfolio Name</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Building Name</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Address</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Trip Code</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Trip Date</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Tenant</td>
								
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Meter Name</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Serial Number</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Meter Make</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Meter Model</td>
		    
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Meter commissioing date</td>
								
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Commissioing report link</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Meter Picture Link</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Meter Manual link</td>
						    							
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Description of load monitored</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Load Traced/Verified (yes/no)</td>
		    
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Load Tracing/Verification date</td>
		    
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Load Tracing/Verification Technician</td>
		    
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Load Tracing/Verification Job #</td>
								
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">CT Ratio</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Meter Voltage</td>
		    
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Remote Read Factor</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Usage Multiplier</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Capicity Multiplier</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Floor</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Location</td>
		    
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Manual Meter</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Previous Month Reading</td>
								
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Previous Month Usage</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Two Months Ago Usage</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Three Months Ago Usage</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Avg 3 Month Usage</td>
								
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">One Year Ago Usage</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Previous Month Peak</td>
		    
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Two Months Ago Peak</td>
		    
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Three Months Ago Peak</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Avg 3 Month Peak</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">One Year Ago Peak</td>
				
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">New Reading</td>
		    
		    <td bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;">Peak New</td>
				
		    
				
		    
		    </tr>
</table>
<table border=0 cellpadding="3" cellspacing="0" bgcolor="#ffffff" width="100%">
  <% 
	Do until rst1.EOF 
		
			%>
  
   		   
		    
		
		    <tr <% %> bgcolor="#ffffff" valign="top" > 
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;" nowrap><%=rst1("billperiod")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("portfolioname")%>&nbsp;</td>
			  <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("bldgname")%>&nbsp;</td>
			  <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("address")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("tripcode")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("tripdate")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("tenant")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("Meter Name")%>&nbsp;</td>
		      		      
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;" nowrap><%=rst1("serialnumber")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("metermake")%>&nbsp;</td>
			  <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("metermodel")%>&nbsp;</td>
			  
			   <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("Metercommissioingdate")%>&nbsp;</td>
			  <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("Commissioingreportlink")%>&nbsp;</td>
			  <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("MeterPictureLink")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("MeterManuallink")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("Descriptionofloadmonitored")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("LoadTracedVerifiedyesno")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("LoadTracingVerificationdate")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("LoadTracingVerificationTech")%>&nbsp;</td>	
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("LoadTracingVerificationJob")%>&nbsp;</td>			  		  
			  
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("ctratio")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("metervoltage")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("remote read factor")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("usagemultiplier")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("capacitymultiplier")%>&nbsp;</td>
		      
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;" nowrap><%=rst1("floor")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("location")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("manualmeter")%>&nbsp;</td>
			  <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("PrevMonthReading")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("prevmonthusage")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("twomonthsagousage")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("threemonthsagousage")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("Avg3monusage")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("Oneyragousage")%>&nbsp;</td>
		      
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;" nowrap><%=rst1("prevmonthpeak")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("2monthagopeak")%>&nbsp;</td>
			  <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("3monthagopeak")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("Avg3monthpeak")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("1yearagopeak")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("newreading")%>&nbsp;</td>
		      <td height="40" bgcolor="#eeefff" class="tblunderline" style="border-left:1px solid #e3e3d3;"><%=rst1("peaknew")%>&nbsp;</td>
		      		      		      
		    </tr>
		    
		  
        <div style="page-break-before:always" />
        <%
		rst1.movenext
    loop
%>

</table>
</body>
<%else %>
<body bgcolor="#eeeeee" leftmargin="0" topmargin="0" class="innerbody">
	<table width="100%" border="0" cellspacing="0" cellpadding="2">
  <tr>
    <td class="standardheader" bgcolor="#999999" align="center">NO TRIP SHEET DATA FOUND FOR BILL YEAR AND BILL PERIOD</td>
  </tr>
</table>

</body>
<% end if %>
</html>
