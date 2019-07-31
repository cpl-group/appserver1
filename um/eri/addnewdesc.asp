<%@Language="VBScript"%>
<!-- #include file="./adovbs.inc" -->

<html>
<head>

<title>Add Description</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function setValue(tenant_no, surveyid){
 
	var temp="survey_detail.asp?tenant_no="+tenant_no+"&surveyid="+surveyid
	parent.frames.tenant.location = temp
}

function setValue(count){
    var amps=document.form1.amps.value
	var volts=document.form1.volts.value
	var ph=document.form1.ph.value
	var pf=document.form1.pf.value
	var watt=amps*volts*ph*pf
    document.form1.watt.value=watt
}

</script>
</head>
<% 
if (Request.QueryString("type") = "Equipment") then
	ph=1.732
else
	ph=1
end if
pf=0.85
%>	 		
<body bgcolor="#FFFFFF" text="#000000">
<form name="form1" method="post" action="adddescription.asp">
  <table width="100%" >
    <tr> 
      <td width="20%" ><font face="Arial, Helvetica, sans-serif"><u><b>Add Description</b></u></font></td>
      <td width="20%" >&nbsp;</td>
    </tr>
    <tr> 
      <td width="20%" ><i><font face="Arial, Helvetica, sans-serif"><b>Type</b></font></i></td>
    </tr>
    <tr> 
      <input type="hidden" name="type" value="<%=Request.QueryString("type")%>">
	  <input type="hidden" name="count" value="<%=Request.QueryString("count")%>">
      <td width="20%"> <font face="Arial, Helvetica, sans-serif"><%=Request.QueryString("type")%></font> 
      </td>
      <td width="20%"> <font face="Arial, Helvetica, sans-serif">&nbsp; </font></td>
    </tr>
    <tr> 
      <td width="20%"><i><font face="Arial, Helvetica, sans-serif"><b>Description</b></font></i></td>
      <td width="20%"><i></i></td>
    </tr>
    <tr> 
      <td width="20%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="description" value="">
        </font></td>
      <td width="20%"><font face="Arial, Helvetica, sans-serif">&nbsp; </font></td>
    </tr>
    <tr> 
      <td width="20%"><i><font face="Arial, Helvetica, sans-serif"><b>Amps</b></font></i>
      <td width="20%"><i></i></td>
    </tr>
    <tr> 
      <td width="20%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="amps" value="" onChange=setValue()>
        </font></td>
      <td width="20%"><font face="Arial, Helvetica, sans-serif">&nbsp; </font></td>
    </tr>
    <tr> 
      <td width="20%"><i><font face="Arial, Helvetica, sans-serif"><b>Volts</b></font></i>
      <td width="20%"><i></i></td>
    </tr>
    <tr> 
      <td width="20%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="volts" value="" onChange=setValue()>
        </font></td>
      <td width="20%"><font face="Arial, Helvetica, sans-serif">&nbsp; </font></td>
    </tr>
    <tr> 
      <td width="20%"><i><font face="Arial, Helvetica, sans-serif"><b>Ph</b></font></i>
      <td width="20%"><i></i></td>
    </tr>
    <tr> 
      <td width="20%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="ph" value="<%=ph%>" onChange=setValue()>
        </font></td>
      <td width="20%"><font face="Arial, Helvetica, sans-serif">&nbsp; </font></td>
    </tr>
    <tr> 
      <td width="20%"><i><font face="Arial, Helvetica, sans-serif"><b>Pf</b></font></i>
      <td width="20%"><i></i></td>
    </tr>
    <tr> 
      <td width="20%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="pf" value="<%=pf%>" onChange=setValue()>
        </font></td>
      <td width="20%"><font face="Arial, Helvetica, sans-serif">&nbsp; </font></td>
    </tr>
    <tr> 
      <td width="20%"><i><font face="Arial, Helvetica, sans-serif"><b>Watt</b></font></i>
      <td width="20%"><i></i></td>
    </tr>
    <tr> 
      <td width="20%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="watt" value="">
        </font></td>
      <td width="20%"><font face="Arial, Helvetica, sans-serif">&nbsp; </font></td>
    </tr>
    <tr>
      <td width="20%"><i><font face="Arial, Helvetica, sans-serif"><b>AdjFactor</b></font></i>
      <td width="20%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="20%"><font face="Arial, Helvetica, sans-serif">
        <input type="text" name="adjfactor" value="">
        </font>
      <td width="20%">&nbsp;</td>
    </tr>
    <tr> 
      <td width="20%"><i><font face="Arial, Helvetica, sans-serif"><b>MonthFactor</b></font></i>
      <td width="20%"><i></i></td>
    </tr>
    <tr> 
      <td width="20%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="monthfactor" value="">
        </font></td>
      <td width="20%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="Submit" name="Submit" value="Save">
        <input type="button" name="Cancel"  value="Cancel" onclick="Javascript:window.close()">
        </font></td>
    </tr>
  </table>
</form>
</body>
</html>

