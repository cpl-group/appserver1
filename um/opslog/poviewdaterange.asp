<%@Language="VBScript"%>
<%option explicit%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<script>
function po(id1){
	document.location="accpofilter.asp?id1="+id1
}

function printResults(){
	document.forms.datepicker.target='_blank';
	document.forms.datepicker.printview.value='yes';
	document.forms.datepicker.submit();
}

function viewResults(){
	document.forms.datepicker.target='results';
	document.forms.datepicker.printview.value='no';
}
</script>
<link rel="Stylesheet" href="../../gEnergy2_Intranet/styles.css" type="text/css">
</head>
<form name="datepicker" target="results" action="posearch.asp">
<input type="hidden" name="select" value="date">
<input type="hidden" name="findvar" value="var">
<input type="hidden" name="caller" value="viewdaterange">
<input type="hidden" name="printview" value="no">
<body bgcolor="#eeeeee" text="#000000">
<table border=0 cellpadding="6" cellspacing="0" width="100%" bgcolor="#eeeeee" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;border-right:1px solid #cccccc">
	<tr>
		<td>
			<a href="corppoview.asp" style="color:#333366;">Approve/Reject Submitted RFs</a> &nbsp;|&nbsp; 
			<a href="acctpoview.asp" style="color:#333366;">View Approved RFs</a> &nbsp;|&nbsp; 
			<b> View All Rfs </b>
		</td>
	</tr>
</table>


<table border=0 cellpadding="2" bgcolor="eeeeee" cellspacing="0" width="100%">
	<tr>
		<td width="18%" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;">
			From Date&nbsp;<input type="text" name="fromdate" value="<%=cdate(date()-30)%>" size="13%">
		</td>
		<td width="25%" align="left" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;border-right:1px solid #cccccc">
			To Date&nbsp;<input type="text" name="todate" value="<%=cdate(date())%>" size="13%">
		</td>
		<td align="left" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;border-left:1px solid #ffffff">
			&nbsp;&nbsp;
			<input type="submit" name="Submit" value="View" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;" 
				onclick="javascript:viewResults();">
		</td>
		<td align="right" style="border-top:1px solid #ffffff;border-bottom:1px solid #cccccc;border-right:1px solid #cccccc">
			<img src="/images/print.gif" onclick="javascript:printResults();" style="cursor:hand">
			<a href="javascript:printResults();" style="text-decoration:none;"><b>&nbsp;&nbsp;Print Results</b></a>
		</td>
	</tr>
</table>

<iframe name="results" src="" height="90%" width="100%" frameborder="0">
</form>