<%Option Explicit%>
<html>
<head>
<title>Import Lefrak Water Bill Adjustments</title>
</head>
<%
	dim pid, bldg, customsrc, action, tid, lid, id, byear, bperiod, creditid, credit_adj, utilid
	pid = request("pid")
	bperiod = request("bperiod")
	byear = request("byear")
	bldg = request("bldgnum")
	tid = request("tid")
	lid = request("lid")
	utilid = request("utilid")
%>
<body>
	<form action="importlefrakadjustments2.asp?bldg=<%=bldg%>&byear=<%=byear%>&bperiod=<%=bperiod%>&utilid=<%=utilid%>" method="post" enctype="multipart/form-data" name="frmMain">
	Upload 
	<input name="file1" type="file">
	<input type="submit" name="Submit" value="Submit">
	<input type="hidden" value="<%= bperiod %>" name="bperiod"/>
	<input type="hidden" value="<%= byear %>" name="byear"/>
	<input type="hidden" value="<%= utilid %>" name="utilid"/>
	<input type="hidden" value="<%= bldg %>" name="bldg"/>
	
	</form>
</body>
</html>
<!--- This file download from www.shotdev.com -->