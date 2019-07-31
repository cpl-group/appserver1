<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<% 
'2/20/2008 N.Ambo amended email messages being sent and changed the email addresses to be notified to project managers of the job rather than account managers

dim jobid, sql, rs, emails, company, invoiceid
set rs = server.createobject("adodb.recordset")
jobid = request("jobid")
if isnumeric(jobid) then jobid = clng(jobid) else jobid = 0
invoiceid = request("invoiceid")
company = request("company")
emails = ""


sql = "SELECT * FROM ("&_
	"SELECT distinct email as emailadd FROM managers m, ["&application("CoreIP")&"].dbCore.dbo.ADusers_GenergyUsers u WHERE isnull(email,'')<>'' and m.userid=u.username and m.mid = (SELECT Project_Manager FROM MASTER_JOB WHERE id='"&jobid&"') "&_
	"UNION ALL "&_
	"SELECT distinct email as emailadd FROM ["&application("CoreIP")&"].dbCore.dbo.master_notes m,  ["&application("CoreIP")&"].dbCore.dbo.ADusers_GenergyUsers u WHERE u.username=m.uid and notefortype = 'arinvoice' and notefor='"&invoiceid&"'"&_
	") l where isnull(emailadd,'') <> ''"

rs.open sql, getConnect(0,0,"Intranet")
do until rs.eof
	emails = emails & "," & rs("emailadd")
	rs.movenext
loop
emails = mid(emails,2)
rs.close

dim message, message2
if trim(emails)<>"" then 
	message = "This is an invoice update request, please update the invoice notes for invoice "&invoiceid&", of job "&jobid
	message2 = "Your request has been sent."
else
	emails = "FinancialServices@Genergy.com"
	'2/20/2008 N.Ambo replaced with more meaningful message
	'message = "This is an invoice update request. However, current contacts are unavailable because they are no longer in the system. Please update the Account Manager setting for the customer under this job."

	message = "An update for invoice notes was requested for this invoice . However, this request could not be generated because no project manager is currently listed for this job or the Project Manager currently listed is no longer a user in our system. Please edit this job to list the proper Project Manager then a new update request should be made."

	message2 =  "Your request for an update could not be completed. Financial Services has been notified of the problem. Please contact someone in financial services regarding this request, then a new request should be made."
end if

sendmail emails, "GSA", "Update Request for Invoice notes for invoice "&invoiceid&", of job "&jobid, message 
response.Write(message2)
%>
<html>
<head>
	<title>Update Request</title>
</head>
<link rel="Stylesheet" href="/genergy2/setup/setup.css" type="text/css">   
<body onload="window.focus();<%'window.close();%>">
</body>
</html>
