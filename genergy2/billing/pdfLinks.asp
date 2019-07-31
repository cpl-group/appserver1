<%
Option Explicit
Response.Buffer = False%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/nonsecure.inc"-->
<%
dim leaseid, ypid, building, pid, byear, bperiod, logo
dim genergy2, devIP, utilityid, detailed, meterbreakdown
Dim SJPproperties, summaryusage, summarydemand, buildpdf
Dim textheader,msg,appTimeout,demo, billid,billurl
dim billCount, logow, logoh,rst11, postback

leaseid = trim(Request("l"))
ypid = trim(request("y"))
building = trim(request("building"))
utilityid = trim(request("utilityid"))
SJPproperties = trim(request("SJPproperties"))
Summaryusage = trim(request("summaryusage"))
Summarydemand = trim(request("summarydemand"))
pid = trim(request("pid"))
byear = trim(request("byear"))
bperiod = trim(request("bperiod"))
logo = trim(request("logo"))
detailed = trim(request("detailed"))
genergy2 = request("genergy2")
meterbreakdown = request("meterbreakdown")
textheader = trim(request("textheader"))
demo = request("demo")
billid = request("billid")
billurl = request("billurl")
billCount = request("billCount")
logow=request("logow")
logoh=request("logoh")
postback = request("blnPostBack")

dim dname : dname = ""
dim SecondsToWait : SecondsToWait = 10
dim StartTime : StartTime = Time()
dim timeunit : timeunit = "sec"

dim maxPageCount
Dim timeToGenerate
Dim PdfRequest
Dim totalPageCount
dim cnn1, rst1, rst2, sql
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

dim utilname
cnn1.Open getLocalConnect(Replace(building,"+"," "))
rst1.open "select utility from tblutility where utilityid = " & utilityid, cnn1
if not rst1.eof then utilname = rst1("utility")
rst1.close

Dim portfolio, cbldg
if building<>"" then
	cbldg = Replace(building, "+", " ")
	cbldg = Replace(cbldg, "%20", " ")
	rst1.open "SELECT location, b.bldgnum, b.portfolioid,billurl,logo, logoh, logow, p.portfolio FROM buildings b, portfolio p, billtemplates bt WHERE b.portfolioid=p.id AND bt.id=p.templateid AND bldgnum='"&cbldg&"'	", cnn1
		if not rst1.eof then 
			pid = rst1("portfolioid")
			billurl = rst1("billurl")
			logo = rst1("logo")
			logoh = rst1("logoh")
			logow = rst1("logow")		
			portfolio=rst1("portfolio")		
		end if
		rst1.close

    sql = " SELECT b.id as billid, b.leaseutilityid, ypid, b.utility, extusg " & _
		  " FROM tblbillbyperiod b " & _
		  " WHERE reject=0 and bldgnum='"&cbldg&"' and billyear="&byear&" and billperiod="&bperiod
		  
    if isnumeric(utilityid) then sql = sql & " and utility="&utilityid
    sql = sql & "  ORDER BY TenantName"
		
	timeToGenerate = 0
	totalPageCount = 0
	dim blgdrset
	Set blgdrset = Server.CreateObject("ADODB.recordset")
	blgdrset.open sql, cnn1
	do until blgdrset.eof
		dim templid, tempypid, temputility
		templid = trim(blgdrset("leaseutilityid"))
		tempypid = trim(blgdrset("ypid"))
		temputility = trim(blgdrset("utility"))
		
		Dim metercountInfo
		Set metercountInfo = Server.CreateObject("ADODB.recordset")
		billid = blgdrset("billid")
		metercountInfo.open "select count(*) as metercount " & _
							"from tblmetersbyperiod tm,buildings b,meters m " & _
							"where tm.bldgnum =b.bldgnum and tm.meternum=m.meternum " & _
							"and b.bldgnum = m.bldgnum and bill_id="&billid, cnn1
		
		dim tempMaxPageCount
		tempMaxPageCount = metercountInfo("metercount") \ 40 + 1

		
		if metercountInfo("metercount") > 5 then
			tempMaxPageCount = tempMaxPageCount + 1
		end if
		' 2 seconds per page to generate
		timeToGenerate = timeToGenerate + ( tempMaxPageCount * 2 )
		totalPageCount = totalPageCount + tempMaxPageCount
        If tempMaxPageCount > MaxPageCount Then
            MaxPageCount = tempMaxPageCount
        End If		
        
		metercountInfo.Close()
		blgdrset.movenext
	loop
	blgdrset.Close()
	timeToGenerate = timetogenerate + 20
	if timeToGenerate > 60 then
		timeToGenerate =	timeToGenerate \ 60 
		timeunit = "min"
	end if
	if timeToGenerate > 1 then
		timeunit=timeunit & "s"
	end if
end if

 %>
 <HTML>
 <HEAD>
<script type="text/javascript">
	function formSubmit()
	{
		document.getElementById("pdfLinks").submit();
	}
</script> 
 </HEAD>
 <BODY>
 <Form name=pdflinks action="pdfLinks.asp">
<table width='100%'  >
<tr>
<td bgcolor='#336699' align='center'>
<font face='Arial, Helvetica, sans-serif' color=white>gEnergyOne<br>PDF Management Server</font></td><td>&nbsp;</td>
</tr>


<%
Function CheckRemoteURL(fileURL)
    ON ERROR RESUME NEXT
    Dim xmlhttp

    Set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")

    xmlhttp.open "GET", fileURL, False
    xmlhttp.send
    If(Err.Number<>0) then
        Response.Write "Could not connect to remote server"
    else
        Select Case Cint(xmlhttp.status)
            Case 200, 202, 302
                Set xmlhttp = Nothing
                CheckRemoteURL = True
            Case Else
                Set xmlhttp = Nothing
                CheckRemoteURL = False
        End Select
    end if
    ON ERROR GOTO 0
End Function



	Dim blnBillsAvailable 
	Dim fso, strFileName, i
	Set fso = CreateObject("Scripting.FileSystemObject")
	dim ctime,absfile
	ctime = year(now)&month(now)&day(now)&hour(now)&minute(now)&second(now)
	Dim IspostBack, sbldg
	IspostBack=false
	IspostBack = Request("blnPostBack")
	sbldg = Replace(building, " ", "")
	blnBillsAvailable = false
	dim pdfdir, pdfname, newpdfname, root, have
	root = "D:\WebSites\isabella\genergyonline.com\pdfmaker\"
	pdfdir =  portfolio & "\" & ucase(cbldg) & "\"
	if detailed then dname = "_Detailed" end if
	newpdfname = ucase(cbldg) &"_"& byear & "." & Right("0" & bperiod, 2) &"_"& utilname & dname &"_TenantBills.pdf"
	for i = 1 to 1
			strFileName	= "D:\WebSites\isabella\genergyonline.com\pdfmaker\" & portfolio & "\" & ucase(cbldg) & "\" & ucase(cbldg) & byear & bperiod & utilityid & i & ".pdf"
			PDFName = ucase(cbldg) & byear & bperiod & utilityid & i & ".pdf"
			if detailed then PDFName = "D_"&PDFName
			absfile = portfolio & "\" & ucase(cbldg) & "\" & ucase(cbldg) & byear & bperiod & utilityid & i & ".pdf"

			if fso.fileexists(root&pdfdir&PDFName) then
				fso.copyFile root&pdfdir&PDFName, root&pdfdir&newpdfname, true
				fso.deletefile(root&pdfdir&pdfname)
			end if	
			if CheckRemoteURL("http://pdfmaker.genergyonline.com/pdfMaker/"&pdfdir&newpdfname) and not have and not postback then
				have = true
				blnBillsAvailable = True
			%>
			<tr><td>&nbsp;</td></tr>
			<tr><td Colspan=4><b>PDF Available:</b>
			<a style="font-family:arial;font-size:12;text-decoration:none;" href="http://pdfmaker.genergyonline.com/pdfMaker/<%=pdfdir%><%=newpdfname%>?dt=<%=ctime%>" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'blue'"><b><%=newpdfname%></b></a> 
			</td></tr>
			<%
			Elseif not postback then %>
				<tr><td Colspan=2>
					<%Response.Write "There are no pdf bill files available."%>
				</td></tr>
		<%	End If
	next
	set fso=nothing
	If blnBillsAvailable = false and not postback  then %>
		<tr><td Colspan=2>
			No pdf files have been generated as yet. 
		</td></tr>
		<tr>
	<%
	End IF		

	If IspostBack = "true" then
		If PdfRequest <> "true" then
			Dim cmd, prm
			set cmd = server.createobject("ADODB.Command") 
			
			cmd.ActiveConnection = cnn1
			cmd.CommandType = adCmdStoredProc
			if detailed then
				cmd.CommandText = "usp_G1ConsolePdfs_PADetail"
			else
				cmd.CommandText = "usp_TestG1ConsolePdfs"
			end if

			Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
			cmd.Parameters.Append prm
			Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
			cmd.Parameters.Append prm
			Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
			cmd.Parameters.Append prm
			Set prm = cmd.CreateParameter("utility", adInteger, adParamInput)
			cmd.Parameters.Append prm


				cmd.parameters("bldg") = Replace(building,"+"," ")
				cmd.parameters("by") = byear
				cmd.parameters("bp") = bperiod
				cmd.parameters("utility") = utilityid

'Response.Write("bldgId: " + building + ", billyear: " + byear + ", bp: " + bperiod + ", utilityId: " + utilityid)
					
			cmd.execute
			
				PdfRequest = true
		End If
						
			%>
		<td Colspan=4>Your request for PDF generation for <b><%=cbldg%></b> is being processed. 
		</td></tr>
		<tr><td>
		Wait time approximately <b><%=timeToGenerate%> <%=timeunit %></b> 
		</td>
		<td>
		<tr><td>&nbsp;</td></tr>
		</tr><tr><td>CLOSING..</td>
		<%
		function sleep(scs)
			Dim lo_wsh, ls_cmd
			Set lo_wsh = CreateObject( "WScript.Shell" )
			ls_cmd = "%COMSPEC% /c ping -n " & 1 + scs & " 127.0.0.1>nul"
			lo_wsh.Run ls_cmd, 0, True 
		End Function
		sleep(8)
		Response.Write ("<script>self.close();</script>")
		Response.End
		%>
		
	<%else%>
		<tr><td>&nbsp;</td></tr>
		<td Colspan=4>Please click here to (Re)Generate PDFs : <INPUT type=submit value="(Re)Generate PDFs"></INPUT></td>
		<td Colspan=4>&nbsp;</td>
	<%end if%>
	</tr>
	<% sbldg = Replace(building, " ", "+")%>
	<INPUT type=hidden name=blnPostBack value=true></INPUT>
	<INPUT type=hidden name=requestPosted value=<%=PdfRequest%>></INPUT>
	<INPUT type=hidden name=building value=<%=sbldg%>></INPUT>
	<INPUT type=hidden name=byear value=<%=byear%>></INPUT>
	<INPUT type=hidden name=bperiod value=<%=bperiod%>></INPUT>
	<INPUT type=hidden name=utilityid value=<%=utilityid%>></INPUT>
	<INPUT type=hidden name=detailed value=<%=detailed%>></INPUT>
	<INPUT type=hidden name=pid value=<%=pid%>></INPUT>
</form>
</body>		
<%
set rst1 = nothing
set cnn1 = nothing
response.End()
%>