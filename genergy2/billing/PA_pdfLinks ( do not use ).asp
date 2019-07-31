<%
Option Explicit
Response.Buffer = False%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim leaseid, ypid, building, pid, byear, bperiod, logo
dim genergy2, devIP, utilityid, detailed, meterbreakdown
Dim SJPproperties, summaryusage, summarydemand, buildpdf
Dim textheader,msg,appTimeout,demo, billid,billurl
dim billCount, logow, logoh,rst11

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

dim maxPageCount
Dim timeToGenerate
Dim PdfRequest
Dim totalPageCount
dim cnn1, rst1, rst2, sql
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getLocalConnect(building)


Dim portfolio
if building<>"" then

	rst1.open "SELECT location, b.bldgnum, b.portfolioid,billurl,logo, logoh, logow, p.portfolio FROM buildings b, portfolio p, billtemplates bt WHERE b.portfolioid=p.id AND bt.id=p.templateid AND bldgnum='"&building&"'	", cnn1
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
		  " WHERE reject=0 and bldgnum='"&building&"' and billyear="&byear&" and billperiod="&bperiod
		  
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
		' 20 seconds per page to generate
		timeToGenerate = timeToGenerate + ( tempMaxPageCount * 20 )
		totalPageCount = totalPageCount + tempMaxPageCount
        If tempMaxPageCount > MaxPageCount Then
            MaxPageCount = tempMaxPageCount
        End If		
        
		metercountInfo.Close()
		blgdrset.movenext
	loop
	blgdrset.Close()
	
	if timeToGenerate > 0 then
		timeToGenerate =	timeToGenerate \ 60 
	end if
end if

 %>
 <HTML>
 <HEAD>
<script type="text/javascript">
	function formSubmit()
	{
		document.getElementById("PA_pdfLinks").submit();
	}
</script> 
 </HEAD>
 <BODY>
 <Form name="PA_pdflinks" action="PA_pdfLinks.asp">
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
	Dim objFSO, strFileName, i
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	dim ctime,absfile
	Dim IspostBack
	IspostBack=false
	IspostBack = Request("blnPostBack")
	
	blnBillsAvailable = false
	for i = 1 to maxPageCount
			strFileName	= "C:\ghnet_websites\genergyonline.com\pdfMaker\" & portfolio & "\" & building & "\d_" & building & byear & bperiod & utilityid & i & ".pdf"	
			absfile = portfolio & "\" & building & "\d_" & building & byear & bperiod & utilityid & i & ".pdf"			
			'If objFSO.FileExists(strFileName) Then 
			if CheckRemoteURL("http://pdfmaker.genergyonline.com/pdfMaker/"&absfile) then
				blnBillsAvailable = True
			%>
			<tr><td Colspan=2>
			<a style="font-family:arial;font-size:12;text-decoration:none;" href="http://pdfmaker.genergyonline.com/pdfMaker/<%=portfolio%>/<%=building%>/<%="d_"%><%=building%><%=byear%><%=bperiod%><%=utilityid%><%=i%>.pdf?dt=<%=ctime%>" target="_blank" onMouseOver="this.style.color= 'lightblue'"; onMouseOut="this.style.color = 'blue'"><b><%=i%> page Bills pdf: <%="d_"%> <%=building%><%=byear%><%=bperiod%><%=utilityid%><%=i%>.pdf</b></a> 
			</td></tr>
			<%
			Else %>
				<tr><td Colspan=2>
					<%Response.Write "There are no " & i & " page bills available."%>
				</td></tr>
		<%	End If
	next
	If blnBillsAvailable = false  then %>
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
					cmd.CommandText = "usp_G1ConsolePdfs_PADetail"

					Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
					cmd.Parameters.Append prm
					Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
					cmd.Parameters.Append prm
					Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
					cmd.Parameters.Append prm
					Set prm = cmd.CreateParameter("utility", adInteger, adParamInput)
					cmd.Parameters.Append prm


						cmd.parameters("bldg") = building
						cmd.parameters("by") = byear
						cmd.parameters("bp") = bperiod
						cmd.parameters("utility") = utilityid
							
					cmd.execute
					
						PdfRequest = true
				End If
						
				%>
			<td Colspan=2>Your request for PDF generation for <%=Building%>  is being processed. There are <%=totalPageCount%> 
						PDF pages to be generated and this process will take approx. <%=timeToGenerate%> Minutes. 
			Please Re-open this page after  <%=timeToGenerate%> Minutes. In case of any problems, please contact Genergy.</td>
			<%else%>
			<td Colspan=2>Please click here to (Re)Generate PDFs : <INPUT type=submit value="(Re)Generate PDFs"></INPUT></td>
		<%end if%>
	</tr>
	<INPUT type=hidden name=blnPostBack value=true></INPUT>
	<INPUT type=hidden name=requestPosted value=<%=PdfRequest%>></INPUT>
	<INPUT type=hidden name=building value=<%=building%>></INPUT>
	<INPUT type=hidden name=byear value=<%=byear%>></INPUT>
	<INPUT type=hidden name=bperiod value=<%=bperiod%>></INPUT>
	<INPUT type=hidden name=utilityid value=<%=utilityid%>></INPUT>
	<INPUT type=hidden name=pid value=<%=pid%>></INPUT>
</form>
</body>		
<%
set rst1 = nothing
set cnn1 = nothing
response.End()
%>