<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Cache-Control", "must-revalidate"

if request("process") = 1 then 	
	On Error Resume Next
	
	Response.AddHeader "content-disposition", "attachment;filename=G1_LoadProfile.csv"
	Dim RS, SQL, Cnn,bldg, meterid, billingid, sd,ed, luid, utility, interval,dsRS,outputRS
	
	bldg = request.querystring("bldg")
	meterid = request.querystring("meterid")
	billingid = request.querystring("billingid")
	sd = request.querystring("sd")
	ed = request.querystring("ed")
	utility = request("utility")
	interval = request("interval")
   		
	if trim(interval) = "" then interval=0
	
	if isdate(ed) then ed = dateadd("n",-1,dateadd("d",1,ed))
	
	set cnn = server.createobject("ADODB.Connection")
	set rs = server.createobject("ADODB.Recordset")
	set outputRS = server.createobject("ADODB.Recordset")
	set dsRS = server.createobject("ADODB.Recordset")
	cnn.Open getLocalConnect(bldg)
	
	'Create Temporary Recordset for CSV File Output
	outputRS.Fields.Append "DESCRIPTION", adVarChar, 50
	outputRS.Fields.Append "TIMESTAMP", adVarChar, 50
	outputRS.Fields.Append "DATA", adChar, 50
	outputRS.Fields.Append "DATAMEASURE", adChar, 50
	
	dim lmpid, lmptype, qryStr1
	
	'get meter list
	if trim(meterid)<>"" then
		lmptype="m"
		sql = "select distinct meternum as description,dmeasure,umeasure,u.utilityid from meters  m inner join tblleasesutilityprices lup on lup.leaseutilityid = m.leaseutilityid inner join tblutility u on u.utilityid = lup.utility where u.utilityid = "&utility&" and meterid = " & meterid
		qryStr1 = meterid 
	elseif trim(billingid)<>"" then
		lmptype="L"
		sql = "select distinct l.billingname as description, dmeasure,umeasure,u.utilityid  from meters m,tblLeases l , tblleasesutilityprices lup,tblutility u WHERE l.billingid=lup.billingid AND lup.utility="&utility&" and l.billingId= "&billingid&" and lup.utility = u.utilityid and online = 1 and m.leaseutilityid = lup.leaseutilityid"
		qryStr1 = billingid
	elseif trim(bldg)<>"" then
		lmptype="b"
		sql = "select distinct strt as description, dmeasure,umeasure,u.utilityid  from meters m inner join tblleasesutilityprices lup on lup.leaseutilityid = m.leaseutilityid inner join tblutility u on u.utilityid = lup.utility inner join  buildings b on b.bldgnum = m.bldgnum where lmp = 1 and u.utilityid = "&utility&" and b.bldgnum = '"&bldg&"' order by description"
		qryStr1 = bldg
	end if
				
	if sql <> "" then 
		rs.open sql, cnn
		if not(rs.eof) then 
			outputRS.Open
				while not rs.eof	
				 sql = "EXEC [sp_LMPDATA] '"&sd&"', '"&ed&"', '"&lmptype&"', '"&qrystr1&"', "&rs("utilityid")&","&interval&", 0, 0, 0, 0"
				dsrs.open sql, cnn
				if not(dsrs.eof) then 
					while not dsrs.eof 
						outputRS.AddNew
						outputRS("DESCRIPTION") = rs("description")
						if len(trim(dsrs("date"))) < 11 then 
							outputRS("TIMESTAMP") = trim(dsrs("date")) & " 12:00:00 AM"
						else 
							outputRS("TIMESTAMP") = dsrs("date")
						end if
						
						select case utility
						
						case 2,1
							outputRS("DATA") = trim(dsrs(trim(rs("UMEASURE"))))
						case 6 
							outputRS("DATA") = cdbl(dsrs(trim(rs("DMEASURE"))))/4			
						case else
							outputRS("DATA") = trim(dsrs(trim(rs("DMEASURE"))))				
						end select 
						
						
						outputRS("DATAMEASURE") = trim(rs("DMEASURE"))
					dsrs.movenext
					wend
				end if
				dsrs.close
				rs.movenext
				wend
		end if
		outputRS.Updatebatch
		outputRS.Movefirst
		Dim F, Head
		For Each F In outputRS.Fields
		  Head = Head & ", " & F.Name
		Next
		Head = Mid(Head,3) & vbCrLf
		Response.ContentType = "text/plain"
		Response.Write Head
		Response.Write outputRS.GetString(,,", ",vbCrLf,"")
	rs.close
	end if
else

	bldg = request.querystring("bldg")
	meterid = request.querystring("meterid")
	billingid = request.querystring("billingid")
	sd = request.querystring("sd")
	ed = request.querystring("ed")
	utility = request("utility")
	interval = request("interval")

%>
<table align="center" height="100%" width="100%"><tr>
  <td align="center" valign="center">Preparing data file for download<br>  <br>
    <a href="accesslmphistory.asp?meterid=<%=meterid%>&bldg=<%=bldg%>&billingid=<%=billingid%>&utility=<%=utility%>&startdate=<%=sd%>">Once completed click here to return to the menu</a></td>
</tr></table>
<script>
  document.location="<%response.write "./csvdownload.asp?process=1&bldg="&request.querystring("bldg")&"&meterid="&request.querystring("meterid")&"&billingid="&request.querystring("billingid")&"&sd="&request.querystring("sd")&"&ed="&request.querystring("ed")&"&utility="&request("utility")&"&interval="&request("interval")%>"
</script>
<%
end if
%>