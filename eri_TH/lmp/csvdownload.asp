<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Cache-Control", "must-revalidate"

if request("file") <> "" then 
Response.AddHeader "content-disposition", "attachment;filename=G1_LoadProfile.csv"
Dim RS, SQL, Cnn,bldg, meterid, billingid, sd,ed, luid, utility, interval,dsarray


bldg = request.querystring("bldg")
meterid = request.querystring("meterid")
billingid = request.querystring("billingid")
sd = request.querystring("sd")
ed = request.querystring("ed")
utility = request("utility")
interval = request("interval")
if isdate(ed) then ed = dateadd("n",-1,dateadd("d",1,ed))

set cnn = server.createobject("ADODB.Connection")
set rs = server.createobject("ADODB.Recordset")
cnn.Open getLocalConnect(bldg)

dim lmpid, lmptype
if trim(meterid)<>"" then
    lmptype="m"
	sql = "select meterid,datasource from meters where meterid = " & meterid
elseif trim(billingid)<>"" then
    lmptype="L"
	sql = "select meterid, datasource from meters m,tblLeases l , tblleasesutilityprices lup WHERE l.billingid=lup.billingid AND lup.utility="&utility&" and l.billingId="&billingid&" and online = 1 and m.leaseutilityid = lup.leaseutilityid order by meterid"
elseif trim(bldg)<>"" then
    lmptype="b"
	sql = "select meterid, datasource from meters where lmp = 1 and bldgnum = '" & bldg & "'"
end if

    rs.open sql, cnn
    if not(rs.eof) then 
		while not rs.eof 
		if lmpid = "" then 
			lmpid = rs("meterid")
		else
			lmpid = lmpid & "," & rs("meterid")
		end if
		dsarray = rs("datasource")
			rs.movenext
		wend
	end if
    rs.close


SQL = "select bldgnum as [BUILDING NUMBER], meterid as [METERID], date as [TIME STAMP], kwh as [DATA], case est when 1 then 'E' else 'A' end as [DATA STATUS] from [10.0.7.149].genergy2.dbo."&dsarray&" where meterid in  ("&lmpid&") and date between '"&sd&"' and '"&ed&"' order by meterid, date"
rs.open sql, cnn

Dim F, Head
For Each F In RS.Fields
  Head = Head & ", " & F.Name
Next
Head = Mid(Head,3) & vbCrLf
Response.ContentType = "text/plain"
Response.Write Head
Response.Write RS.GetString(,,", ",vbCrLf,"")
else
%>
<a href="testdownload.asp?file=show&&bldg=71&billingid=686&utility=2&sd=6/1/2005&ed=6/8/2005">Click to download</a><%end if%>