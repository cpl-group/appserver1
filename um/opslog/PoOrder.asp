<%option explicit%>
<!--#INCLUDE Virtual="/genergy2/secure.inc"-->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
if not allowgroups("Department Supervisors,IT Services,gAccounting")  then  
	Response.Redirect "acctpoview.asp"
end if

Dim cnn1, rst1,rst2, sqlstr,rst3,poNumz,compType,sql666,result,vendorSelect,cmd,cmd2,cmd3,sql777,rsscomp
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
Set rst3 = Server.CreateObject("ADODB.recordset")
Set rsscomp = Server.CreateObject("ADODB.recordset")

set cmd = server.createobject("adodb.command")
set cmd2 = server.createobject("adodb.command")
set cmd3 = server.createobject("adodb.command")
cnn1.Open getConnect(0,0,"intranet")

sql666="delete from po_temp"
rst2.Open sql666, cnn1
sql666="insert into PO_TEMP ([Commitment], [PONum], [Vendor], [JobNum], [JobName], [JobAddr], [ShipAddr], [PODate], [Requistioner], [submittedby], [PO_Total], [submitted], [accepted], [description], [ship_amt], [admin_comment], [closed], [closed_user], [tax], [question], [accepted_user], [acct_ponum], [approved_comment], [approved], [approved_user], [vid]) select  convert(varchar,jobnum)+'.'+convert(varchar,ponum)as [Commitment], [PONum], [Vendor], [JobNum], [JobName], [JobAddr], [ShipAddr], replace(convert(varchar, [podate],101),'/',''), [Requistioner], [submittedby], [PO_Total], [submitted], [accepted], [description], [ship_amt], [admin_comment], [closed], [closed_user], [tax], [question], [accepted_user], [acct_ponum], [approved_comment], [approved], [approved_user], [vid]  from  po  where accepted=1 and closed=0 and vid <> '0'"
rst2.Open sql666, cnn1
'response.write sql666
'response.end
'rst3.close


rst1.open "select * from companycodes where active = 1 and code <> 'AC' order by name", getConnect(0,0,"intranet")
do until rst1.eof
	vendorSelect = vendorSelect & "SELECT [name], vendor, '"&rst1("code")&"' as comp FROM ["&rst1("code")&"_MASTER_APM_VENDOR] UNION all "
	rst1.movenext
loop
rst1.close
vendorSelect = "(SELECT distinct * FROM (" & vendorSelect
vendorSelect = left(vendorSelect,len(vendorSelect)-10) & ") v)"
sqlstr = "select ltrim(str(po.Jobnum)) + '.' + ltrim(str(po.POnum)) as ponumber,po.*, employees.[first name]+' '+employees.[last name] as req, case when po.vid<>'0' then vs.name else po.vendor end as vendorname from po INNER JOIN master_job m ON po.jobnum=m.id join employees on po.requistioner=substring(employees.username,7,20) LEFT JOIN "&vendorSelect&" vs ON vs.vendor=vid and vs.comp=m.company WHERE accepted = 1 and closed=0 order by podate desc"

rst1.Open sqlstr, cnn1, 0, 1, 1

if not rst1.EOF then 

While not rst1.EOF

		poNumz=split(rst1("ponumber"),".")(0)
		sql666="select company from master_job where id = '"&poNumz&"'"
				
		rst3.Open sql666, cnn1
		if not rst3.EOF then 
		compType=rst3("company")
		end if
		rst3.close  

dim GYresult,GEresult,NYresult,GYflag,GEflag,NYflag,resultlist,GYPoFList,GEPoFList,NYPoFList
Select Case compType
        
		Case "GY"
            
 GYresult=GYresult+ "SET ANSI_WARNINGS OFF insert into PoMaster_FILE([C],[Commitment], [PO Type],[description],[Vendor], [podate],[noos],[qwerty],[Closed],[Name], [Address_1], [Address_2], [City], [State], [ZIP],[a] ,[b1] ,[b2] , [b3] ,[b4] ,[b5] ,[b6] ,[b7] ,[b8] ,[b9] ,[b10] ,[bq] ,[bw] ,[be] ,[br] ,[bt] ,[by] ,[bu] ,[bi] ,[bo] ,[bp] ,[ba] ,[bs] ,[bd] ,[bf] ,[bg] ,[bh] ,[bj] ,[bk] ,[bl] ,[bz] ,[bx] ,[bc] ,[bv] ,[bb] ,[qb] ,[wb] ,[eb] ,[rb] ,[tb] ,[yb] ,[ybqqqq] ,[ub] ,[ib] ,[ob] ,[b] ,[pb] ,[ab] ,[sb] ,[db] ,[gb] ,[fb] ,[hb] ,[jb] ,[kaab] ,[kb] ,[lb] ,[zb] ,[xb] ,[cb] ,[vb] ,[nb] ,[mb] ,[qqb] ,[bqq] ,[wwb] ,[eeb] ,[rrb] ,[ttb] ,[yyb] ,[uub] ,[iib] ,[oob] ,[ppb] ,[aab] ,[ssb] ,[ddb] ,[ffb] ,[ggb] ,[hhb] ,[jjb] ,[kkb] ,[llb] ,[zzb] ,[xxb] ,[ccb] ,[vvb] ,[bbbaaa] ,[nnb] ,[mmb] ,[bqqeeeee] ,[bww] ,[bee] ,[brr] ,[btt] ,[byy] ,[buu] ,[bii] ,[boo] ,[bpp] ,[baa] ,[bss] ,[bdd] ,[bff] ,[bgg] ,[bhh] ,[bjj] ,[bkk] ,[bll] ,[bzz] ,[bxx] ,[bcc] ,[bvv] ,[bbb] ,[bnn] ,[bmm] ,[bmmm] ,[bmmmm] ,[bmmmmm] ,[bmmmmmmm] ,[bmmmmmmmmmm] ,[bmmmmmmmmmmmmm] ,[bmmmmmmmmmmmmmmm] ,[bmmmmmmmmmmmmmmmmm] ,[bqqqqqqqqqq] ,[bqqqqqqqqqqqqq] ,[bwwwwwwwwwwwww] ,[beeeeeeeeeee] ,[beeeeeee] ) select distinct isnull('C','""c""'),'""'+po.commitment+'""','""2""','""'+left(jobaddr,30)+'""', '""'+v.Vendor+'""','""'+isnull(convert(varchar,podate)+'""',getdate()),'""""','""Y""','""N""','""""', '""'+Address_1+'""', '""'+Address_2+'""','""'+left(City,15)+'""', '""'+State+'""', '""'+ZIP+'""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""' from GY_MASTER_APM_VENDOR v inner join PO_TEMP po on po.vid = v.vendor where accepted=1 and closed=0  and po.commitment='"&rst1("ponumber")&"';"  
		if GYflag <> "1" then
		resultlist = resultlist + "GYresult,"
        end if
	    GYflag="1"
	 	GYPoFList=GYPoFList+rst1("ponumber")+","
	
	   Case "GE"
         
		   
 GEresult=GEresult+ "SET ANSI_WARNINGS OFF insert into PoMaster_FILE([C],[Commitment], [PO Type],[description],[Vendor], [podate],[noos],[qwerty],[Closed],[Name], [Address_1], [Address_2], [City], [State], [ZIP],[a] ,[b1] ,[b2] , [b3] ,[b4] ,[b5] ,[b6] ,[b7] ,[b8] ,[b9] ,[b10] ,[bq] ,[bw] ,[be] ,[br] ,[bt] ,[by] ,[bu] ,[bi] ,[bo] ,[bp] ,[ba] ,[bs] ,[bd] ,[bf] ,[bg] ,[bh] ,[bj] ,[bk] ,[bl] ,[bz] ,[bx] ,[bc] ,[bv] ,[bb] ,[qb] ,[wb] ,[eb] ,[rb] ,[tb] ,[yb] ,[ybqqqq] ,[ub] ,[ib] ,[ob] ,[b] ,[pb] ,[ab] ,[sb] ,[db] ,[gb] ,[fb] ,[hb] ,[jb] ,[kaab] ,[kb] ,[lb] ,[zb] ,[xb] ,[cb] ,[vb] ,[nb] ,[mb] ,[qqb] ,[bqq] ,[wwb] ,[eeb] ,[rrb] ,[ttb] ,[yyb] ,[uub] ,[iib] ,[oob] ,[ppb] ,[aab] ,[ssb] ,[ddb] ,[ffb] ,[ggb] ,[hhb] ,[jjb] ,[kkb] ,[llb] ,[zzb] ,[xxb] ,[ccb] ,[vvb] ,[bbbaaa] ,[nnb] ,[mmb] ,[bqqeeeee] ,[bww] ,[bee] ,[brr] ,[btt] ,[byy] ,[buu] ,[bii] ,[boo] ,[bpp] ,[baa] ,[bss] ,[bdd] ,[bff] ,[bgg] ,[bhh] ,[bjj] ,[bkk] ,[bll] ,[bzz] ,[bxx] ,[bcc] ,[bvv] ,[bbb] ,[bnn] ,[bmm] ,[bmmm] ,[bmmmm] ,[bmmmmm] ,[bmmmmmmm] ,[bmmmmmmmmmm] ,[bmmmmmmmmmmmmm] ,[bmmmmmmmmmmmmmmm] ,[bmmmmmmmmmmmmmmmmm] ,[bqqqqqqqqqq] ,[bqqqqqqqqqqqqq] ,[bwwwwwwwwwwwww] ,[beeeeeeeeeee] ,[beeeeeee] ) select distinct isnull('C','""c""'),'""'+po.commitment+'""','""2""','""'+left(jobaddr,30)+'""', '""'+v.Vendor+'""','""'+isnull(convert(varchar,podate)+'""',getdate()),'""""','""Y""','""N""','""""', '""'+Address_1+'""', '""'+Address_2+'""','""'+left(City,15)+'""', '""'+State+'""', '""'+ZIP+'""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""' from GE_MASTER_APM_VENDOR v inner join PO_TEMP po on po.vid = v.vendor where accepted=1 and closed=0 and po.commitment='"&rst1("ponumber")&"';"  
		if GEflag <> "1" then
		resultlist = resultlist + "GEresult,"
        end if
		GEflag="1"
		GEPoFList=GEPoFList+rst1("ponumber")+","
		
		Case "NY"
		  
 NYresult=NYresult+ "SET ANSI_WARNINGS OFF insert into PoMaster_FILE([C],[Commitment], [PO Type],[description],[Vendor], [podate],[noos],[qwerty],[Closed],[Name], [Address_1], [Address_2], [City], [State], [ZIP],[a] ,[b1] ,[b2] , [b3] ,[b4] ,[b5] ,[b6] ,[b7] ,[b8] ,[b9] ,[b10] ,[bq] ,[bw] ,[be] ,[br] ,[bt] ,[by] ,[bu] ,[bi] ,[bo] ,[bp] ,[ba] ,[bs] ,[bd] ,[bf] ,[bg] ,[bh] ,[bj] ,[bk] ,[bl] ,[bz] ,[bx] ,[bc] ,[bv] ,[bb] ,[qb] ,[wb] ,[eb] ,[rb] ,[tb] ,[yb] ,[ybqqqq] ,[ub] ,[ib] ,[ob] ,[b] ,[pb] ,[ab] ,[sb] ,[db] ,[gb] ,[fb] ,[hb] ,[jb] ,[kaab] ,[kb] ,[lb] ,[zb] ,[xb] ,[cb] ,[vb] ,[nb] ,[mb] ,[qqb] ,[bqq] ,[wwb] ,[eeb] ,[rrb] ,[ttb] ,[yyb] ,[uub] ,[iib] ,[oob] ,[ppb] ,[aab] ,[ssb] ,[ddb] ,[ffb] ,[ggb] ,[hhb] ,[jjb] ,[kkb] ,[llb] ,[zzb] ,[xxb] ,[ccb] ,[vvb] ,[bbbaaa] ,[nnb] ,[mmb] ,[bqqeeeee] ,[bww] ,[bee] ,[brr] ,[btt] ,[byy] ,[buu] ,[bii] ,[boo] ,[bpp] ,[baa] ,[bss] ,[bdd] ,[bff] ,[bgg] ,[bhh] ,[bjj] ,[bkk] ,[bll] ,[bzz] ,[bxx] ,[bcc] ,[bvv] ,[bbb] ,[bnn] ,[bmm] ,[bmmm] ,[bmmmm] ,[bmmmmm] ,[bmmmmmmm] ,[bmmmmmmmmmm] ,[bmmmmmmmmmmmmm] ,[bmmmmmmmmmmmmmmm] ,[bmmmmmmmmmmmmmmmmm] ,[bqqqqqqqqqq] ,[bqqqqqqqqqqqqq] ,[bwwwwwwwwwwwww] ,[beeeeeeeeeee] ,[beeeeeee] ) select distinct isnull('C','""c""'),'""'+po.commitment+'""','""2""','""'+left(jobaddr,30)+'""', '""'+v.Vendor+'""','""'+isnull(convert(varchar,podate)+'""',getdate()),'""""','""Y""','""N""','""""', '""'+Address_1+'""', '""'+Address_2+'""','""'+left(City,15)+'""', '""'+State+'""', '""'+ZIP+'""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""' from NY_MASTER_APM_VENDOR v inner join PO_TEMP po on po.vid = v.vendor where accepted=1 and closed=0 and po.commitment='"&rst1("ponumber")&"';"  
		if NYflag <> "1" then
		resultlist = resultlist + "NYresult,"
		end if
		NYflag="1"
		NYPoFList=NYPoFList+rst1("ponumber")+","
		End Select

 'result=result+ "SET ANSI_WARNINGS OFF insert into PoMaster_FILE([C],[Commitment], [PO Type],[description],[Vendor], [podate],[noos],[qwerty],[Closed],[Name], [Address_1], [Address_2], [City], [State], [ZIP],[a] ,[b1] ,[b2] , [b3] ,[b4] ,[b5] ,[b6] ,[b7] ,[b8] ,[b9] ,[b10] ,[bq] ,[bw] ,[be] ,[br] ,[bt] ,[by] ,[bu] ,[bi] ,[bo] ,[bp] ,[ba] ,[bs] ,[bd] ,[bf] ,[bg] ,[bh] ,[bj] ,[bk] ,[bl] ,[bz] ,[bx] ,[bc] ,[bv] ,[bb] ,[qb] ,[wb] ,[eb] ,[rb] ,[tb] ,[yb] ,[ybqqqq] ,[ub] ,[ib] ,[ob] ,[b] ,[pb] ,[ab] ,[sb] ,[db] ,[gb] ,[fb] ,[hb] ,[jb] ,[kaab] ,[kb] ,[lb] ,[zb] ,[xb] ,[cb] ,[vb] ,[nb] ,[mb] ,[qqb] ,[bqq] ,[wwb] ,[eeb] ,[rrb] ,[ttb] ,[yyb] ,[uub] ,[iib] ,[oob] ,[ppb] ,[aab] ,[ssb] ,[ddb] ,[ffb] ,[ggb] ,[hhb] ,[jjb] ,[kkb] ,[llb] ,[zzb] ,[xxb] ,[ccb] ,[vvb] ,[bbbaaa] ,[nnb] ,[mmb] ,[bqqeeeee] ,[bww] ,[bee] ,[brr] ,[btt] ,[byy] ,[buu] ,[bii] ,[boo] ,[bpp] ,[baa] ,[bss] ,[bdd] ,[bff] ,[bgg] ,[bhh] ,[bjj] ,[bkk] ,[bll] ,[bzz] ,[bxx] ,[bcc] ,[bvv] ,[bbb] ,[bnn] ,[bmm] ,[bmmm] ,[bmmmm] ,[bmmmmm] ,[bmmmmmmm] ,[bmmmmmmmmmm] ,[bmmmmmmmmmmmmm] ,[bmmmmmmmmmmmmmmm] ,[bmmmmmmmmmmmmmmmmm] ,[bqqqqqqqqqq] ,[bqqqqqqqqqqqqq] ,[bwwwwwwwwwwwww] ,[beeeeeeeeeee] ,[beeeeeee] ) select distinct isnull('C','""c""'),'""'+po.commitment+'""','""2""','""'+description+'""', '""'+v.Vendor+'""','""'+isnull(convert(varchar,podate)+'""',getdate()),'""""','""Y""','""Y""','""""', '""'+Address_1+'""', '""'+Address_2+'""','""'+ City+'""', '""'+State+'""', '""'+ZIP+'""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""','""""' from "&compType&"_MASTER_APM_VENDOR v inner join PO_TEMP po on po.vid = v.vendor where accepted=1 and closed=0;"  

	rst1.movenext
  	Wend
	'response.write GEresult
	'response.end
'GYPoFList=left(GYPoFList,len(GYPoFList)-1)
'response.write  GYPoFList
'response.end 
'response.write GEresult
'response.end
dim companyresult,list,prm,GYlist,GEList,NYList,GYPonum,GEPonum,NYPonum,GYPonumlist,GEPonumlist,NYPonumlist,setzero
'sql666 = "SET ANSI_WARNINGS OFF insert into po_file([Record ID], [Commitment ID], [Item number], [Description], [Retain percent], [Delivery  Date], [Scope of Work], [Job], [Extra], [Cost Code], [Category], [tax grp], [Tax], [units], [unit cost], [unit Description], [amount],[bought out],varies) select distinct '""CI""' as CI,'""'+convert(varchar,commitment)+'""' as commitment,'""""','""'+[description]+'""','""""','""""','""""','""'+convert(varchar,jobnum)+'""','""""','""'+'110006'+'""','""'+'005'+'""','""""','""""','""""','""""','""""','""'+convert(varchar,po_total)+'""','""""','""""' from po_temp where accepted=1 and closed=0 and commitment='"&GYPonum&"'; "

list=split(resultlist,",")
For Each companyresult in list
'response.write companyresult
'response.end
select case companyresult
'sql666="SET ANSI_WARNINGS OFF"

case "GYresult"

rst3.Open GYresult, cnn1 
'rst3.Open result, cnn1
GYPoFList=left(GYPoFList,len(GYPoFList)-1)
GYPonumlist=split(GYPoFList,",")
sql666=""
'dim setzero
for each GYPonum in GYPonumlist
sql777="Select left(job,2)as prefix ,id from master_job where id = '"&split(GYPonum,".")(0)&"'"
rsscomp.Open sql777, cnn1
if len(rsscomp("id"))=4 then
setzero="00"
else
setzero="0"
end if

'sql666=sql666 + "SET ANSI_WARNINGS OFF insert into po_file([Record ID], [Commitment ID], [Item number], [Description], [Retain percent], [Delivery  Date], [Scope of Work], [Job], [Extra], [Cost Code], [Category], [tax grp], [Tax], [units], [unit cost], [unit Description], [amount],[bought out],varies) select distinct '""CI""' as CI,'""'+convert(varchar,commitment)+'""' as commitment,'""""','""'+[description]+'""','""""','""""','""""','""GY0'+convert(varchar,jobnum)+'""','""""','""'+'110006'+'""','""'+'005'+'""','""""','""""','""""','""""','""""','""'+convert(varchar,po_total)+'""','""""','""""' from po_temp where accepted=1 and closed=0 and commitment='"&GYPonum&"'; "
sql666=sql666 + "SET ANSI_WARNINGS OFF insert into po_file([Record ID], [Commitment ID], [Item number], [Description], [Retain percent], [Delivery  Date], [Scope of Work], [Job], [Extra], [Cost Code], [Category], [tax grp], [Tax], [units], [unit cost], [unit Description], [amount],[bought out],varies) select distinct '""CI""' as CI,'""'+convert(varchar,commitment)+'""' as commitment,'""""','""'+left(description,30)+'""','""""','""""','""""','"""&rsscomp("prefix")&setzero&"'+convert(varchar,jobnum)+'""','""""','""'+'110006'+'""','""'+'005'+'""','""""','""""','""""','""""','""""','""'+convert(varchar,po_total)+'""','""""','""""' from po_temp where accepted=1 and closed=0 and commitment='"&GYPonum&"'; "
sql777=""
rsscomp.close
next
'response.write sql666
'response.end
rst3.Open sql666, cnn1
	
	set cmd.ActiveConnection = cnn1
	cmd.commandText = "RunPOs"
	cmd.CommandType = adCmdStoredProc
	Set prm = cmd.CreateParameter("company", adVarChar, adParamInput, 50)
    cmd.Parameters.Append prm
	cmd.Parameters("company") = "GY"
    cmd.execute
	sql666=""

 case "GEresult"

rst3.Open GEresult, cnn1 
'rst3.Open result, cnn1


GEPoFList=left(GEPoFList,len(GEPoFList)-1)
GEPonumlist=split(GEPoFList,",")
'response.write "1" & GEPoFList &"<BR>"
'response.end
sql666=""
'dim setzero
for each GEPonum in GEPonumlist
sql777="Select left(job,2)as prefix ,id from master_job where id = '"&split(GEPonum,".")(0)&"'"
rsscomp.Open sql777, cnn1
if len(rsscomp("id"))=4 then
setzero="00"
else
setzero="0"
end if


sql666 = sql666 + "SET ANSI_WARNINGS OFF insert into po_file([Record ID], [Commitment ID], [Item number], [Description], [Retain percent], [Delivery  Date], [Scope of Work], [Job], [Extra], [Cost Code], [Category], [tax grp], [Tax], [units], [unit cost], [unit Description], [amount],[bought out],varies) select distinct '""CI""' as CI,'""'+convert(varchar,commitment)+'""' as commitment,'""""','""'+left(description,30)+'""','""""','""""','""""','"""&rsscomp("prefix")&setzero&"'+convert(varchar,jobnum)+'""','""""','""'+'020006'+'""','""'+'005'+'""','""""','""""','""""','""""','""""','""'+convert(varchar,po_total)+'""','""""','""""' from po_temp where accepted=1 and closed=0 and commitment='"&GEPonum&"'; "
'response.write "2)" & GEPonum &"<BR>"

'sql666= "SET ANSI_WARNINGS OFF insert into po_file([Record ID], [Commitment ID], [Item number], [Description], [Retain percent], [Delivery  Date], [Scope of Work], [Job], [Extra], [Cost Code], [Category], [tax grp], [Tax], [units], [unit cost], [unit Description], [amount],[bought out],varies) select distinct '""CI""' as CI,'""'+convert(varchar,commitment)+'""','""""','""'+[description]+'""','""""','""""','""""','""'+convert(varchar,jobnum)+'""','""""','""'+'110006'+'""','""'+'005'+'""','""""','""""','""""','""""','""""','""'+convert(varchar,po_total)+'""','""""','""""' from po_temp where accepted=1 and closed=0"
'response.write sql666
'response.end
rsscomp.close
next
'response.write sql666
'response.end
rst3.Open sql666, cnn1
	set cmd2.ActiveConnection = cnn1
	cmd2.commandText = "RunPOs"
	cmd2.CommandType = adCmdStoredProc
	Set prm = cmd2.CreateParameter("company", adVarChar, adParamInput, 50)
    cmd2.Parameters.Append prm
	cmd2.Parameters("company") = "GE"
    cmd2.execute
   	sql666=""

case "NYresult"

rst3.Open NYresult, cnn1 
'rst3.Open result, cnn1
NYPoFList=left(NYPoFList,len(NYPoFList)-1)
NYPonumlist=split(NYPoFList,",")
sql666=""
'dim setzero
for each NYPonum in NYPonumlist
sql777="Select left(job,2)as prefix ,id from master_job where id = '"&split(NYPonum,".")(0)&"'"
rsscomp.Open sql777, cnn1
if len(rsscomp("id"))=4 then
setzero="00"
else
setzero="0"
end if
sql666 = sql666 + "SET ANSI_WARNINGS OFF insert into po_file([Record ID], [Commitment ID], [Item number], [Description], [Retain percent], [Delivery  Date], [Scope of Work], [Job], [Extra], [Cost Code], [Category], [tax grp], [Tax], [units], [unit cost], [unit Description], [amount],[bought out],varies) select distinct '""CI""' as CI,'""'+convert(varchar,commitment)+'""' as commitment,'""""','""'+left(description,30)+'""','""""','""""','""""','"""&rsscomp("prefix")&setzero&"'+convert(varchar,jobnum)+'""','""""','""'+'010006'+'""','""'+'005'+'""','""""','""""','""""','""""','""""','""'+convert(varchar,po_total)+'""','""""','""""' from po_temp where accepted=1 and closed=0 and commitment='"&NYPonum&"'; "
rsscomp.close
next
'response.write sql666
'response.end
'sql666= "SET ANSI_WARNINGS OFF insert into po_file([Record ID], [Commitment ID], [Item number], [Description], [Retain percent], [Delivery  Date], [Scope of Work], [Job], [Extra], [Cost Code], [Category], [tax grp], [Tax], [units], [unit cost], [unit Description], [amount],[bought out],varies) select distinct '""CI""' as CI,'""'+convert(varchar,commitment)+'""','""""','""'+[description]+'""','""""','""""','""""','""'+convert(varchar,jobnum)+'""','""""','""'+'110006'+'""','""'+'005'+'""','""""','""""','""""','""""','""""','""'+convert(varchar,po_total)+'""','""""','""""' from po_temp where accepted=1 and closed=0"
rst3.Open sql666, cnn1
	set cmd3.ActiveConnection = cnn1
	cmd3.commandText = "RunPOs"
	cmd3.CommandType = adCmdStoredProc
	Set prm = cmd3.CreateParameter("company", adVarChar, adParamInput, 50)
    cmd3.Parameters.Append prm
	cmd3.Parameters("company") = "NY"
    cmd3.execute
	sql666=""
	End Select
next
end if


rst1.close
response.redirect "acctpoview.asp?POflag=1"
%>