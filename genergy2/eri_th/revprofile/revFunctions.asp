<%'functions for load revenue/expense data in to arrays 
  'used in MakeChart.asp, monthlyDetails.asp
dim farev, aexp, eri, exps, subm, urae, urar, mac, plp, net, preferences, pidsession
eri		=1
exps	=2
subm	=3
urar	=4
urae	=5
mac		=6
plp		=7
net		=8
dim ArrPrefs(8)

dim cnn1, rst1, rst2, cmd, prm
Set rst1 = Server.CreateObject("ADODB.recordset")
set cnn1 = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rst2 = server.createobject("ADODB.Recordset")
if trim(request("bldg"))<>"" then cnn1.Open getConnect(0,request("bldg"),"billing")
cnn1.CursorLocation = adUseClient
cmd.CommandText = "sp_RevProfile"
cmd.CommandType = adCmdStoredProc

Set prm = cmd.CreateParameter("bldg", adVarChar, adParamInput, 20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("util", adVarChar, adParamInput, 20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("by", adVarChar, adParamInput, 10)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("eri", adTinyInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("exp", adTinyInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("subm", adTinyInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("urar", adTinyInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("urae", adTinyInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("mac", adTinyInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("plp", adTinyInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("net", adTinyInt, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("user", adVarChar, adParamInput, 50)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("ngain", adVarChar, adParamOutput,20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("texp", adVarchar, adParamOutput,20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("nett", adVarChar, adParamOutput, 20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("bldgnum", adVarChar, adParamOutput, 20)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("byear", adInteger, adParamOutput)
cmd.Parameters.Append prm
cmd.Name = "test"
'Set cmd.ActiveConnection = cnn1

sub getdataSets(byear, bldg, utype, download)
'response.write getLocalConnect(bldg)
'response.end
  set cnn1 = server.createobject("ADODB.Connection")
  cnn1.open getConnect(0,bldg,"billing")
  Set cmd.ActiveConnection = cnn1
	if download=0 then initarrays()
	checkprefs()
	pidsession=getPID(bldg)
  	'response.write " exec sp_RevProfile "&bldg&", "&utype&", "&byear&", "&ArrPrefs(1)&", "&ArrPrefs(2)&", "&ArrPrefs(3)&", "&ArrPrefs(4)&", "&ArrPrefs(5)&", "&ArrPrefs(6)&", "&ArrPrefs(7)&", "&ArrPrefs(8)&", 1, "&session("pid")&"<br>"
  	'response.end
  cnn1.test bldg, utype, byear, ArrPrefs(1), ArrPrefs(2), ArrPrefs(3), ArrPrefs(4), ArrPrefs(5), ArrPrefs(6), ArrPrefs(7), 1,pidsession, rst1
'	if download=0 then
		Dim recordnum
		recordnum=0
		do until rst1.eof' or recordnum=12
		  recordnum=cint(rst1("billperiod"))
			if recordnum<13 and recordnum>0 then
				if ArrPrefs(eri) then 	ArrDataSeriesERI(recordnum)=ArrDataSeriesERI(recordnum)+rst1("eri_rev")/1000
				if ArrPrefs(exps) then 	ArrDataSeriesExpenses(recordnum)=ArrDataSeriesExpenses(recordnum)+rst1("Expenses")/1000
				if ArrPrefs(subm) then	ArrDataSeriesSubmetered(recordnum)=ArrDataSeriesSubmetered(recordnum)+clng(rst1("Submetered"))/1000
				if ArrPrefs(mac) then	ArrDataSeriesMac(recordnum)=ArrDataSeriesMac(recordnum)+clng(rst1("Mac_rev"))/1000
				if ArrPrefs(plp) then	ArrDataSeriesPlp(recordnum)=ArrDataSeriesPlp(recordnum)+clng(rst1("PLP"))/1000
				if ArrPrefs(urae) then 	ArrDataSeriesUnreportedExp(recordnum)=ArrDataSeriesUnreportedExp(recordnum)+clng(rst1("UnreportedEXPAmt"))/1000
				if ArrPrefs(urar) then 	ArrDataSeriesUnreportedRev(recordnum)=ArrDataSeriesUnreportedRev(recordnum)+clng(rst1("UnreportedRevAmt"))/1000
				if ArrPrefs(net)  then 	
					dim tempnet
					tempnet = 0
					if trim(rst1("net"))<>"" then tempnet = clng(rst1("net"))/1000
					ArrDataSeriesNet(recordnum)=ArrDataSeriesNet(recordnum)+tempnet
				end if
			end if
			rst1.movenext
		loop
'end if
	rst1.close
  cnn1.close
end sub

sub initarrays()
	for i=1 to 12
		ArrDataSeriesERI(i)=0
		ArrDataSeriesSubMetered(i)=0
		ArrDataSeriesExpenses(i)=0
		ArrDataSeriesUnreportedExp(i)=0
		ArrDataSeriesUnreportedRev(i)=0
		ArrDataSeriesMac(i)=0
		ArrDataSeriesPlp(i)=0
		if i<4 then ArrPieRevenue(i)=0
		if i<3 then ArrPieExpenses(i)=0
	next
end sub

function checkprefs()
	dim i
	checkprefs=0
	for i = 1 to 8
		ArrPrefs(i) = 0
	next
	if session("ERI")=1 and (utype=2 or utype=0) then									'eri
		ArrPrefs(eri)=1
		checkprefs = checkprefs + 1
	end if
	if session("Expenses")=1 then								'exp
	ArrPrefs(exps)=1
		checkprefs = checkprefs + 1
	end if
	if session("Submeter")=1 then								'subm
		ArrPrefs(subm)=1
		checkprefs = checkprefs + 1
	end if
	if session("Revenue_Adjustments")=1 then			'ura
		ArrPrefs(urar)=1
		checkprefs = checkprefs + 1
	end if
	if session("Expense_Adjustments")=1 then			'ura
		ArrPrefs(urae)=1
		checkprefs = checkprefs + 1
	end if
	if session("Mac_Revenue")=1 and (utype=2 or utype=0) then							'mac
		ArrPrefs(mac)=1
		checkprefs = checkprefs + 1
	end if
	if session("PLP_Revenue")=1 and (utype=2 or utype=0) then							'plp
		ArrPrefs(plp)=1
		checkprefs = checkprefs + 1
	end if
'	if session("Net")=1 then							'net
		ArrPrefs(net)=1
		checkprefs = checkprefs + 1
'	end if
	if instr(request.querystring("preferences"),",")>0 then
		preferences = split(request.querystring("preferences"),",")
		ArrPrefs(0) = preferences(0)
		ArrPrefs(1) = preferences(1)
		ArrPrefs(2) = preferences(2)
		ArrPrefs(3) = preferences(3)
		ArrPrefs(4) = preferences(4)
		ArrPrefs(5) = preferences(5)
		ArrPrefs(6) = preferences(6)
		ArrPrefs(7) = preferences(7)
		ArrPrefs(8) = preferences(8)
		pidsession=preferences(9)
	end if
end function
%>