<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim by, lmpdate,lmpdate1,lmpdate2, bldg, meterid, x, rgblist,y,graphtype,region, ratelabel, bldglabel, qrytype,costlabel
dim rs, cnn, cmd,prm,rpt, NumberFormat,NameStart,rpttype,strsql,pid
bldg 		= trim(request("bldgid")) '"99"
lmpdate		= cdate(trim(request("lmpdate")))   '"02/19/2003")
lmpdate1 	= cdate(lmpdate) + 1 
meterid 	= trim(request("meterid")) '"3257"
region 		= "N.Y.C."
qrytype		= trim(request("qrytype"))

graphtype	= trim(request("graphtype")) '"6"

'make chart

Set cnn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.recordset")

cnn.Open getLocalConnect(bldg)

strsql = "select strt,pid from buildings where bldgnum = '" & bldg &"'"

rs.Open strsql, cnn,0

if not rs.eof then 
	bldglabel 	= rs("strt")
	pid 		= rs("portfolioid") 
end if
rs.close

if meterid = "" then 
	strsql = "select meterid from meters where bldgnum = '" & bldg &"' and pp = 1"
	rs.Open strsql, cnn,0
	
	if not rs.eof then 
		meterid = rs("meterid")
	end if
	rs.close

end if

select case lcase(qrytype)

	case "actual"
		
		strsql= "select datepart(hour,date) as hour,sum(kwh) as kwh, rate as rate, sum(kwh) * (rate) as appliedrate,ratetype  from (select top 1 case when a.rate is null then b.rate else a.rate end as rate, case when a.rate is null then 'avg' else 'bill' end as ratetype from (select avgkwh as rate from utilitybill where ypid in (select top 1 ypid from billyrperiod where bldgnum = '"&bldg&"' and utility = 2 and '"&lmpdate&"' between datestart and dateend)) as a full join (select avg(avgkwh) as rate from utilitybill where ypid in (select top 3 ypid from billyrperiod where bldgnum = '"&bldg&"' and utility = 2 and dateend < '"&lmpdate&"' order by dateend desc)) as b on a.rate = b.rate) as Ratetable, pulse_"&bldg&" where meterid = "&meterid&" and (date >= '"&lmpdate&"' and date < '"&lmpdate1&"')  group by datepart(hour,date), rate,ratetype order by datepart(hour,date)"
		ratelabel 	= "Actual Cost"
		if lmpdate = date() then 
			costlabel 	= "Is"
		else
			costlabel 	= "Was"
		end if
		detaillabel = "Rate comprised of the prior 3 month actual average unit cost per KWHR inclusive of T&D and Commodity, exclusive of sales tax."
	case "dam"
		strsql = "select datepart(hour,date) as hour, sum(kwh) as kwh, avgkwh+rate as rate, sum(kwh) * (avgkwh+rate) as appliedrate, 'avg' as ratetype  from (SELECT avg((case tdwithtax when '1' then (tdtotalamt - tdsalesamt) else tdtotalamt end)/totalkwh) as avgkwh FROM UtilityBill WHERE (ypId IN (select top 3 ypid from billyrperiod where bldgnum = '"&bldg&"' and utility = 2 and dateend < '"&lmpdate&"' order by dateend desc))) as kwhtable, pulse_"&bldg&" inner join (select b.hour, ((rateA + rateB) / 1000) rate from (select top 24 datepart(hour,date) as hour, (tenminsr + tenminnsr + thirtyminOR + regulation) as rateB from ["&getpidip(pid)&"].mainmodule.dbo.real_time_price_aux where (date >= '"&lmpdate&"' and date <'"&lmpdate1&"') order by date) as b inner join (select top 24 datepart(hour,date) as hour, lbmp as rateA from ["&getpidip(pid)&"].mainmodule.dbo.real_time_price where (date >= '"&lmpdate&"' and date <'"&lmpdate1&"')  and name = '"&region&"' order by hour) as A on a.hour = b.hour) as Ratetable on ratetable.hour = datepart(hour,pulse_"&bldg&".date) where meterid = "&meterid&" and (date >= '"&lmpdate&"' and date <'"&lmpdate1&"')  group by datepart(hour,date), Ratetable.rate, avgkwh order by datepart(hour,date)"
		ratelabel 	= "NYISO LBMP for Zone J Comparison"
		costlabel 	= "Would Have Been"
		detaillabel = "Rate comprised of NYISO LBMP for Zone J inclusive of ancillary charges plus 3 moth actual average T&D costs, exclusive of sales tax"
	case else
		Response.write "No Query Type Provided"
		response.end
	end select
'response.write strsql
'response.end

rs.Open strsql, cnn,0

x=0
dim objChart,label, tempdata, detaillabel
set objChart = Server.CreateObject("Dundas.ChartServer2D.2")
if not rs.eof then 
			while not rs.EOF 
				select case lcase(rs("ratetype"))
					case "bill"
						detaillabel = "Rate comprised of the actual average unit cost per KWHR inclusive of T&D and Commodity, exclusive of sales tax."					
				end select 						
				tempdata = rs("kwh")
				label = rs("hour")
				objChart.AddData tempdata, 0, label
				label = formatcurrency(rs("appliedrate"),2)
				x = x + label
				objChart.AddData 0, 1, label	
				rs.movenext
				
			wend 
			rs.close
			'objChart.SetColorFromPoint (0)
			objChart.SetSeriesColor 0, RGB(40,120,255)	'first series, light blue
			objChart.ChartArea(0).AddChart graphtype,0,0, , 0
			'objChart.ChartArea(1).AddChart graphtype,1,1, , 0
			objChart.Rectangle3DEffect()
			objChart.ChartArea(0).Axis(1).Angle = 0
			objChart.ChartArea(0).Axis(0).enabled=true
			objChart.ChartArea(0).Axis(1).enabled=true
			'objChart.ChartArea(0).Axis(1).title="HOUR" 
			objChart.ChartArea(0).Axis(0).title="KILO WATT HOURS"
			objChart.ChartArea(0).Axis(2).enabled=false
			objChart.ChartArea(0).Axis(3).enabled=false
			objChart.ChartArea(0).Axis(0).color=RGB(00,00,00)
			objChart.ChartArea(0).Axis(1).color=RGB(00,00,00)
			objChart.ChartArea(0).Axis(2).color=RGB(00,00,00)
			objChart.ChartArea(0).Axis(3).color=RGB(00,00,00)
			objChart.ChartArea(0).LineWidth = 3
			objChart.ChartArea(0).GridVEnabled = true
			objChart.ChartArea(0).GridHEnabled = true
			'objChart.ChartArea(0).Axis(0).SetNumberFormat 2,2
			objChart.AntiAlias()
			objChart.ChartArea(0).LineWidth = 2
			objChart.BackgroundColor= RGB(255,255,255)
			objChart.ChartArea(0).BackgroundColor = RGB(255,255,255)
			'Chart 2 setup
			objChart.ChartArea(1).AddChart graphtype,1,1, , 0
			objChart.Rectangle3DEffect()
			objChart.ChartArea(1).Axis(3).Angle = 45
			objChart.ChartArea(1).Axis(0).enabled=false
			objChart.ChartArea(1).Axis(1).enabled=false
			objChart.ChartArea(1).Axis(2).enabled=false
			objChart.ChartArea(1).Axis(3).enabled=true
			objChart.ChartArea(1).Axis(3).title="HOURLY RATES"
			objChart.ChartArea(1).Axis(0).color=RGB(00,00,00)
			objChart.ChartArea(1).Axis(1).color=RGB(00,00,00)
			objChart.ChartArea(1).Axis(2).color=RGB(00,00,00)
			objChart.ChartArea(1).Axis(3).color=RGB(00,00,00)
			objChart.ChartArea(1).LineWidth = 3
			objChart.ChartArea(1).GridVEnabled = false
			objChart.ChartArea(1).GridHEnabled = false
			objChart.ChartArea(1).Transparent = true
			objChart.ChartArea(1).Axis(3).SetNumberFormat 2,2
			objChart.AntiAlias()
			objChart.ChartArea(1).LineWidth = 2
			objChart.BackgroundColor= RGB(255,255,255)
			objChart.ChartArea(1).BackgroundColor = RGB(255,255,255)

'			objChart.AddStaticText rateLabel,0,100,RGB(100,100,100),"Arial",8,1,2,90
			objChart.AddStaticText rateLabel &" For "& bldglabel &". Total Cost For " &lmpdate& " " & costlabel & " " & formatcurrency(x,2),10,0,RGB(100,100,100),"Arial",8,1
			objChart.AddStaticText detaillabel,0,289,RGB(100,100,100),"Arial",7,1
			objChart.AddStaticText "HOUR",300,275,RGB(100,100,100),"Arial",7,1

'			if x = 0 then 
'			objChart.AddStaticText by1,525,0,RGB(100,100,100),"Arial",8,1
'			else
'			objChart.AddStaticText by1 & " vs " & by2,480,0,RGB(100,100,100),"Arial",8,1
'			end if
			
			objChart.ChartArea(0).SetPosition 70,80,620,255
			objChart.ChartArea(1).SetPosition 70,80,620,255
else
			objChart.AddStaticText "No Data Available For Selected Dates" ,300,15,RGB(0,0,100),"Arial",13,1,2
			objChart.BackgroundColor = RGB(255,255,255)
end if
objChart.SendJpeg 650,300

set cnn = nothing
%>
