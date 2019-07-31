<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim storedproc, by1, by2, bldg, FunctionLabel,bldglabel, rgblist,y,graphtype, annBudget
dim rs, cnn, cmd,prm,rpt, NumberFormat,NameStart,rpttype,vardollar,varpercent,applyvar,PerChange(12,2),axisvariable,tempArray

tempArray 		= split(trim(request("storedproc")),"|")
storedproc 		= tempArray(0)
FunctionLabel 	= tempArray(1)
bldg 			= trim(request("bldg"))
bldglabel		= ucase(trim(request("bldgname")))
by1 			= trim(request("by1"))
by2 			= trim(request("by2"))
annBudget		= trim(request("budget"))

applyvar 	= trim(request("applyvariant"))

if by2 = "" then 
	by2 = null
end if 

graphtype	= request("graphtype")
rpt 		= trim(request("rpt"))

NameStart	= Len(rpt) + 4 
axisvariable= 0
if applyvar = "true" then 
	varpercent	= trim(request("varpercent"))
	vardollar	= trim(request("vardollar"))
	if vardollar = "" then 
		vardollar = 0
	else
		if varpercent = "" then 
			varpercent = 0
		else
		end if
	end if
else
	varpercent = 0 
	vardollar = 0
end if


'FunctionLabel = Mid(replace(storedproc, "_", " "), NameStart)

'make chart

Set cnn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.recordset")
set cmd = server.createobject("ADODB.Command")

cnn.Open getLocalConnect(bldg)

cnn.CursorLocation = adUseClient
cmd.CommandType = adCmdStoredProc
cmd.Name = "get"
if trim(by1)<>"" then
    cmd.CommandText = storedproc

    Set prm = cmd.CreateParameter("bldg", adChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("by1", adChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("by2", adinteger, adParamInput)
    cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("vardollar", addouble, adParamInput,18,4)
    cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("varpercent", addouble, adParamInput,18,2)
    cmd.Parameters.Append prm	
	Set prm = cmd.CreateParameter("rpttype", adVarChar, adParamOutput, 50)
	cmd.Parameters.Append prm
    Set cmd.ActiveConnection = cnn
    'return set to recordset rs
	 cmd.Parameters("bldg") = bldg
	 cmd.Parameters("by1") = by1
	 cmd.Parameters("by2") = by2
	 cmd.Parameters("vardollar") = vardollar
	 cmd.Parameters("varpercent") = varpercent

'	 response.write "exec "&storedproc&" "&cmd.Parameters("bldg")&","&cmd.Parameters("by1")&","&cmd.Parameters("by2")&","&cmd.Parameters("vardollar")&","&cmd.Parameters("varpercent")
'	 response.end

	 cnn.get rs
end if


x=0
if rs.state = adstateopen then
if not rs.eof then 
				dim objChart, cf
				set objChart = Server.CreateObject("Dundas.ChartServer2D.2")
				rpttype = cmd.Parameters("rpttype")
				cf 		= 0
				while not rs.EOF 
				dim label, tempdata
				for y=1 to 12 
				Select case lcase(rpttype)
				
				 Case "cost" 
				
						tempdata 		= formatcurrency(rs("p"&y),4)
						NumberFormat 	= 2
					
				
				 Case "forecast"
				 	if lcase(trim(rs("billyear"))) <> "forecast" then 
						tempdata = formatcurrency(rs("p"&y),4)
						NumberFormat 	= 2
					else
						if rs("p"&y) = 0 then 
							cf 				= cf + PerChange(y,1) 
							tempdata 		= cf
							NumberFormat 	= 2
						else
							cf 				= cf + formatcurrency(rs("p"&y),0)  
							tempdata 		= cf
							NumberFormat 	= 2
						end if
					end if				
				 case "operations"
						tempdata 		= formatnumber(rs("p"&y),4)
						NumberFormat 	= 1
				 Case else
					tempdata = formatnumber(rs("p"&y))
					NumberFormat = 1
				End Select
				if tempdata > axisVariable then 
					AxisVariable = tempdata
				end if
				label = Mid(MonthName(y),1,3)
				objChart.AddData tempdata, x, label 
				PerChange(y,x) = rs("p"&y)
				next
				if x = 0 then 
					by1 = rs("billyear")
				end if
				rs.movenext
				if not rs.eof then
						x=x+1
'						if rs("billyear") <> "Forecast" then 
							by2 = rs("billyear")
'						end if
				end if
			wend 
			rs.movefirst
			objChart.Rectangle3DEffect()
			objChart.ChartArea(0).Axis(1).Angle = 0
			objChart.ChartArea(0).Axis(0).enabled=true
			objChart.ChartArea(0).Axis(1).enabled=true
			objChart.ChartArea(0).Axis(2).enabled=false
			objChart.ChartArea(0).Axis(0).color=RGB(00,00,00)
			objChart.ChartArea(0).Axis(1).color=RGB(00,00,00)
			objChart.ChartArea(0).Axis(2).color=RGB(00,00,00)
			objChart.ChartArea(0).Axis(3).color=RGB(00,00,00)
			objChart.ChartArea(0).GridVEnabled = true
			objChart.ChartArea(0).GridHEnabled = true
			Select case lcase(rpttype)
				 Case "cost"  
						if x = 1 then 
							objChart.SetSeriesColor 0, rgb(192,192,192)
							objChart.SetSeriesColor 1, rgb(00,153,255)
						else
							objChart.SetSeriesColor 1, rgb(192,192,192)
							objChart.SetSeriesColor 0, rgb(00,153,255)			
						end if
						objChart.ChartArea(0).AddChart graphtype,0,x, , 0
						objChart.ChartArea(0).LineWidth = 3
						if x = 0 then 
							objChart.AddStaticText by1,525,0,RGB(100,100,100),"Arial",8,1
						else
							objChart.AddStaticText by1 & " vs " & by2,480,0,RGB(100,100,100),"Arial",8,1
						end if				 
				Case "forecast"
				 		x=x+1
						for y = 1 to 12 
							objChart.AddData annBudget, x, label 
						next

						if x > 0 then 
							objChart.SetSeriesColor 2, rgb(128,0,0)
							objChart.SetSeriesColor 1, rgb(00,153,255)
							objChart.SetSeriesColor 3, rgb(00,255,0)
						else
							objChart.SetSeriesColor 1, rgb(192,192,192)
							objChart.SetSeriesColor 0, rgb(00,153,255)			
						end if
				 		objChart.ChartArea(0).AddChart graphtype,1,1, , 0
				 		if x>=2 then objChart.ChartArea(0).AddChart 1,2,2, , 0
						if x>=3 then objChart.ChartArea(0).AddChart 1,3,3, , 0
						objChart.ChartArea(0).LineWidth = 1
						if by2 <> "" then 
						objChart.AddStaticText by2,525,0,RGB(100,100,100),"Arial",8,1
						end if 
				 Case Else
						if x = 1 then 
							objChart.SetSeriesColor 0, rgb(192,192,192)
							objChart.SetSeriesColor 1, rgb(00,153,255)
						else
							objChart.SetSeriesColor 1, rgb(192,192,192)
							objChart.SetSeriesColor 0, rgb(00,153,255)			
						end if
						objChart.ChartArea(0).AddChart graphtype,0,x, , 0
						objChart.ChartArea(0).LineWidth = 3				 		
			End Select
			
				
			if cdbl(axisVariable) < 1 then
				objChart.ChartArea(0).Axis(0).SetNumberFormat NumberFormat,4
			else
				if cdbl(axisVariable) > 1000 then 
				objChart.ChartArea(0).Axis(0).SetNumberFormat NumberFormat,0
				else
				objChart.ChartArea(0).Axis(0).SetNumberFormat NumberFormat,2
				end if
			end if
			objChart.AntiAlias()
			objChart.ChartArea(0).LineWidth = 2
			objChart.BackgroundColor= RGB(238,238,238)
			objChart.ChartArea(0).BackgroundColor = RGB(238,238,238)
			objChart.AddStaticText FunctionLabel,0,100,RGB(100,100,100),"Arial",8,1,2,90
			objChart.AddStaticText ucase(FunctionLabel) &" FOR "&bldglabel,70,0,RGB(100,100,100),"Arial",8,1	
			
			'--------------------------------------------
			' setup the legend
			'--------------------------------------------
			objChart.Legend.Enabled = true  'enable the legend (it is disabled by default)
			objChart.Legend.FontSize = 8
			Select case rpttype
				 Case "cost"  
					if x=0 then
						objChart.Legend.Add by1, rgb(00,153,255)
						objChart.Legend.SetPosition 70,235,300,280 
					else
						objChart.Legend.Add by1, rgb(192,192,192)
						objChart.Legend.Add by2, rgb(00,153,255)
						objChart.Legend.SetPosition 70,235,300,280 
					end if
				 Case "forecast"
				 		if cf <> 0 and annbudget > 0 then 		
							tempdata = formatcurrency(cf,0) & " ("&formatpercent(((cf-annbudget)/cf))&")"
						else
							tempdata = formatcurrency(cf,0) 
						end if
						objChart.Legend.Add "Current Year: " & by2, rgb(00,153,255)
						objChart.Legend.Add "Annual Energy Budget: " & formatcurrency(annBudget,0), rgb(0,255,0)
						objChart.Legend.Add "Forecasted Annual Energy Cost Assuming Current Usage Patterns: " & tempdata, rgb(128,0,0)
						objChart.Legend.SetPosition 70,235,600,280 
				 Case Else
					if x=0 then
						objChart.Legend.Add by1, rgb(00,153,255)
						objChart.Legend.SetPosition 70,235,300,280 
					else
						objChart.Legend.Add by1, rgb(192,192,192)
						objChart.Legend.Add by2, rgb(00,153,255)
						objChart.Legend.SetPosition 70,235,300,280 
					end if
			End Select
			'optional legend settings
			objChart.Legend.BorderColor = RGB(110,0,0)
			objChart.Legend.BackgroundColor = RGB(230,230,230)
			objChart.Legend.FontColor = RGB(0,0,110)
			objChart.Legend.FontSize = 8
			objChart.Legend.Transparent = true 'set to false, so that the background color
												' can be seen
			
			objChart.ChartArea(0).SetPosition 70,20,550,210
			objChart.SendJpeg 600,300
else
	response.write "<body bgcolor=eeeeee> NO DATA AVAILABLE FOR <br>" & bldglabel & " FOR SELECTED DATES</body>"

			end if


end if

set cnn = nothing
%>
