<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<%
dim bldg
dim rs, cnn, cmd,graphtype,strsql
bldg 		= trim(request("bldgid")) '"99"
graphtype	= trim(request("graphtype")) '"6"

'make chart

Set cnn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.recordset")

cnn.Open getConnect(0,0,"Engineering")

strsql = "select * from graph where bldgnum = '" & bldg &"' order by orderno"

rs.Open strsql, cnn,0
			dim objChart,label, tempdata
			set objChart = Server.CreateObject("Dundas.ChartServer2D.2")

if not rs.eof then 

			while not rs.EOF 						
				tempdata = rs("wsqft")
				label = rs("floor")
				objChart.AddData tempdata, 0, label
				label = rs("wsqft")
				objChart.AddData 0, 1, label	
				rs.movenext
				
			wend 
			rs.close
			'objChart.SetColorFromPoint (0)
			objChart.SetSeriesColor 0, RGB(40,120,255)	'first series, light blue
			objChart.ChartArea(0).AddChart graphtype,0,0, , 0
			'objChart.ChartArea(1).AddChart graphtype,1,1, , 0
			objChart.Rectangle3DEffect()
			objChart.ChartArea(0).Axis(1).Angle = 90
			objChart.ChartArea(0).Axis(0).enabled=true
			objChart.ChartArea(0).Axis(1).enabled=true
			objChart.ChartArea(0).Axis(1).title="FLOOR" 
			objChart.ChartArea(0).Axis(0).title="WATT PER SQFT"
			objChart.ChartArea(0).Axis(2).enabled=false
			objChart.ChartArea(0).Axis(3).enabled=false
			objChart.ChartArea(0).Axis(0).color=RGB(00,00,00)
			objChart.ChartArea(0).Axis(1).color=RGB(00,00,00)
			objChart.ChartArea(0).Axis(2).color=RGB(00,00,00)
			objChart.ChartArea(0).Axis(3).color=RGB(00,00,00)
			objChart.ChartArea(0).LineWidth = 3
			objChart.ChartArea(0).GridVEnabled = false
			objChart.ChartArea(0).GridHEnabled = true
			'objChart.ChartArea(0).Axis(0).SetNumberFormat 2,2
			objChart.AntiAlias()
			objChart.ChartArea(0).LineWidth = 2
			objChart.BackgroundColor= RGB(255,255,255)
			objChart.ChartArea(0).BackgroundColor = RGB(255,255,255)
			'Chart 2 setup
			objChart.ChartArea(1).AddChart graphtype,1,1, , 0
			objChart.Rectangle3DEffect()
			objChart.ChartArea(1).Axis(3).Angle = 90
			objChart.ChartArea(1).Axis(0).enabled=false
			objChart.ChartArea(1).Axis(1).enabled=false
			objChart.ChartArea(1).Axis(2).enabled=false
			objChart.ChartArea(1).Axis(3).enabled=true
			'objChart.ChartArea(1).Axis(3).title="WATTS PER SQFT"
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

			objChart.AddStaticText "WATT PER SQFT",0,300,RGB(00,00,00),"Arial",8,1,2,90
			'objChart.AddStaticText rateLabel &" For "& bldglabel &". Total Cost For " &lmpdate& " " & costlabel & " " & formatcurrency(x,2),10,0,RGB(100,100,100),"Arial",8,1
			'objChart.AddStaticText detaillabel,0,289,RGB(100,100,100),"Arial",7,1
			objChart.AddStaticText "FLOOR",400,580,RGB(00,00,00),"Arial",8,1,2

'			if x = 0 then 
'			objChart.AddStaticText by1,525,0,RGB(100,100,100),"Arial",8,1
'			else
'			objChart.AddStaticText by1 & " vs " & by2,480,0,RGB(100,100,100),"Arial",8,1
'			end if
			
			objChart.ChartArea(0).SetPosition 100,75,750,500
			objChart.ChartArea(1).SetPosition 100,75,750,500
else
			objChart.AddStaticText "No Data Available For Selected Dates" ,300,15,RGB(0,0,100),"Arial",13,1,2
			objChart.BackgroundColor = RGB(255,255,255)
end if
objChart.SendJpeg 800,600

set cnn = nothing
%>
