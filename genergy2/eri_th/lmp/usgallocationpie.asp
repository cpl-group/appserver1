<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
Dim bldgnum, by, bp, varpercent, usgunit,totalBuildingUsg,totalBuildingBill,tenantname,misclabel,coor, explode,graphtype,  datacolors(14)

bldgnum		= Request("bldgnum")
by 			= Request("by")
bp 			= Request("bp")

if by = "" then by = 0
if bp = "" then bp = 0

varpercent 	= ".00"

'Colors
Dim tempcolor, tempR, tempG, tempB, count

        tempR = 0.3
        tempG = 0.61
        tempB = 0.79

        For count = 0 To 14 'static color generation
            tempR = (1999 / tempR) Mod 255
            tempG = (1777 / tempG) Mod 255
            tempB = (2003 / tempB) Mod 255
            datacolors(count) = RGB(CInt(tempR), CInt(tempG), CInt(tempB))
            tempR = Abs(Sin(tempR))
            tempG = Abs(Cos(tempG))
            tempB = Abs(Sin(tempB))
        Next 


Dim cnn, cmd, rs,sqlstr, prm

set cnn 	= server.createobject("ADODB.Connection")
set cmd 	= server.createobject("ADODB.Command")
set rs 		= server.createobject("ADODB.Recordset")

cnn.Open getLocalConnect(bldgnum)
cmd.CommandType 	= adCmdStoredProc
cnn.CursorLocation 	= adUseClient

Set cmd.ActiveConnection = cnn

    cmd.CommandText = "sp_tenant_usage_allocation"
    ' set parameter type and append for tenant contribution pie
    Set prm = cmd.CreateParameter("building", adChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("by", adinteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("bp", adinteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("percent", adchar, adParamInput,10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("buildingusg", adinteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("buildingamt", adinteger, adParamOutput)
    cmd.Parameters.Append prm
x=30
do until x <= 15 	
	cmd.Parameters("building") 	= bldgnum
    cmd.Parameters("by") 		= by
    cmd.Parameters("bp") 		= bp
    cmd.Parameters("percent") 	= varpercent
	
    set rs = cmd.execute
	totalBuildingUsg 			= cmd.Parameters("buildingusg")
	totalBuildingBill			= cmd.Parameters("buildingamt")
	usgunit 					= "KWH"
	if rs.recordcount > 15 then 
		x = rs.recordcount
		varpercent = varpercent + .01
		rs.close
	else 
		x = rs.recordcount
	end if 	
loop
'chart vars
dim objChart
set objChart = Server.CreateObject ("Dundas.ChartServer2D.2")
if not rs.EOF and (totalbuildingusg <> "" and totalbuildingbill <> "") then 
	dim dataset(), datanames(),index, leftover
	index = 0

	redim dataset(rs.recordcount)
	redim datanames(rs.recordcount)
	
	
	leftover=0
	'dim labelname ' for determining whether to use table column name "billingname" or 
	do while not(rs.EOF)
		if trim(cdbl(rs("totalamt"))) > 0 then 
				
			if lcase(rs("tenantname")) = "small tenants" then  
					misclabel = "Tenants under " & Formatpercent(varpercent,2) &" usage" 
			else 
					misclabel = rs("tenantname") 
			end if
'			datacolors(index) = cstr(255-index) & "," & cstr(255/index) & "," cstr((255-index)/index)
				objChart.AddData formatnumber(rs("usgpercent")),0, misclabel,datacolors(index)
				objChart.AddData formatnumber(rs("amtpercent")),1, misclabel,datacolors(index)

			datanames(index) = misclabel
			leftover=leftover+cDBL(dataset(index))
			index = index + 1
		end if
		rs.movenext()
	loop
  leftover=100-leftover
'  if leftover>0 then
'      dim loname
'      objChart.AddData formatnumber(leftover),0, "Non-Metered Load", RGB(200,200,200)
'      dataset(index) = leftover
'  end if
else
  objChart.AddStaticText "No Data Available",750,10,RGB(100,100,100),"Arial",14,1,2
  objChart.SendJPEG 650, 300, 50
  response.end
end if
rs.close

objChart.ChartArea(0).AddChart 0, 0, 0
objChart.ChartArea(1).AddChart 0, 1, 1
objChart.ChartArea(0).SetPosition 0, 50,350, 300
objChart.ChartArea(1).SetPosition 300,50,650, 300

	
objChart.AddStaticText "Usage Allocation",175, 20,RGB(100,100,100),"Arial",10,1,2
objChart.AddStaticText "Cost Allocation",475, 20,RGB(100,100,100),"Arial",10,1,2

objChart.Legend.Enabled = false
objChart.Legend.FontSize =8
objChart.Legend.SetPosition 400,50,650,300


objChart.AntiAlias


objChart.SendJPEG 650, 300, 0


%>