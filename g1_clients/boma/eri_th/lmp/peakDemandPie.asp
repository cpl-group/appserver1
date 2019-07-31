<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim b, explode, coor, byear, bperiod, luid
b = request.querystring("b")
luid = request.querystring("luid")
byear = request.querystring("byear")
bperiod = request.querystring("bperiod")
explode = request.querystring("explode")
coor = request.querystring("coor")
'ado vars
dim cnn, cmd, rs, prm
set cnn = server.createobject("ADODB.Connection")
set cmd = server.createobject("ADODB.Command")
set rs = server.createobject("ADODB.Recordset")
cnn.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

dim graphtypename, tenantname
if trim(luid)<>"" then
    graphtypename = "Tenant "
    rs.open "SELECT BillingName FROM tblLeases WHERE billingId in (SELECT billingid FROM tblLeasesUtilityPrices WHERE leaseutilityid="& luid &")", cnn
    if not(rs.eof) then
        graphtypename = rs("BillingName")
        rs.close
        
    end if
else
    graphtypename = "Building "
end if

'chart vars
dim objChart
set objChart = Server.CreateObject ("Dundas.ChartServer2D.2")


cnn.CursorLocation = adUseClient
cmd.CommandType = adCmdStoredProc

' assign internal name to stored procedure
cmd.Name = "test"

if trim(luid)<>"" then
    cmd.CommandText = "sp_peak_metercontribution"
    ' set parameter type and append for tenant contribution pie
    Set prm = cmd.CreateParameter("bldg", adChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("lid", adChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("byear", adinteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("bperiod", adinteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("max", adChar, adParamOutput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("by", adinteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("bp", adinteger, adParamOutput)
    cmd.Parameters.Append prm

    Set cmd.ActiveConnection = cnn
    'return set to recordset rs
    cnn.test b, luid, byear, bperiod, rs
else
    cmd.CommandText = "sp_peak_contribution"
    ' set parameter type and append for building contribution pie
    Set prm = cmd.CreateParameter("bldg", adChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("byear", adinteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("bperiod", adinteger, adParamInput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("max", adChar, adParamOutput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("dmax", adDBTimeStamp, adParamOutput, 11)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("tmax", adChar, adParamOutput, 8)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("by", adinteger, adParamOutput)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("bp", adinteger, adParamOutput)
    cmd.Parameters.Append prm

    Set cmd.ActiveConnection = cnn
    'return set to recordset rs
    cnn.test b, byear, bperiod, rs
	dim dmaxdate
	dmaxdate = left(cmd.Parameters("dmax")&" ",instr(cmd.Parameters("dmax")&" ", " ")-1)
end if



dim dataset(), datanames(), datademand(), dataluid(), index, leftover
index = 0
redim dataset(rs.recordcount)
redim datanames(rs.recordcount)
redim datademand(rs.recordcount)
redim dataluid(rs.recordcount)
leftover=0
'dim labelname ' for determining whether to use table column name "billingname" or 
do while not(rs.EOF)
    if formatnumber(rs("percentage"))>0 then 
        objChart.AddData formatnumber(rs("percentage")),0, "Demo Tenant" 
        dataset(index) = rs("percentage")
        datanames(index) = rs("labelname")
        datademand(index) = rs("demand")
        if trim(luid)="" then dataluid(index) = rs("leaseutilityid")
        leftover=leftover+cDBL(dataset(index))
        index = index + 1
    end if
    rs.movenext()
loop
rs.close
leftover=100-leftover
if leftover>0 then
    dim loname
    objChart.AddData formatnumber(leftover),0, "Non-Submetered Load", RGB(200,200,200)
    dataset(index) = leftover
    datanames(index) = "Non-Submetered Load"
    datademand(index) = "0"
    dataluid(index) = "0"
end if
objChart.ChartArea(0).AddChart 0, 0, 0
objChart.ChartArea(0).SetPosition 50, 50, 550, 260
if coor <> "" then
   coor = Mid(coor,2)
    
   dim ArrXYposition, Xposition, Yposition
   ArrXYposition = split(coor , ",")

   Xposition = cint(trim(ArrXYposition(0)))
   Yposition = cint(trim(ArrXYposition(1)))

   objChart.Selection 600, 310, Xposition, Yposition

   if objChart.SelectedDataSeries <> -1 and objChart.SelectedDataPoint <> -1 then
      objChart.SetExploded 0, objChart.SelectedDataPoint
      explode = objChart.SelectedDataPoint
   else
      explode = -1
   end if

elseif explode>-1 and explode<>"" then
    objChart.SetExploded 0, cint(explode)
end if

objChart.AddStaticText graphtypename& " Peak Demand Contributions",300, 10,RGB(100,100,100),"Arial",14,1,2
if explode<>"" and explode<>"-1" then
    objChart.AddStaticText "Demo Tenant",10,260,RGB(100,100,100),"Arial",8,1
    objChart.AddStaticText formatnumber(dataset(explode))&" percent of total",10,270,RGB(100,100,100),"Arial",8,1
    rs.open "SELECT meternum,kwh FROM pulse_"&b&" p INNER join meters on p.meterid=meters.meterid WHERE meters.LeaseUtilityId="& dataluid(explode) &" and datediff(minute,[date],'"& dmaxdate &" "& cmd.Parameters("tmax") &"')=0 ORDER BY kwh desc", cnn
    if not(rs.EOF) and cDBL(datademand(explode))>0 then 'create table of individual tenant meter info
        dim chY, chInterval, chLoop, chLimit
        chY=70
        chInterval=10
        chLoop=0
        chLimit=22
        objChart.AddStaticText "Tenant Contribution Details",430,58,RGB(100,100,100),"Arial",8,1
        objChart.AddStaticText "Meter Number",430,chY,RGB(100,100,100),"Arial",7,1
        objChart.AddStaticText "Demand",540,chY,RGB(100,100,100),"Arial",7,1,1
        objChart.AddStaticText "Percentage",600,chY,RGB(100,100,100),"Arial",7,1,1
        do Until rs.EOF or chLimit < rs.AbsolutePosition
            dim tempmeter
            tempmeter = rs("meternum")
            if len(tempmeter)>11 then tempmeter=left(tempmeter,9)&"..."
            chY = chY+chInterval
            objChart.AddStaticText tempmeter,430,chY,RGB(100,100,100),"Arial",7,1
            objChart.AddStaticText (cDBL(rs("kwh"))*4),540,chY,RGB(100,100,100),"Arial",7,1,1
            'response.write datademand(explode)
            'response.end
            objChart.AddStaticText formatnumber(cDBL(rs("kwh"))*400/cDBL(datademand(explode)))&"%",600,chy,RGB(100,100,100),"Arial",7,1,1
'            response.write datademand(explode)
'            response.end
            rs.movenext()
        loop
        if chLimit < rs.AbsolutePosition then 'display that there are more meter not shown
            chY = chY+chInterval
            objChart.AddStaticText "More meters not shown...",430,chY,RGB(100,100,100),"Arial",7,1
        end if
    end if
    rs.close()
else
    objChart.Legend.Enabled = true
    objChart.Legend.FontSize = 6
    objChart.Legend.SetPosition 450,35,600,310 
end if

objChart.AntiAlias


objChart.SendJPEG 600, 310, 50


%>