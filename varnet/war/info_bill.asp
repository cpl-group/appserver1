<%@ LANGUAGE="VBSCRIPT" %>

<html>
<!-- #include file ="adovbs.inc" -->
<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<title>New Page 1</title>

<meta name="Microsoft Theme" content="none, default">
</head><style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>


<%
bldg = Request("bldg")
yr = request("yr")
pr = request("pr")

Set cnn1 = Server.CreateObject("ADODB.Connection")

serv = "10.0.7.16"
'serv = "home"

cnn1.Open "driver={SQL Server};server=" & serv & ";uid=genergy1;pwd=g1appg1;database=genergy1;"
%>

<%
Set cmd = Server.CreateObject("ADODB.command")
set rs12 = Server.CreateObject("ADODB.recordset")

cmd.ActiveConnection = cnn1
cmd.CommandText = "qpaLidByBldg"
cmd.CommandType = 4
Set prm1 = cmd.CreateParameter("bldg", adVarChar, adParamInput, 10, bldg)
cmd.Parameters.Append prm1
Set rs12 = cmd.Execute
Do While Not rs12.EOF

demand = calcdemandprice(bldg,rs12(0),yr,pr)
%>

<body>

<table border="1" width="100%">
<td width="10%" align="right"><font face="Arial"><font size="2"><%=rs12(0)%></font></font></td>
<td width="10%" align="right"><font face="Arial"><font size="2"><%=rs12(1)%></font></font></td>
<td width="70%" align="left"><font face="Arial"><font size="2"><%=rs12(2)%></font></font></td>
<td width="10%" align="right"><font face="Arial"><font size="2"><%=demand%></font></font></td>
</table>
 
 </body>

 <%
rs12.MoveNext  
Loop
rs12.close
set rs12 = nothing
set cnn1 = nothing
%>

<%
public Function CalcDemandPrice(bldgNum,LeaseUtilityId,BillYr,BillPeriod)

set rs11 = Server.CreateObject("ADODB.recordset")

    Dim dteUtilityFrom 
    Dim dteUtilityTo
    Dim dteTrans
    Dim dtePeak
    Dim varDemand
    Dim strRateTenant
    Dim strFind
    Dim strSumWinPrefix
    Dim dblGrossReceipt
    Dim dblUnitCostKW
    Dim dblDemand
    Dim dblCopyDemand
    Dim dblPrice
    Dim dblAddPrice
    Dim lngNumDays
    Dim fCoincident
    Dim intRType
    Dim fCalcBldg    
    Dim initFlag, loopFlag, blend
    Dim itrvl1, itrvl2, dmark
    Dim strDesc, strSuff
    Dim arr(3), i
    Dim lngLevel
    Dim tempPrice, pbuffer, flatsum
    fCalcBldg = False
    gdetailCostKW = ""
    
    

    If Left(BldgNum, 5) = "Calc:" Then
    ' Calculating Unit Cost for KW for Building
        fCalcBldg = True
        BldgNum = Mid(BldgNum, 6)
        strRateTenant = DLookup("[RateBldg]", "Buildings", BuildCriteria("[BldgNum]", dbText, BldgNum))
        fCoincident = False
    Else
        ' look for lease change, 'LC', first
        
Dim ii

cmd.ActiveConnection = cnn1
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "qprmLCLeaseUtlinfo"

        While Not cmd.Parameters.Count = 0
        cmd.Parameters.Delete 0
        Wend

Set p1 = cmd.CreateParameter("luid", adChar, adParamInput, 10, LeaseUtilityId)
cmd.Parameters.Append p1
Set p2 = cmd.CreateParameter("yr", adChar, adParamInput, 10, BillYr)
cmd.Parameters.Append p2
Set p3 = cmd.CreateParameter("per", adChar, adParamInput, 10, BillPeriod)
cmd.Parameters.Append p3

Set rs11 = cmd.Execute
        
    ' if none found, use old query
        If rs11.EOF Then
        
        While Not cmd.Parameters.Count = 0
        cmd.Parameters.Delete 0
        Wend
        
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "qprmLeaseUtilityInfo"
        Set p1 = cmd.CreateParameter("lID", adChar, adParamInput, 10, LeaseUtilityId)
        cmd.Parameters.Append p1
        Set rs11 = cmd.Execute
 
        If rs11.EOF Then
          
        End If
        End If
        If IsNull(rs11("rateTenant")) Then
        
          CalcDemandPrice = 0
          exit function
          
        End If
        strRateTenant = rs11("rateTenant")
        fCoincident = rs11("Coincident")
        rs11.Close
        
    End If
    
    Select Case strRateTenant
    Case "Avg"
        CalcDemandPrice = 0
        Exit Function
    Case "Avg Cost"
        ' Perform calculation using Unit Cost from Utility Bill for Building
    End Select

    ' Determine if Summer or Winter
    
        While Not cmd.Parameters.Count = 0
        cmd.Parameters.Delete 0
        Wend
        
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "qselutlinf"
        Set p1 = cmd.CreateParameter("bldg", adChar, adParamInput, 10, BldgNum)
        cmd.Parameters.Append p1
        Set p2 = cmd.CreateParameter("yr", adChar, adParamInput, 10, BillYr)
        cmd.Parameters.Append p2
        Set p3 = cmd.CreateParameter("bldg", adChar, adParamInput, 10, BillPeriod)
        cmd.Parameters.Append p3
        Set rs11 = cmd.Execute
        
 
 
        If rs11.EOF Then
        
            ' Values Not Entered
            CalcDemandPrice = 0
            exit function
        Else
            dblGrossReceipt = rs11.Fields("GrossReceipt")
            dblUnitCostKW = rs11.Fields("UnitCostKW")
            
            dteUtilityFrom = DateAdd("d", -1, rs11("DateStart"))            
            dteUtilityTo = rs11("DateEnd")
            dteTrans = "6/1/2001" 'Nz(rs11("DateTrans"), 1/1/1905)
           
            strSumWinPrefix = rs11("SumWin")
            lngNumDays = DateDiff("d", dteUtilityFrom, dteUtilityTo)
                       

                      
            If fCalcBldg Then

                dblDemand = rs11("TotalKW")
                dblCopyDemand = dblDemand
                Select Case strSumWinPrefix
                Case "SumWin", "WinSum"
                    ' Transition Month
                    strSumWinPrefix = Left(strSumWinPrefix, 3)
                Case Else
                End Select
            End If
        End If
        rs11.Close
    

 
    If Not fCalcBldg Then
    

  
        If fCoincident Then

        While Not cmd.Parameters.Count = 0
        cmd.Parameters.Delete 0
        Wend
        cmd.CommandType = adCmdStoredProc
        cmd.CommandText = "qselutlinf"
        Set p1 = cmd.CreateParameter("LID", adChar, adParamInput, 10, LeaseUtilityId)
        cmd.Parameters.Append p1
        Set p2 = cmd.CreateParameter("yr", adChar, adParamInput, 10, BillYr)
        cmd.Parameters.Append p2
        Set p3 = cmd.CreateParameter("bldg", adChar, adParamInput, 10, BillPeriod)
        cmd.Parameters.Append p3
        Set rs11 = cmd.Execute
        
     	If rs11.EOF Then
       varDemand = Null
       else
        
            varDemand = rs11("demand") 'DLookup("[Demand]", "CoincidentDemand", strFind)
            dtePeak = rs11("datePeak") 'DLookup("[DatePeak]", "CoincidentDemand", strFind)
       end if
       
        Else
        
        While Not cmd.Parameters.Count = 0
        cmd.Parameters.Delete 0
        Wend
        
        cmd.CommandType = 4
        cmd.CommandText = "qselTotalDemand"
        
        Set p1 = cmd.CreateParameter("preD", adchar, adParamInput, 10, dteUtilityFrom)
        cmd.Parameters.Append p1
        Set p2 = cmd.CreateParameter("endD", adchar, adParamInput, 10, dteUtilityTo)
        cmd.Parameters.Append p2
        Set p3 = cmd.CreateParameter("LID", adChar, adParamInput, 10, LeaseUtilityId)
        cmd.Parameters.Append p3
        Set p4 = cmd.CreateParameter("yr", adChar, adParamInput, 10, BillYr)
        cmd.Parameters.Append p4
        Set p5 = cmd.CreateParameter("per", adChar, adParamInput, 10, BillPeriod)
        cmd.Parameters.Append p5
        rs11.cursortype=3
        Set rs11 = cmd.Execute
        
        
        
      If rs11.EOF Then
       varDemand = Null
        Else
        
                varDemand = rs11("TotalDemand")
                dtePeak = rs11("PeakDate")
              
              
                Select Case strSumWinPrefix
                Case "SumWin", "WinSum"
                ' Transition Month, calculate Peak Demand using Rate in which Peak occured
                    If rs11("PeakDate") < dteTrans Then
                        strSumWinPrefix = Left(strSumWinPrefix, 3)
                    Else
                        strSumWinPrefix = Right(strSumWinPrefix, 3)
                    End If
                Case Else
                End Select
            End If
            rs11.Close
            
        End If
        
     if IsNull(varDemand) Then
            CalcDemandPrice = 0
            exit function
        Else
            dblDemand = varDemand 
            
        

 
    
        End If
        dblCopyDemand = dblDemand
       
        
    End If
 
    dblPrice = 0
   
strsql="SELECT DISTINCT " &_
"tblUtilityRates.*, tblUtility.UtilitySuffix AS utilitysuffix " &_
"FROM tblUtilityRates LEFT OUTER JOIN " &_
"tblUtility ON tblUtilityRates.Utility = tblUtility.UtilityDisplay " &_
"WHERE (tblUtilityRates.DateEnd >= '" & dteUtilityFrom & "') AND " &_
"(tblUtilityRates.RateTenant = '" & strRateTenant & "') AND " &_
"(tblUtilityRates.Utility = 'demand') AND " &_
"(tblUtilityRates.DateEffective <= '"& dteUtilityTo &"') order BY tblUtilityRates.[level], tblUtilityRates.dateeffective"                            


rs11.open strsql,cnn1,adopendynamic,adlockoptimistic,adcmdtext
 
 
 
    ' set flags that control blend
    initFlag = False
    blend = False
     
    Do Until rs11.EOF
        ' blend loop
        strDesc = rs11("RateDesc")        
        strSuff = rs11("utilitysuffix")        
        lngLevel = rs11(5)

        tempPrice = 0
        flatsum = 0
        dmark = rs11("DateEffective")
        If initFlag = False Then
           loopFlag = True
           i = 0
           tempPrice = 0
           
           While Not rs11.EOF And (lngLevel = rs11(5))
                
                If rs11("DateEffective") < dteUtilityFrom Then
                   itrvl1 = dteUtilityFrom
                Else
                   itrvl1 = rs11("DateEffective")
                End If
                
                If rs11("DateEnd") > dteUtilityTo Then
                   itrvl2 = dteUtilityTo
                Else
                   itrvl2 = rs11("DateEnd")
                End If
                
                If i = 0 Then
                  arr(i) = DateDiff("d", itrvl1, itrvl2)
                Else
                  arr(i) = DateDiff("d", itrvl1, itrvl2) + 1
                End If
                
                If lngLevel <> 0 Then
                  tempPrice = tempPrice + rs11(strSumWinPrefix & "PeakPrice") * arr(i)
                  i = i + 1
                Else
                  pbuffer = rs11(strSumWinPrefix & "PeakPrice")
                End If
                
             
                rs11.MoveNext
                

                
                If lngLevel = 0 Then
                  If rs11.EOF Then
                    tempPrice = tempPrice + (flatsum + pbuffer) * arr(i)
                    flatsum = 0
                    i = i + 1
                    
                    rs11.MovePrevious
                    
                                       
                    lngLevel = -1
                  ElseIf dmark <> rs11("DateEffective") Then
                    tempPrice = tempPrice + (flatsum + pbuffer) * arr(i)
                    flatsum = 0
                    i = i + 1
                    dmark = rs11("DateEffective")
                  Else
                    flatsum = flatsum + pbuffer
                    dmark = rs11("DateEffective")
                  End If
                End If
           Wend
           
           arr(i) = -1
           tempPrice = tempPrice / lngNumDays
           
           If i > 1 Then
              blend = True
           End If
           If lngLevel <> -1 Then                   
           
           rs11.MovePrevious
           Else
             lngLevel = 0
           End If
           initFlag = True
        ElseIf blend = True Then
          'loop thru vals
          i = 1
          tempPrice = rs11(strSumWinPrefix & "PeakPrice") * arr(0)
          While arr(i) <> -1
            rs11.MoveNext
            tempPrice = tempPrice + rs11(strSumWinPrefix & "PeakPrice") * arr(i)
            i = i + 1
          Wend
        tempPrice = tempPrice / lngNumDays
       

        Else
          tempPrice = rs11(strSumWinPrefix & "PeakPrice")
        End If
        ' end blend loop
        
        If lngLevel = 0 Then
            ' Price applied to Total Demand Used
            dblAddPrice = (tempPrice * dblDemand)
            'gdetailCostKW = gdetailCostKW & vbCrLf & dblDemand & " KW @ $" & Format(tempPrice, "0.00000") & " (" & strDesc & ")"
        Else
            ' Price applied to portion of Demand Used
            If dblCopyDemand >= lngLevel Then
                'gdetailCostKW = gdetailCostKW & vbCrLf & lngLevel & " KW @ $" & Format(tempPrice, "0.00000") & " (" & strDesc & ")"
                
                
                dblAddPrice = Formatnumber((tempPrice * lngLevel), 2)
                
               
                
                dblCopyDemand = dblCopyDemand - lngLevel
            Else
                'gdetailCostKW = gdetailCostKW & vbCrLf & dblCopyDemand & " KW @ $" & Format(tempPrice, "0.00000") & " (" & strDesc & ")"
                If lngLevel = 5 Then
                  If dblCopyDemand < 5 And UCase(strRateTenant) = "SC9R1" Then
                    dblAddPrice =(tempPrice * 5)
                  Else
                    dblAddPrice = (tempPrice * dblCopyDemand)
                  End If
                Else
                  dblAddPrice = (tempPrice * dblCopyDemand)
                End If
                dblCopyDemand = 0
            End If
        End If

        If rs11("Prorate") Then
            'gdetailCostKW = gdetailCostKW & " ProRated At " & lngNumDays & "/30 = $" & Format(dblAddPrice * Format((lngNumDays / 30), "0.00####"), "#,##0.00")
            
            dblPrice = dblPrice + (dblAddPrice * (lngNumDays / 30))
            

        Else
            'gdetailCostKW = gdetailCostKW & " = $" & dblAddPrice
            dblPrice = dblPrice + dblAddPrice
        End If
        rs11.MoveNext
      Loop
    
    
 rs11.Close
 

    If ucase(Left(strRateTenant, 3)) <> "AVG" Then
    
        'gdetailCostKW = gdetailCostKW & vbCrLf & "Gross Receipts @ " & _
        'Format(dblGrossReceipt, "0.00####") & " = " & Format(dblPrice * dblGrossReceipt, "$#,##0.00")
        dblPrice = dblPrice + (dblPrice * dblGrossReceipt)
        'gdetailCostKW = gdetailCostKW & vbCrLf & "TOTAL DEMAND = " & Format(dblPrice, "$#,##0.00")
    Else
    ' Only AvgCost Rate Class will come this far, not AVG
        'gdetailCostKW = dblDemand & " KW At An Avg Cost Of $" & dblUnitCostKW & " = $" & Format(dblDemand * dblUnitCostKW, "0.00")
        dblPrice = (dblDemand * dblUnitCostKW)
    End If

    CalcDemandPrice = Formatnumber(dblPrice, 2)

   
End Function

%>    


