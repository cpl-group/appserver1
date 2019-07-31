<%

Function billByPeriod(bldg As String, yr As String, per As Integer)

Dim rst1 As New ADODB.Recordset
Dim rst2 As New ADODB.Recordset
Dim rst3 As New ADODB.Recordset
Dim rst4 As New ADODB.Recordset
Dim rs As New ADODB.Recordset

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")
Set rst3 = Server.CreateObject("ADODB.recordset")
Set rst4 = Server.CreateObject("ADODB.recordset")
Set rs = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"


Dim fieldname, update_str, update_cmd As String

Dim energy, Demand, Subtotal, tax, total As Double
Dim metercount, ypId, leaseid As Integer


cmd.ActiveConnection = cnn1
cmd.CommandType = adCmdStoredProc
cmd.CommandText = "sp_billsByPeriod"

        While Not cmd.Parameters.Count = 0
        cmd.Parameters.Delete 0
        Wend


Set p1 = cmd.CreateParameter("bldg", adChar, adParamInput, 10, bldg)
cmd.Parameters.Append p1
Set p2 = cmd.CreateParameter("yr", adChar, adParamInput, 10, yr)
cmd.Parameters.Append p2
Set p3 = cmd.CreateParameter("per", adChar, adParamInput, 10, per)
cmd.Parameters.Append p3

Set rst1 = cmd.Execute


rst3.Open "tblBillbyPeriod", cnn1, adOpenDynamic, adLockOptimistic
rst4.Open "tblmetersbyPeriod", cnn1, adOpenDynamic, adLockOptimistic

While Not rst1.EOF

energy = 0
Demand = 0
Subtotal = 0
tax = 0
total = 0
metercount = 0
    ypId = rst1("ypId")
    leaseid = rst1("LeaseUtilityId")

    energy = CalcEnergyPrice(bldg, leaseid, yr, per)
    Demand = calcdemandprice(bldg, leaseid, yr, per)
    
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT COUNT (MeterId)as a FROM Meters WHERE (DateOffLine > DateLastRead) AND (LeaseUtilityId = " & leaseid & ")"
    Set rst2 = cmd.Execute
    
    metercount = rst2(0)
    
    Subtotal = (energy + Demand) + ((energy + Demand) * rst1("adminfee")) + (rst1("addonfee") * metercount)
  
    If Not rst1("taxexempt") Then
    tax = Subtotal * rst1("salestax")
    Else
    tax = 0
    End If
    
    
    If IsNull(tax) Then
    tax = 0
    End If
    
    total = Subtotal + tax
    
 
        rst3.AddNew
        
            For i = 0 To rst1.Fields.Count - 1
            fieldname = rst1.Fields.Item(i).Name
            rst3(fieldname) = rst1(i)
            Next i
            
        rst3("subtotal") = Subtotal
        rst3("tax") = tax
        rst3("totalamt") = total
        rst3("energy") = energy
        rst3("demand") = Demand
        rst3.Update

        
        ' insert into tblmetersbyperiod info about meters
        
        Set rs = rst("sp_tbl_meter", leaseid, rst1("ypId"))
            If Not rs.EOF Then
                
                While Not rs.EOF
                    rst4.AddNew
                    For i = 0 To rs.Fields.Count - 1
                    fieldname = rs.Fields.Item(i).Name
                    rst4(fieldname) = rs(i)
                    Next i
                    rs.MoveNext
                Wend
                rst4.Update
            End If
            
        ' update consumption data
        
            Set rs = rst("sp_tbl_consumption", leaseid, rst1("ypId"))
            
             While Not rs.EOF
                update_str = ""
                          
                    For i = 0 To rs.Fields.Count - 3
                    fieldname = rs.Fields.Item(i).Name
                    update_str = update_str & fieldname & " =" & rs(i) & ", "
                    
                    Next i
                    fieldname = rs.Fields.Item(i).Name
                    update_str = update_str & fieldname & " =" & rs(i)
                    
                    cnn1.Execute "UPDATE tblmetersbyperiod SET " & update_str & " where meterid =" & rs("meterid") & " and ypid =" & ypId

                    rs.MoveNext
            Wend
            
        ' insert peak demand data
            
        Set rs = rst("sp_tbl_peakdemand", leaseid, rst1("ypId"))
            
        While Not rs.EOF
            cnn1.Execute "UPDATE tblmetersbyperiod SET DatePeak_P ='" & rs("datepeak") & "', Demand_P =" & rs("demand") & " where meterid =" & rs("meterid") & " and ypid =" & ypId
            rs.MoveNext
        Wend

        ' insert coincident demand data
            
        Set rs = rst("sp_tbl_coincident", leaseid, rst1("ypId"))
            
        While Not rs.EOF
        
            cnn1.Execute "UPDATE tblmetersbyperiod SET DatePeak_C ='" & rs("datepeak") & "', Demand_C =" & rs("demand") & " where leaseutilityid =" & leaseid & " and ypid =" & ypId
            rs.MoveNext
            
        Wend
        
    
  
            
            
            
            
        
    rst1.MoveNext
    
   
    
Wend

MsgBox (bldg & "  " & yr & "/" & per & "     Done!")

pre_exit_bill:

Set cnn1 = Nothing
Debug.Print Time()
GoTo exit_bill


error_bill:
 Select Case Err.Number
 
    Case -2147217873
    MsgBox ("You are Trying to generate existing bills ! Check your parameter and try again ")
    GoTo pre_exit_bill
    
    
    Case Else
  
    Resume Next
    
  
    End Select
exit_bill:

End Function

%>