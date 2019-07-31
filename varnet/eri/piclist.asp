<%@Language="VBScript"%>
<!-- #include file="./adovbs.inc" -->
<html>
<head>
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'					Response.Write "../index.asp"
   	    end if	
		
%>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Genergy ERI Management</title>
</head>
<body bgcolor="#FFFFFF">
<div>
  <table border="1" width="100%" height="100%">
    <tr> 
      <td width="10%" height="20" align="center" bgcolor="#66CCFF"> 
        <p align="center"><font face="Arial, Helvetica, sans-serif"><font color="#000000">Tenant 
          #</font></font>
      </td>
      <td width="30%" height="20" align="center" bgcolor="#66CCFF"> 
        <p align="center"><font face="Arial, Helvetica, sans-serif"><font color="#000000">Tenant 
          Name</font></font>
      </td>
      <td width="10%" height="20" align="center" bgcolor="#66CCFF"><font color="#000000" face="Arial, Helvetica, sans-serif">Sqft</font></td>
      <td width="10%" height="20" align="center" bgcolor="#66CCFF"><font color="#000000" face="Arial, Helvetica, sans-serif">Monthly 
        Charge</font></td>
      <td width="10%" height="20" align="center" bgcolor="#66CCFF"><font color="#000000" face="Arial, Helvetica, sans-serif">Yearly 
        Charge</font></td>
      <td width="10%" height="20" align="center" bgcolor="#66CCFF"> <font color="#000000" face="Arial, Helvetica, sans-serif"> 
        $ / Sqft</font></td>
      <td width="10%" height="20" align="center" bgcolor="#66CCFF"><font color="#000000" face="Arial, Helvetica, sans-serif">Lease 
        Exp. Date</font></td>
      <td width="10%" height="20" align="center" bgcolor="#66CCFF"><font color="#000000" face="Arial, Helvetica, sans-serif">Move 
        Out Date</font></td>
    </tr>
    <%
bldg = Request("bldg")


Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"


Set rst1 = Server.CreateObject("ADODB.Recordset")

sql = "SELECT * FROM tenant_info WHERE (bldg_no='" & bldg & "') order by tenantname"

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly

' Write a browser-side script to update another frame (named
' detail) within the same frameset that displays this page.
Response.Write "<script>" & vbCrLf




If rst1.EOF Then
  Response.Write "parent.frames.info.location = " & _
                 Chr(34) & "null.htm" & Chr(34) & vbCrLf
Else
  Response.Write "parent.frames.info.location = " & _
                  Chr(34) & _
                  "info.asp?qcatnr=" & rst1("Tenant_no") & _
                  Chr(34) & vbCrLf
End If

Response.Write "</script>" & vbCrLf

If not rst1.EOF then 
rst1.movefirst
end if

Do While Not rst1.EOF
%>
    <tr> 
      <td width="10%" height="20" align="center"> 
        <p align="center"><%  If Session("eri") > 2 then  %> <a href= "ti_edit.asp?tenant_no=<% =rst1("Tenant_no") %>" ><img src="../images/Leaf.gif" width="16" height="16" border="0"></a>  <% End if %>
          <%
		 				Response.write "<a href=" & chr(34) & "info.asp?" & "qcatnr=" & rst1("Tenant_no") & chr(34) & " target=info>" & rst1("Tenant_no")
		 	%>
      </td>
      <td width="30%" height="20" align="center"> 
        <p align="left"><%=rst1("Tenantname")%>
      </td>
      <% If IsNull(rst1("sqft")) then %>
      <td width="10%" height="20" align="center"></td>
      <% else %>
      <td width="10%" height="20" align="center"> 
        <p align="right"><%=Formatnumber(rst1("sqft"),0)%>
      </td>
      <%end If%>
      <% If IsNull(rst1("ccm")) then %>
      <td width="10%" height="20" align="center"></td>
      <% else %>
      <td width="10%" height="20" align="center"> 
        <p align="right"><%=FormatCurrency(rst1("ccm"),2)%>
      </td>
      <%end If%>
      <% If IsNull(rst1("ccy")) then %>
      <td width="10%" height="20" align="center"></td>
      <% else %>
      <td width="10%" height="20" align="center"> 
        <p align="right"><%=FormatCurrency(rst1("ccy"),2)%>
      </td>
      <%end If%>
      <% If IsNull(rst1("cost_sqft")) then %>
      <td width="10%" height="20" align="center"></td>
      <% else %>
      <td width="10%" height="20" align="center"> 
        <p align="right"><%=formatcurrency(rst1("cost_sqft"),2)%>
      </td>
      <%end If%>
     <td width="10%" height="20" align="center"><font face="Arial, Helvetica, sans-serif" size="2"><%if rst1("lease_exp_date")<>"1/1/1900"  then response.write rst1("lease_exp_date")%></font></td>
    <td width="10%" height="20" align="center"> 
      <p align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%if rst1("Move_out_date")<>"1/1/1900" then response.write rst1("Move_out_date")%> 
        </font>
    </td>
    </tr>
    <%
  rst1.MoveNext  
Loop

'Close and destroy the recordset and connection objects.
rst1.Close
Set rst1 = Nothing
cnn1.Close
Set cnn1 = Nothing
%>
  </table>
</div>
</body>
</html>

