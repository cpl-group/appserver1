<%@Language="VBScript"%>
<!-- #include file="./adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<%
    if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
      '         Response.Write "../index.asp"
        end if  
    

bldg = Request("bldg")
bldgstr = Chr(34) & bldg & Chr(34)
%>

<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">

<title>Genergy ERI Management</title>

<script language="JavaScript" type="text/javascript">
function addnew(){
  // Load Add new tenant form
  //document.frames.title.location.href="null.htm";
  parent.frames.info.location.href="tenantlist.asp?bldg=" + <%=bldgstr%>;
  location.href="ti_add.asp?bldg=" + <%=bldgstr%>;
}

</script>
<link rel="Stylesheet" href="styles.css" type="text/css">   
</head>
<body bgcolor="#FFFFFF">

<%	  if Session("eri") > 2 then %>
<table border=0 cellpadding="3" cellspacing="0" width="100%">
<tr bgcolor="#eeeeee">
  <td><b>Tenants</b></td>
  <td align="right"><input type="button" name="Submit2" value="Add New Tenant" onclick="addnew()"></td>
</tr>
</table>
<%	   end if  	%>
<table border=0 cellpadding="3" cellspacing="1" width="100%">
<tr bgcolor="#dddddd" style="font-weight:normal;"> 
  <td>Tenant Num.</td>
  <td>Tenant Name</td>
  <td>Sqft</td>
  <td>Monthly Charge</td>
  <td>Yearly Charge</td>
  <td>$ / Sqft</td>
  <td>Lease Exp. Date</td>
  <td>Move Out Date</td>
</tr>
<%


Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.open getConnect(0,0,"Engineering")


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
<tr onMouseOver="this.style.backgroundColor = 'lightgreen'" style="cursor:hand" onMouseOut="this.style.backgroundColor = 'white'" onClick="javascript:parent.info.document.location='info.asp?qcatnr=<%=rst1("Tenant_no")%>'" > 
  <td> <%=rst1("Tenant_no")%></td>
  <td><%=rst1("Tenantname")%></td>
  <% If IsNull(rst1("sqft")) then %>
  <td></td>
  <% else %>
  <td><%=Formatnumber(rst1("sqft"),0)%></td>
  <%end If%>
  <% If IsNull(rst1("ccm")) then %>
  <td></td>
  <% else %>
  <td><%=FormatCurrency(rst1("ccm"),2)%></td>
  <%end If%>
  <% If IsNull(rst1("ccy")) then %>
  <td></td>
  <% else %>
  <td><%=FormatCurrency(rst1("ccy"),2)%></td>
  <%end If%>
  <% If IsNull(rst1("cost_sqft")) then %>
  <td></td>
  <% else %>
  <td><%=formatcurrency(rst1("cost_sqft"),2)%></td>
  <%end If%>
  <td><%if rst1("lease_exp_date")<>"1/1/1900"  then response.write rst1("lease_exp_date")%></td>
  <td><%if rst1("Move_out_date")<>"1/1/1900" then response.write rst1("Move_out_date")%></td>
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

</body>
</html>

