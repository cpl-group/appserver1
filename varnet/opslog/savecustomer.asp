<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
choice=Request.Form("save")
company=Request.Form("CompanyName")
first=Request.Form("first")
last=Request.Form("last")
t=Trim(Request.Form("title"))
addr=Request.Form("addr")
city=Request.Form("city")
state=Request.Form("state")
phone=Request.Form("phone")
fax=Request.Form("fax")
zip=Request.Form("zip")
country=Request.Form("country")
cid=Request.Form("cid")
mkid=Request.Form("mkid")
'response.write(manager)
'response.write(billdate)

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
cnn1.Open application("cnnstr_main")

cnum=0
rst1.open("SELECT max(ID) as id FROM customers"), cnn1
if not(rst1.eof) then cnum=rst1("id")
rst1.close

rst1.open("SELECT phoneNumber, PostalCode, customerid, ContactLastName, ContactFirstName ,companyname FROM customers WHERE phonenumber like '%"& phone &"%' and PostalCode like '%"& zip &"%'"), cnn1
if not(rst1.eof) then
'response.write "SELECT phoneNumber, PostalCode, id, first_name, last_name FROM customers WHERE phonenumber like '%"& phone &"%' and 'PostalCode like '%"& zip &"%'"
%>
<HTML>
<body bgcolor="#FFFFFF">
<font face="Arial, Helvetica, sans-serif">
<div align="center">
  <table width="100%" border="0" bgcolor="#3399CC">
    <tr>
      <td>
        <div align="center"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><b>THIS CUSTOMER ALREADY EXISTS</b></font></div>
      </td>
    </tr>
  </table>
  <p>
<%
do until rst1.eof
    response.write "<a href=""javascript:window.opener.parent.document.location.href='newrfp.asp?mkid="& mkid &"&cnum=" &rst1("customerid") &"';window.close()"">"& rst1("ContactLastName") &" "& rst1("ContactFirstName") &" "& rst1("companyname") &"</a><BR>"
    rst1.movenext
loop
%>
    <input type="button" name="Button" value="CLOSE THIS WINDOW" onclick="javascript:window.opener.parent.document.location.href='newrfp.asp?mkid=<%=mkid%>&cnum=<%=cnum%>';window.close();">
  </p>
</div>
    </body></html>
    <%response.end
end if
strsql = "insert customers (customerid,CompanyName,contactfirstname,contactlastname,billingaddress,city,stateorprovince,postalcode,country,contacttitle,phonenumber,faxnumber)values (" & cid & ",'" & company & "', '" & first & "', '" &last & "', '" & addr & "', '" & city & "', '" & state & "', '" & zip & "','" &country & "','" & t & "', '" & phone & "', '" & fax & "')"
cnum=cint(cnum)+1

'response.write strsql
cnn1.execute strsql

if trim(mkid)<>"" then
    tmpMoveFrame = "window.opener.parent.document.location.href='newrfp.asp?mkid="& mkid &"&cnum=" &cid &"';window.close();"
else
    tmpMoveFrame =  "document.location = " & Chr(34) & "newcustomer.asp?job="& job & chr(34) & vbCrLf 
end if

cnn1.close

set cnn1=nothing

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>