<body bgcolor="#FFFFFF">
<div align="center">
  <table width="100%" border="0" bgcolor="#3399CC">
    <tr>
      <td>
        <div align="center"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><b>CUSTOMER 
          HAS BEEN UPDATED SUCCESSFULLY</b></font></div>
      </td>
    </tr>
  </table>
  <p> 
    <input type="button" name="Button" value="CLOSE THIS WINDOW" onclick="javascript:window.close()">
  </p>
</div>
<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
choice=Request.Form("update")
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
email=Request.Form("email")
cid=Request.Form("cid")
'response.write(manager)
'response.write(billdate)

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"intranet")

strsql="update customers set CompanyName='" & company & "',contactfirstname='" & first & "',contactlastname='" &last & "',billingaddress='" & addr & "',city='" & city & "',stateorprovince='" & state & "',postalcode='" & zip & "',country='" & country & "',contacttitle='" & t & "',phonenumber='" & phone & "',faxnumber='" & fax& " ',email='" & email & "' where customerid=" & cid & ""
'response.write strsql
'response.end



cnn1.execute strsql
set cnn1=nothing


Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>