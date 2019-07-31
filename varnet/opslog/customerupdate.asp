<body bgcolor="#FFFFFF">
<div align="center">
  <table width="100%" border="0" bgcolor="#3399CC">
    <tr>
      <td>
        <div align="center"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><b>CONTACT HAS BEEN UPDATED SUCCESSFULLY</b></font></div>
      </td>
    </tr>
  </table>
  <p> 
    <input type="button" name="Button" value="CLOSE THIS WINDOW" onclick="javascript:window.close()">
  </p>
</div>
<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
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
cnn1.Open application("cnnstr_main")

strsql="update contacts set CompanyName='" & company & "',first_name='" & first & "',last_name='" &last & "',address='" & addr & "',city='" & city & "',state='" & state & "',zip='" & zip & "',country='" & country & "',title='" & t & "',phone='" & phone & "',fax='" & fax& " ',email='" & email & "' where customerid=" & cid & ""
'response.write strsql
'response.end



cnn1.execute strsql
set cnn1=nothing
tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "contactview.asp?cid="& cid & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>