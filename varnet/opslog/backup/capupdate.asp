<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
bldgname=Request.Form("bldgname")
floor=Request.Form("floor")
bldgnum=Request.Form("bldgnum")
riser=Request.Form("riser")
size=Request.Form("size")
metal=Request.Form("metal")
volts=Request.Form("volts")
insulation=Request.Form("insulation")
sets=Trim(Request.Form("sets"))
sframe=Request.Form("sframe")
sfuse=Request.Form("sfuse")
wc=Request.Form("wc")
choice=Request.Form("choice")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=Capacity_db;"
if choice="Update" then
strsql = "Update tblriser Set size='" & size & "', metal='" & metal & "', insulation='" & insulation & "', sets='" & sets & "', volts='" &volts & "', sw_frame='" &sframe & "', sw_fuse='" & sfuse & "' where (riser_name='"& riser &"' and bldgnum='"& bldgnum &"')"
end if


response.write strsql
'cnn1.execute strsql
'cnn1.execute strsql2
set cnn1=nothing
tmpMoveFrame =  "document.location = " & Chr(34) & _
				"capdetail.asp?bldgname="& bldgname &_     
				"&bldgnum="&bldgnum&"&floor="&floor& " &riser="&riser&chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
'Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>