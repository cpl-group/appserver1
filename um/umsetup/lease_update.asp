<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
			Response.Redirect "http://www.genergyonline.com"
		end if		
%>
<%
Bill= request("b")
bldg= request("bldg")
ten= request("ten")

billingname = Request("billingname")
flr = Request("flr")
sqft = Request("sqft")
taxexempt= Request("taxexempt")
leaseexpired= Request("leaseexpired")

if taxexempt = "on" or taxexempt = "off" then
taxexempt=1
else
taxexempt=0
end if

if leaseexpired= "on" or leaseexpired= "off" then
leaseexpired=1
else
leaseexpired=0
end if



Set cnn1 = Server.CreateObject("ADODB.Connection")

cnn1.Open "driver={SQL Server};server=10.0.7.8;uid=genergy1;pwd=g1appg1;database=genergy1;"

strsql = "UPDATE tblleases SET billingname='" & billingname & "', leaseexpired='" & leaseexpired & "', flr='" & flr & "', sqft=" & sqft & ", taxexempt='" & taxexempt& "' where billingid = " & bill


cnn1.execute strsql

'response.write strsql
set cnn1=nothing

response.redirect "leases_info.asp?bldg=" & bldg &"&ten=" & ten
%>
























