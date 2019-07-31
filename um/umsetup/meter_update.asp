<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
			Response.Redirect "http://www.genergyonline.com"
		end if		
%>
<%
meterid= request("M")
lui = request("L")

meternum = Request("meternum")
datestart = Request("datestart")
dateoffline = Request("dateoffline")
datelastread = Request("datelastread")
timelastread = Request("timelastread")
multiplier = Request("multiplier")
manualmultiplier = Request("manualmultiplier")
demandmultiplier = Request("demandmultiplier")
location = Request("location")
riser = Request("riser")
online_A = Request("online")
metercomments = Request("metercomments")



if online_A = "on" or online_A = "off" then
online=1
else
online=0
end if




Set cnn1 = Server.CreateObject("ADODB.Connection")

cnn1.Open "driver={SQL Server};server=10.0.7.8;uid=genergy1;pwd=g1appg1;database=genergy1;"

strsql = "UPDATE meters SET meternum='" & meternum & "', online='" & online & "', datestart='" & datestart & "', dateoffline='" & dateoffline & "', datelastread='" & datelastread & "', multiplier=" & multiplier & ", manualmultiplier=" & manualmultiplier & ", demandmultiplier=" & demandmultiplier & ", location='" & location & "', riser='" & riser & "', metercomments='" & metercomments & "' where meterid = " & meterid
cnn1.execute strsql

'response.write online_A & "      " & online & "      " & strsql
set cnn1=nothing

response.redirect "meter_info.asp?lui=" & lui
%>
























