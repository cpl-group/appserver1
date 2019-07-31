<%@Language="VBScript"%>
<!-- #include file="../opslog/adovbs.inc" -->
<%
job=Request.Form("job")
entry_type=Request.Form("entry type")
companyname=Request.Form("companyname")
contactname=Trim(Request.Form("contactname"))
reqname=Request.Form("reqname")
reqphone=Request.Form("reqphone")
refby=Request.Form("refby")
custometphone=Request.Form("customerphone")
customerfax=Request.Form("customerfax")
floorroom=Request.Form("floorroom")
recdate=Request.Form("recdate")
reqtargetdate=Request.Form("reqtargetdate")
description=Request.Form("description")
enteredby=Request.Form("EnteredBy")
manager=Request.Form("manager")
status=Request.Form("status")
percentcomp=Request.Form("percentcomp")
billdate=Request.Form("billdate")
comments2=Request.Form("comments2")

response.write(job)
response.write(status)
response.write(comments2)
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=security;"


if choice="Update" then
'strsql = "Update [Job Log] Set userid='" & userid & "', region='" & region & "', regioncount='" &regioncount & "', bldgid='" & bldgid & "', bldg_name='" & bldg_name & "', contact_name='" & contact_name & "', bldgid_link='" & bldg_link & "', pgi_link='" &pgi_link & "', pgi_target='" & pgi_target & "', eri_link='" & eri_link & "', eri_target='" & eri_target & "', meter_link='" & meter_link & "', meter_target='" &meter_target & "', ovthvac_link='" & ovthvac_link & "', ovthvac_target='" & ovthvac_target & "', pow_ava_link='" & pow_ava_link & "', pow_ava_target='" & pow_ava_target & "', pa_chart_link='" &pa_chart_link & "', pa_chart_target='" & pa_chart_target & "', rev_prof_link='" & rev_prof_link & "', rev_prof_target='" & rev_prof_target & "', iri_link='" & iri_link & "',  iri_target='" & iri_target & "', mep_link='" &mep_link & "', mep_target='" & mep_target & "', pq_link='" & pq_link & "', pq_target='" & pq_target & "', plp_link='" & plp_link & "', plp_target='" &plp_target & "', msi_link='" & msi_link & "', msi_target='" & msi_target & "', lmp_link='" & lmp_link & "', lmp_target='" &lmp_target & "', ca_link='" & ca_link & "', ca_target='" & ca_target & "', pgi='" & pgi & "', eri='" & eri & "', meter='" &meter & "', ovthvac='" & ovthvac & "', pow_ava='" & pow_ava & "', pa_chart='" & pa_chart & "', rev_prof='" & rev_prof & "', iri='" &iri & "', mep='" & mep & "', pq='" & pq & "', plp='" & plp & "', msi='" & msi & "', lmp='" &lmp & "', ca='" & ca & "' where (clientsitekey='"& sitekey &"')"


end if

'response.write choice
'Response.Write strsql
'cnn1.execute strsql
set cnn1=nothing
tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "usrsite.asp?username="& userid & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
'Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>