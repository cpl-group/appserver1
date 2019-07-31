<%@Language="VBScript"%>
<!-- #include file="../opslog/adovbs.inc" -->
<%
sitekey=Request.Form("sitekey")
choice=Request.Form("choice")
userid=Request.Form("username")
bldg_name=Trim(Request.Form("bldg_name"))
if bldg_name="" then
    response.redirect "usrsite.asp?username="&username
end if
region=Request.Form("region")
regioncount=Request.Form("regioncount")
bldgid=Request.Form("bldgid")
contact_name=Request.Form("contact_name")
bldgid_link=Request.Form("bldgid_link")
pgi_link=Request.Form("pgi_link")
pgi_target=Request.Form("pgi_target")
eri_link=Request.Form("eri_link")
eri_target=Request.Form("eri_target")
meter_link=Request.Form("meter_link")
meter_target=Request.Form("meter_target")
ovthvac_link=Request.Form("ovthvac_link")
ovthvac_target=Request.Form("ovthvac_target")
pow_ava_link=Request.Form("pow_ava_link")
pow_ava_target=Request.Form("pow_ava_target")
pa_chart_link=Request.Form("pa_chart_link")
pa_chart_target=Request.Form("pa_chart_target")
rev_prof_link=Request.Form("rev_prof_link")
rev_prof_target=Request.Form("rev_prof_target")
iri_link=Request.Form("iri_link")
iri_target=Request.Form("iri_target")
mep_link=Request.Form("mep_link")
mep_target=Request.Form("mep_target")
pq_link=Request.Form("pq_link")
pq_target=Request.Form("pq_target")
plp_link=Request.Form("plp_link")
plp_target=Request.Form("plp_target")
msi_link=Request.Form("msi_link")
msi_target=Request.Form("msi_target")
lmp_link=Request.Form("lmp_link")
lmp_target=Request.Form("lmp_target")
ca_link=Request.Form("ca_link")
ca_target=Request.Form("ca_target")


pgi=Request.Form("pgi")
response.write(pgi)
eri=Request.Form("eri")
meter=Request.Form("meter")
ovthvac=Request.Form("ovthvac")
pow_ava=Request.Form("pow_ava")
pa_chart=Request.Form("pa_chart")
rev_prof=Request.Form("rev_prof")
iri=Request.Form("iri")
mep=Request.Form("mep")
pq=Request.Form("pq")
plp=Request.Form("plp")
msi=Request.Form("msi")
lmp=Request.Form("lmp")
ca=Request.Form("ca")
if meter <> 1 then
    meter=0
end if
if ovthvac <> 1 then
    ovthvac=0
end if
if pow_ava <> 1 then
    pow_ava=0
end if
if pa_chart <> 1 then
	pa_chart=0
end if
if rev_prof <> 1 then
    rev_prof=0
end if
if iri <> 1 then
    iri=0
end if
if mep <> 1 then
    mep=0
end if
if pq <> 1 then
    pq=0
end if
if msi <> 1 then
    msi=0
end if
if lmp <> 1 then
    lmp=0
end if
if plp <> 1 then
    plp=0
end if
if ca <> 1 then
    ca=0
end if
if pgi <> 1 then
    pgi=0
end if
if eri <> 1 then
    eri=0
end if
'response.write(ca_link)
'response.write(iri_link)
'response.write(userid)
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=security;"
if(choice = "Add") then
strsql = "Insert into clientsites (userid, region, regioncount, bldgid, bldg_name, contact_name, bldgid_link, pgi_link, pgi_target, eri_link, eri_target, meter_link, meter_target, ovthvac_link, ovthvac_target, pow_ava_link, pow_ava_target, pa_chart_link, pa_chart_target, rev_prof_link, rev_prof_target, iri_link, iri_target, mep_link, mep_target, pq_link, pq_target, plp_link, plp_target,  msi_link, msi_target, lmp_link, lmp_target, ca_link, ca_target, pgi, eri, meter, ovthvac, pow_ava, pa_chart, rev_prof, iri, mep, pq, plp,msi, lmp, ca) "_
	& "values ("_
	& "'" & userid & "', '" & region & "', '" & regioncount & "',"_
	& "'" & bldgid & "', '" & bldg_name & "', '" & contact_name & "', "_
	& "'" & bldgid_link & "', '" & pgi_link & "', '" & pgi_target & "',"_
	& "'" & eri_link & "', '" & eri_target & "', '" & meter_link & "', "_
	& "'" & meter_target & "', '" & ovthvac_link & "', '" & ovthvac_target & "', "_
	& "'" & pow_ava_link & "', '" & pow_ava_target & "', '" & pa_chart_link & "', "_
	& "'" & pa_chart_target & "', '" & rev_prof_link & "', '" & rev_prof_target & "', "_
	& "'" & iri_link & "', '" & iri_target & "', '" & mep_link & "', "_
	& "'" & mep_target & "', '" & pq_link & "', '" & pq_target & "', "_
	& "'" & plp_link & "', '" & plp_target & "', '" & msi_link & "', "_
	& "'" & msi_target & "', '" & lmp_link & "', '" & lmp_target & "', "_
	& "'" & ca_link & "', '" & ca_target & "', '" & pgi & "', "_
	& "'" & eri & "', '" & meter & "', '" & ovthvac & "', "_
	& "'" & pow_ava & "', '" & pa_chart & "', '" & rev_prof & "', "_
	& "'" & iri & "', '" & mep & "', '" & pq & "', "_
	& "'" & plp & "', '" & msi & "', '" & lmp & "', "_
	& "'" & ca & "')"

'Response.Write strsql	

end if
if choice="Save" then
strsql = "Update clientsites Set userid='" & userid & "', region='" & region & "', regioncount='" &regioncount & "', bldgid='" & bldgid & "', bldg_name='" & bldg_name & "', contact_name='" & contact_name & "', bldgid_link='" & bldgid_link & "', pgi_link='" &pgi_link & "', pgi_target='" & pgi_target & "', eri_link='" & eri_link & "', eri_target='" & eri_target & "', meter_link='" & meter_link & "', meter_target='" &meter_target & "', ovthvac_link='" & ovthvac_link & "', ovthvac_target='" & ovthvac_target & "', pow_ava_link='" & pow_ava_link & "', pow_ava_target='" & pow_ava_target & "', pa_chart_link='" &pa_chart_link & "', pa_chart_target='" & pa_chart_target & "', rev_prof_link='" & rev_prof_link & "', rev_prof_target='" & rev_prof_target & "', iri_link='" & iri_link & "',  iri_target='" & iri_target & "', mep_link='" &mep_link & "', mep_target='" & mep_target & "', pq_link='" & pq_link & "', pq_target='" & pq_target & "', plp_link='" & plp_link & "', plp_target='" &plp_target & "', msi_link='" & msi_link & "', msi_target='" & msi_target & "', lmp_link='" & lmp_link & "', lmp_target='" &lmp_target & "', ca_link='" & ca_link & "', ca_target='" & ca_target & "', pgi='" & pgi & "', eri='" & eri & "', meter='" &meter & "', ovthvac='" & ovthvac & "', pow_ava='" & pow_ava & "', pa_chart='" & pa_chart & "', rev_prof='" & rev_prof & "', iri='" &iri & "', mep='" & mep & "', pq='" & pq & "', plp='" & plp & "', msi='" & msi & "', lmp='" &lmp & "', ca='" & ca & "' where (clientsitekey='"& sitekey &"')"


end if
if choice="Delete" then
strsql= "Delete from clientsites where clientsitekey = '"& sitekey &"'"
end if
'Response.Write strsql
cnn1.execute strsql
set cnn1=nothing
tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "usrsite.asp?username="& userid & chr(34) & vbCrLf 

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>