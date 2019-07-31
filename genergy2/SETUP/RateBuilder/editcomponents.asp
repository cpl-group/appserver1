<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/nonsecure.inc"-->
<%
'if not(allowGroups("Genergy Users")) then
'%><%
'end if

	dim cnn1, rst1, rst2, strsql, insertSql
	set cnn1 = server.createobject("ADODB.connection")
	set rst1 = server.createobject("ADODB.recordset")
	cnn1.open getConnect(0,0,"dbCore")
%>

<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:2359ice:2359ice"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 15">
<link id=Main-File rel=Main-File href="../Rate%20Builder.htm">
<link rel=File-List href=filelist.xml>
<link rel=Stylesheet href=stylesheet.css>
<style>
<!--table
	{mso-displayed-decimal-separator:"\.";
	mso-displayed-thousand-separator:"\,";}
@page
	{margin:.75in .7in .75in .7in;
	mso-header-margin:.3in;
	mso-footer-margin:.3in;}
-->
</style>

</head>

<%
	dim rbid,rbcid,rbrid,rateperiod,createdBy,createdOn,modifiedBy,modifiedOn, exact

	dim e_edc_r,e_edc_sc9_1,e_edc_sc9_2,e_edc_sc9_3,e_edc_sc12_2,e_edc_sc1_f250_1,e_edc_sc1_o250_1,e_tra_r,e_tra_sc9_1,e_tra_sc9_2,e_tra_sc9_3,e_tra_sc12_2,e_tra_sc1_f250_1,e_tra_sc1_o250_1,e_mac_r,e_mac_1,e_sbc_r,e_sbc_1,e_rpsp_r,e_rpsp_1,e_psls_r,e_psls_sc9_1,e_psls_sc12_2,e_psls_sc1_1,e_rdm_r,e_rdm_sc9_1,e_rdm_sc12_2,e_rdm_sc1_1,e_drs_r,e_drs_sc9_1,e_drs_sc12_2,e_drs_sc1_1,e_mfc_r,e_mfc_sc912_1,e_mfc_sc1_1
	dim d_mc_r,d_mc_1,d_o5_r,d_o5_1,d_mf86_r,d_mf86_sc9_2,d_mf86_sc9_3,d_mf86_sc12_2,d_mf810_r,d_mf810_sc9_2, d_mf810_sc9_3,d_mf810_sc12_2,d_all_r,d_all_sc9_2,d_all_sc9_3,d_all_sc12_2,d_tra_mc_r,d_tra_mc_1,d_tra_o5_r,d_tra_o5_1,d_tra_mf86_r,d_tra_mf86_sc9_2,d_tra_mf86_sc9_3
	dim d_tra_mf86_sc12_2,d_tra_mf810_r,d_tra_mf810_sc9_2,d_tra_mf810_sc9_3,d_tra_mf810_sc12_2,d_tra_all_r,d_tra_all_sc9_2,d_tra_all_sc9_3,d_tra_all_sc12_2,d_msccap_r,d_msccap_sc9r1_1,d_msccap_sc9r2_1,d_msccap_m_1,d_rpd_r,d_rpd_1
	dim s_cms_r,s_cms_sc9_m_1,s_cms_sc9_1,s_cms_sc9_2,s_cms_sc1_1,s_cms_sc1_low_1,s_bppc_r,s_bppc_el_1,s_bppc_elgs_1, e_cesds_1, e_cesss_1, e_cesds_r, e_cesss_r, s_cms_sc9_3, s_cms_sc9_m_3, s_cms_sc12_2, d_dlms_r, d_dlms_sc9_1, d_dlms_sc9_2_3, d_dlms_sc12_2
	dim e_sosc_sc9_1_3, d_sosc_sc9_1_3, e_sosc_sc9_2, d_sosc_sc9_2, e_sosc_sc12_2, d_sosc_sc12_2
	
	dim today, year, month, day
	dim eyear, emonth
	
	year = datepart("yyyy", now)
	month = datepart("m", now)
	day = datepart("d", now)
	if day > 20 then month = month + 1
	if month = 13 then 
		month = 1
		year = year + 1
	end if
	today = dateserial(year,month,1)
	
	eyear = request.form("eyear")
	emonth = request.form("emonth")
	
	if eyear <> "" and emonth <> "" then
		today = dateserial(eyear, emonth, 1)
		year = datepart("yyyy", today)
		month = datepart("m", today)
	end if

%>

<%
	rbid = trim(secureRequest("rbid"))
	
	if rbid <>"" then
		strsql = "select * from ratebuilder where rbid =" & rbid
	else
		strsql = "select top 1 *, case when rateperiod='"& today &"' then 1 else 0 end as exacting from ratebuilder where rateperiod <= '" &today& "' order by rateperiod desc, modifiedOn desc"
	end if

	rst1.Open strsql, cnn1
	if not rst1.EOF then
		rbid = rst1("rbid")
		rbcid = rst1("rbcid")
		rateperiod = rst1("rateperiod")
		createdBy = rst1("createdBy")
		createdOn = rst1("createdOn")
		modifiedBy = rst1("modifiedBy")
		modifiedOn = rst1("modifiedOn")
		exact = rst1("exacting")
	end if
	rst1.Close
	
	if rbcid <> "" then
		strsql = "select * from ratebuildercomponents where rbcid =" & rbcid
		rst1.Open strsql, cnn1
		if not rst1.EOF then 
			e_edc_r = rst1("e_edc_r")
			e_edc_sc9_1 = rst1("e_edc_sc9_1")
			e_edc_sc9_2 = rst1("e_edc_sc9_2")
			e_edc_sc9_3 = rst1("e_edc_sc9_3")
			e_edc_sc12_2 = rst1("e_edc_sc12_2")
			e_edc_sc1_f250_1 = rst1("e_edc_sc1_f250_1")
			e_edc_sc1_o250_1 = rst1("e_edc_sc1_o250_1")
			e_tra_r = rst1("e_tra_r")
			e_tra_sc9_1 = rst1("e_tra_sc9_1")
			e_tra_sc9_2 = rst1("e_tra_sc9_2")
			e_tra_sc9_3 = rst1("e_tra_sc9_3")
			e_tra_sc12_2 = rst1("e_tra_sc12_2")
			e_tra_sc1_f250_1 = rst1("e_tra_sc1_f250_1")
			e_tra_sc1_o250_1 = rst1("e_tra_sc1_o250_1")
			e_mac_r = rst1("e_mac_r")
			e_mac_1 = rst1("e_mac_1")
			e_sbc_r = rst1("e_sbc_r")
			e_sbc_1 = rst1("e_sbc_1")
			e_rpsp_r = rst1("e_rpsp_r")
			e_rpsp_1 = rst1("e_rpsp_1")
			e_psls_r = rst1("e_psls_r")
			e_psls_sc9_1 = rst1("e_psls_sc9_1")
			e_psls_sc12_2 = rst1("e_psls_sc12_2")	
			e_psls_sc1_1 = rst1("e_psls_sc1_1")	
			e_rdm_r = rst1("e_rdm_r")
			e_rdm_sc9_1 = rst1("e_rdm_sc9_1")
			e_rdm_sc12_2 = rst1("e_rdm_sc12_2")
			e_rdm_sc1_1 = rst1("e_rdm_sc1_1")
			e_cesds_r = rst1("e_cesds_r")
			e_cesds_1 = rst1("e_cesds_1")
			e_drs_r = rst1("e_drs_r")
			e_drs_sc9_1 = rst1("e_drs_sc9_1")
			e_drs_sc12_2 = rst1("e_drs_sc12_2")
			e_drs_sc1_1 = rst1("e_drs_sc1_1")
			e_mfc_r = rst1("e_mfc_r")
			e_mfc_sc912_1 = rst1("e_mfc_sc912_1")
			e_mfc_sc1_1 = rst1("e_mfc_sc1_1")
			e_cesss_r = rst1("e_cesss_r")
			e_cesss_1 = rst1("e_cesss_1")			
			e_sosc_sc9_1_3 = rst1("e_sosc_sc9_1_3")
			e_sosc_sc9_2 = rst1("e_sosc_sc9_2")
			e_sosc_sc12_2 = rst1("e_sosc_sc12_2")
			
			d_mc_r = rst1("d_mc_r")
			d_mc_1 = rst1("d_mc_1")
			d_o5_r = rst1("d_o5_r")
			d_o5_1 = rst1("d_o5_1")
			d_mf86_r = rst1("d_mf86_r")
			d_mf86_sc9_2 = rst1("d_mf86_sc9_2")
			d_mf86_sc9_3 = rst1("d_mf86_sc9_3")
			d_mf86_sc12_2 = rst1("d_mf86_sc12_2")
			d_mf810_r = rst1("d_mf810_r")
			d_mf810_sc9_2 = rst1("d_mf810_sc9_2")
			d_mf810_sc9_3 = rst1("d_mf810_sc9_3")
			d_mf810_sc12_2 = rst1("d_mf810_sc12_2")
			d_all_r = rst1("d_all_r")
			d_all_sc9_2 = rst1("d_all_sc9_2")
			d_all_sc9_3 = rst1("d_all_sc9_3")
			d_all_sc12_2 = rst1("d_all_sc12_2")
			
			d_tra_mc_r = rst1("d_tra_mc_r")
			d_tra_mc_1 = rst1("d_tra_mc_1")
			d_tra_o5_r = rst1("d_tra_o5_r")
			d_tra_o5_1 = rst1("d_tra_o5_1")
			d_tra_mf86_r = rst1("d_tra_mf86_r")
			d_tra_mf86_sc9_2 = rst1("d_tra_mf86_sc9_2")
			d_tra_mf86_sc9_3 = rst1("d_tra_mf86_sc9_3")
			d_tra_mf86_sc12_2 = rst1("d_tra_mf86_sc12_2")
			d_tra_mf810_r = rst1("d_tra_mf810_r")
			d_tra_mf810_sc9_2 = rst1("d_tra_mf810_sc9_2")
			d_tra_mf810_sc9_3 = rst1("d_tra_mf810_sc9_3")
			d_tra_mf810_sc12_2 = rst1("d_tra_mf810_sc12_2")
			d_tra_all_r = rst1("d_tra_all_r")
			d_tra_all_sc9_2 = rst1("d_tra_all_sc9_2")
			d_tra_all_sc9_3 = rst1("d_tra_all_sc9_3")
			d_tra_all_sc12_2 = rst1("d_tra_all_sc12_2")
			
			d_msccap_r = rst1("d_msccap_r")
			d_msccap_sc9r1_1 = rst1("d_msccap_sc9r1_1")
			d_msccap_sc9r2_1 = rst1("d_msccap_sc9r2_1")
			d_msccap_m_1 = rst1("d_msccap_m_1")
			
			d_rpd_r = rst1("d_rpd_r")
			d_rpd_1 = rst1("d_rpd_1")
			
			d_dlms_sc9_1 = rst1("d_dlms_sc9_1")
			d_dlms_sc9_2_3 = rst1("d_dlms_sc9_2_3")
			d_dlms_sc12_2 = rst1("d_dlms_sc12_2")
			d_sosc_sc9_1_3 = rst1("d_sosc_sc9_1_3")
			d_sosc_sc9_2 = rst1("d_sosc_sc9_2")
			d_sosc_sc12_2 = rst1("d_sosc_sc12_2")
			
			s_cms_r = rst1("s_cms_r")
			s_cms_sc9_1 = rst1("s_cms_sc9_1")
			s_cms_sc9_m_1 = rst1("s_cms_sc9_m_1")
			s_cms_sc9_2 = rst1("s_cms_sc9_2")
			s_cms_sc9_3		= rst1("s_cms_sc9_3")
			s_cms_sc9_m_3	= rst1("s_cms_sc9_m_3")
			s_cms_sc12_2	= rst1("s_cms_sc12_2")
			s_cms_sc1_1 = rst1("s_cms_sc1_1")
			s_cms_sc1_low_1 = rst1("s_cms_sc1_low_1")
			s_bppc_r = rst1("s_bppc_r")
			s_bppc_el_1 = rst1("s_bppc_el_1")
			s_bppc_elgs_1 = rst1("s_bppc_elgs_1")
			
			
		end if
		rst1.Close
		if exact = false then 
			rbid = null
			rbcid = null
		end if
	end if
%>

<body link="#0563C1" vlink="#954F72" class=xl97>
	<table border=0 cellpadding=0 cellspacing=0 width=1001 style='border-collapse:collapse;table-layout:fixed;width:753pt'>
		<col class=xl97 width=72 style='width:54pt'>
		<col class=xl97 width=72 style='width:54pt'>
		<col class=xl109 width=78 span=5 style='mso-width-source:userset;mso-width-alt:
		2496;width:59pt'>
		<col class=xl109 width=46 style='mso-width-source:userset;mso-width-alt:1472;
		width:35pt'>
		<col class=xl97 width=87 span=4 style='mso-width-source:userset;mso-width-alt:
		2784;width:65pt'>
		<col class=xl97 width=73 style='mso-width-source:userset;mso-width-alt:2336;
		width:55pt'>
		<col class=xl97 width=72 span=6 style='mso-width-source:userset;mso-width-alt:
		2304;width:54pt'>
		<tr height=43 style='mso-height-source:userset;height:32.25pt'>
			<td height=43 class=xl97 width=72 style='height:32.25pt;width:54pt'></td>
			<td colspan=12 class=xl148 width=929 style='border-right:1.0pt solid black;width:699pt'>
				Rate Builder
			</td>
		</tr>
		<tr height=21 style='height:15.75pt'>
			<td height=21 class=xl97 style='height:15.75pt'></td>
			<td class=xl97></td>
			<td class=xl109></td>
			<td class=xl109></td>
			<td class=xl109></td>
			<td class=xl109></td>
			<td class=xl109></td>
			<td class=xl109></td>
			<td class=xl97></td>
			<td class=xl97></td>
			<td class=xl97></td>
			<td class=xl97></td>
			<td class=xl97><%= rateperiod %></td>
		</tr>
		<tr height=31 style='height:23.25pt'>
			<form name="EditPreviousRate" method="post" action="editcomponents.asp">
				<td height=31 class=xl97 style='height:23.25pt'></td>
				<td class=xl97></td>
				<td colspan=4 class=xl154 style='border-right:1.0pt solid black'><input class=box name="emonth" value="<%=emonth%>" /><input class=box name="eyear" value="<%=eyear%>" />&nbsp;&nbsp;<input type="submit" name="action" value="Edit" class="standard" /></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td colspan=4 class=xl239 style='border-right:1.0pt solid black'><%= monthname(month, true) & " " & year %> //  <%= rbid %> || <%= rbcid %></td>
				<td class=xl97></td>
			</form>
		</tr>

		<form name="RateBuilderComponents" method="post" action="saveComponents.asp">
			<tr height=26 style='height:19.5pt'>
				<td height=26 class=xl97 style='height:19.5pt'></td>
				<td class=xl97></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl97></td>
				<td class=xl97></td>
				<td class=xl97></td>
				<td class=xl97></td>
				<td class=xl98></td>
			</tr>
			<tr height=26 style='height:19.5pt'>
				<td height=26 class=xl97 style='height:19.5pt'></td>
				<td class=xl99>&nbsp;</td>
				<td class=xl110>&nbsp;</td>
				<td class=xl110>&nbsp;</td>
				<td class=xl110>&nbsp;</td>
				<td class=xl110>&nbsp;</td>
				<td class=xl110>&nbsp;</td>
				<td class=xl110>&nbsp;</td>
				<td class=xl101>&nbsp;</td>
				<td class=xl101>&nbsp;</td>
				<td class=xl101>&nbsp;</td>
				<td class=xl101>&nbsp;</td>
				<td class=xl102>&nbsp;</td>
			</tr>
			<tr height=25 style='height:18.75pt'>
				<td height=25 class=xl97 style='height:18.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=10 class=xl151 style='border-right:1.0pt solid black'>Energy Rate (in cents)</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl135 style='border-top:none'>&nbsp;</td>
				<td class=xl136 style='border-top:none'>&nbsp;</td>
				<td class=xl136 style='border-top:none'>&nbsp;</td>
				<td class=xl136 style='border-top:none'>&nbsp;</td>
				<td class=xl136 style='border-top:none'>&nbsp;</td>
				<td class=xl136 style='border-top:none'>&nbsp;</td>
				<td class=xl136 style='border-top:none'>&nbsp;</td>
				<td class=xl136 style='border-top:none'>&nbsp;</td>
				<td class=xl136 style='border-top:none'>&nbsp;</td>
				<td class=xl137 style='border-top:none'>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='mso-height-source:userset;height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>Energy Delivery Charge</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="e_edc_r" value="<%=trim(e_edc_r)%>" <% if trim(e_edc_r) = "true" then %> checked <% end if %>/></td>
				<td class=xl227>SC9-I</td>
				<td class=xl227>SC9-II</td>
				<td class=xl227>SC9-III</td>
				<td class=xl228>SC12-II</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl229><input class=box name="e_edc_sc9_1" value="<%=e_edc_sc9_1%>" autofocus/></td>
				<td class=xl229><input class=box name="e_edc_sc9_2" value="<%=e_edc_sc9_2%>" /></td>
				<td class=xl229><input class=box name="e_edc_sc9_3" value="<%=e_edc_sc9_3%>" /></td>
				<td class=xl230><input class=box name="e_edc_sc12_2" value="<%=e_edc_sc12_2%>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='mso-height-source:userset;height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146></td>
				<td rowspan=2 class=xl226></td>
				<td class=xl227></td>
				<td class=xl227></td>
				<td class=xl227>SC1-I<br>< 250kwh</td>
				<td class=xl228>SC1-I<br>> 250kwh</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl229></td>
				<td class=xl229></td>
				<td class=xl229><input class=box name="e_edc_sc1_f250_1" value="<%=e_edc_sc1_f250_1%>" /></td>
				<td class=xl230><input class=box name="e_edc_sc1_o250_1" value="<%=e_edc_sc1_o250_1%>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>			
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td class=xl130></td>
				<td class=xl130></td>
				<td class=xl130></td>
				<td class=xl138>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='mso-height-source:userset;height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>Temporary Rate Adjustment</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="e_tra_r" value="<%= e_tra_r %>" /></td>
				<td class=xl227>SC9-I</td>
				<td class=xl227>SC9-II</td>
				<td class=xl227>SC9-III</td>
				<td class=xl228>SC12-II</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl231><input class=box name="e_tra_sc9_1" value="<%= e_tra_sc9_1 %>" /></td>
				<td class=xl231><input class=box name="e_tra_sc9_2" value="<%= e_tra_sc9_2 %>" /></td>
				<td class=xl231><input class=box name="e_tra_sc9_3" value="<%= e_tra_sc9_3 %>" /></td>
				<td class=xl232><input class=box name="e_tra_sc12_2" value="<%= e_tra_sc12_2  %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>			
			<tr height=21 style='mso-height-source:userset;height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146></td>
				<td rowspan=2 class=xl226></td>
				<td class=xl227></td>
				<td class=xl227></td>
				<td class=xl227>SC1-I<br>< 250kwh</td>
				<td class=xl228>SC1-I<br>> 250kwh</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl231></td>
				<td class=xl231></td>
				<td class=xl231><input class=box name="e_tra_sc1_f250_1" value="<%= e_tra_sc1_f250_1 %>" /></td>
				<td class=xl232><input class=box name="e_tra_sc1_o250_1" value="<%= e_tra_sc1_o250_1 %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>MAC</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="e_mac_r" value="<%= e_mac_r %>" /></td>
				<td class=xl231>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl228>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl232><input class=box name="e_mac_1" value="<%= e_mac_1 %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>Systems Benefits Charge</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="e_sbc_r" value="<%= e_sbc_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl228>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl232><input class=box name="e_sbc_1" value="<%= e_sbc_1 %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>Renewable Portfolio Standard Program</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="e_rpsp_r" value="<%= e_rpsp_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl228>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl232><input class=box name="e_rpsp_1" value="<%= e_rpsp_1 %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='mso-height-source:userset;height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>PSL Surcharge</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="e_psls_r" value="<%= e_psls_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl227>SC9</td>
				<td class=xl227>SC12</td>
				<td class=xl228>SC1</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231><input class=box name="e_psls_sc9_1" value="<%= e_psls_sc9_1 %>" /></td>
				<td class=xl231><input class=box name="e_psls_sc12_2" value="<%= e_psls_sc12_2 %>" /></td>
				<td class=xl232><input class=box name="e_psls_sc1_1" value="<%= e_psls_sc1_1 %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='mso-height-source:userset;height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>Revenue Decoupling Mechanism<span
				style='mso-spacerun:yes'> </span></td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="e_rdm_r" value="<%= e_rdm_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl227>SC9</td>
				<td class=xl227>SC12</td>
				<td class=xl228>SC1</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231><input class=box name="e_rdm_sc9_1" value="<%=  e_rdm_sc9_1 %>" /></td>
				<td class=xl231><input class=box name="e_rdm_sc12_2" value="<%= e_rdm_sc12_2 %>" /></td>
				<td class=xl232><input class=box name="e_rdm_sc1_1" value="<%= e_rdm_sc1_1 %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>Clean Energy Standard Delivery Surcharge</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="e_cesds_r" value="<%= e_cesds_r %>" /></td>
				<td class=xl231>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl228>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl232><input class=box name="e_cesds_1" value="<%= e_cesds_1 %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>		
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='mso-height-source:userset;height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl157>Delivery Revenue Surcharge</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="e_drs_r" value="<%= e_drs_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl227>SC9</td>
				<td class=xl227>SC12</td>
				<td class=xl228>SC1</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231><input class=box name="e_drs_sc9_1" value="<%= e_drs_sc9_1 %>" /></td>
				<td class=xl231><input class=box name="e_drs_sc12_2" value="<%=e_drs_sc12_2 %>" /></td>
				<td class=xl232><input class=box name="e_drs_sc1_1" value="<%=e_drs_sc1_1 %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>Merchant Function Charge (SC9/12)</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="e_mfc_r" value="<%= e_mfc_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl227>SC9/12</td>
				<td class=xl228>SC1</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231><input class=box name="e_mfc_sc912_1" value="<%= e_mfc_sc912_1 %>" /></td>
				<td class=xl232><input class=box name="e_mfc_sc1_1" value="<%= e_mfc_sc1_1 %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl140>&nbsp;</td>
				<td class=xl133></td>
				<td class=xl133></td>
				<td class=xl133></td>
				<td class=xl133></td>
				<td class=xl132></td>
				<td></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146 style='border-bottom:1.0pt solid black'>Clean Energy Standard Supply Surcharge</td>
				<td rowspan=2 class=xl226 style='border-bottom:1.0pt solid black'>
					R<input tabindex="999" type="checkbox" name="e_cesss_r" value="<%= e_cesss_r %>" /></td>
				<td class=xl231>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl228>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=22 style='height:16.5pt'>
				<td height=22 class=xl97 style='height:16.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl235>&nbsp;</td>
				<td class=xl235>&nbsp;</td>
				<td class=xl236>&nbsp;</td>
				<td class=xl237><input class=box name="e_cesss_1" value="<%= e_cesss_1 %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl140>&nbsp;</td>
				<td class=xl133></td>
				<td class=xl133></td>
				<td class=xl133></td>
				<td class=xl133></td>
				<td class=xl132></td>
				<td></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='mso-height-source:userset;height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>Statement of Sur-credit</td>
				<td rowspan=2 class=xl226 >
					R<input tabindex="999" type="checkbox" name="d_msccap_r" value="<%= d_msccap_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl227>SC9 I & III</td>
				<td class=xl227>SC9 II</td>
				<td class=xl228>SC12 II</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231><input class=box name="e_sosc_sc9_1_3" value="<%= e_sosc_sc9_1_3 %>" /></td>
				<td class=xl231><input class=box name="e_sosc_sc9_2" value="<%=   e_sosc_sc9_2 %>" /></td>
				<td class=xl232><input class=box name="e_sosc_sc12_2" value="<%=  e_sosc_sc12_2  %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>

			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl103>&nbsp;</td>
				<td class=xl111>&nbsp;</td>
				<td class=xl111>&nbsp;</td>
				<td class=xl111>&nbsp;</td>
				<td class=xl111>&nbsp;</td>
				<td class=xl111>&nbsp;</td>
				<td class=xl111>&nbsp;</td>
				<td class=xl104>&nbsp;</td>
				<td class=xl104>&nbsp;</td>
				<td class=xl104>&nbsp;</td>
				<td class=xl104>&nbsp;</td>
				<td class=xl105>&nbsp;</td>
			</tr>

			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl97></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl97></td>
				<td class=xl97></td>
				<td class=xl97></td>
				<td class=xl97></td>
				<td class=xl97></td>
			</tr>
			<tr height=22 style='height:16.5pt'>
				<td height=22 class=xl97 style='height:16.5pt'></td>
				<td class=xl99>&nbsp;</td>
				<td class=xl112>&nbsp;</td>
				<td class=xl112>&nbsp;</td>
				<td class=xl112>&nbsp;</td>
				<td class=xl112>&nbsp;</td>
				<td class=xl112>&nbsp;</td>
				<td class=xl112>&nbsp;</td>
				<td class=xl106>&nbsp;</td>
				<td class=xl106>&nbsp;</td>
				<td class=xl106>&nbsp;</td>
				<td class=xl106>&nbsp;</td>
				<td class=xl107>&nbsp;</td>
			</tr>
			<tr height=25 style='height:18.75pt'>
				<td height=25 class=xl97 style='height:18.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=10 class=xl151 style='border-right:1.0pt solid black'>Demand Rate (in dollars)</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl141>&nbsp;</td>
				<td class=xl134></td>
				<td class=xl134></td>
				<td class=xl134></td>
				<td class=xl134></td>
				<td class=xl134></td>
				<td class=xl134></td>
				<td class=xl134></td>
				<td class=xl134></td>
				<td class=xl108>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=25 style='height:18.75pt'>
				<td height=25 class=xl97 style='height:18.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>first 5 kW(or less)(minimum charge)</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="d_mc_r" value="<%= d_mc_r %>" /></td>
				<td class=xl238>&nbsp;</td>
				<td class=xl238>&nbsp;</td>
				<td class=xl238>&nbsp;</td>
				<td class=xl228>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl232><input class=box name="d_mc_1" value="<%= d_mc_1 %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>Over 5 kW</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="d_o5_r" value="<%= d_o5_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl228>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl232><input class=box name="d_o5_1" value="<%= d_o5_1 %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=22 style='mso-height-source:userset;height:16.5pt'>
				<td height=22 class=xl97 style='height:16.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>M-F 8am to 6pm</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="d_mf86_r" value="<%= d_mf86_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl227>SC9-II</td>
				<td class=xl227>SC9-III</td>
				<td class=xl228>SC12-II</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231><input class=box name="d_mf86_sc9_2" value="<%= d_mf86_sc9_2 %>" /></td>
				<td class=xl231><input class=box name="d_mf86_sc9_3" value="<%= d_mf86_sc9_3 %>" /></td>
				<td class=xl232><input class=box name="d_mf86_sc12_2" value="<%=d_mf86_sc12_2 %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='mso-height-source:userset;height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>M-F 8am to 10pm</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="d_mf810_r" value="<%= d_mf810_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl227>SC9-II</td>
				<td class=xl227>SC9-III</td>
				<td class=xl228>SC12-II</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231><input class=box name="d_mf810_sc9_2" value="<%=  d_mf810_sc9_2 %>" /></td>
				<td class=xl231><input class=box name="d_mf810_sc9_3" value="<%=  d_mf810_sc9_3 %>" /></td>
				<td class=xl232><input class=box name="d_mf810_sc12_2" value="<%= d_mf810_sc12_2  %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='mso-height-source:userset;height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>all hours of all days</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="d_all_r" value="<%= d_all_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl227>SC9-II</td>
				<td class=xl227>SC9-III</td>
				<td class=xl228>SC12-II</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231><input class=box name="d_all_sc9_2" value="<%= d_all_sc9_2 %>" /></td>
				<td class=xl231><input class=box name="d_all_sc9_3" value="<%= d_all_sc9_3 %>" /></td>
				<td class=xl232><input class=box name="d_all_sc12_2" value="<%=d_all_sc12_2  %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>Temp. Rate Adj.-first 5 kW(or less)(minimum charge)</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="d_tra_mc_r" value="<%= d_tra_mc_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl228>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl232><input class=box name="d_tra_mc_1" value="<%= d_tra_mc_1 %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>Temp. Rate Adj.-Over 5 kW</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="d_tra_o5_r" value="<%= d_tra_o5_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl228>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl232><input class=box name="d_tra_o5_1" value="<%= d_tra_o5_1 %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='mso-height-source:userset;height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>Temp. Rate Adj.-M-F 8am to 6pm</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="d_tra_mf86_r" value="<%= d_tra_mf86_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl227>SC9-II</td>
				<td class=xl227>SC9-III</td>
				<td class=xl228>SC12-II</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231><input class=box name="d_tra_mf86_sc9_2" value="<%=  d_tra_mf86_sc9_2 %>" /></td>
				<td class=xl231><input class=box name="d_tra_mf86_sc9_3" value="<%=  d_tra_mf86_sc9_3 %>" /></td>
				<td class=xl232><input class=box name="d_tra_mf86_sc12_2" value="<%= d_tra_mf86_sc12_2  %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='mso-height-source:userset;height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>Temp. Rate Adj.-M-F 8am to 10pm</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="d_tra_mf810_r" value="<%= d_tra_mf810_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl227>SC9-II</td>
				<td class=xl227>SC9-III</td>
				<td class=xl228>SC12-II</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231><input class=box name="d_tra_mf810_sc9_2" value="<%=  d_tra_mf810_sc9_2 %>" /></td>
				<td class=xl231><input class=box name="d_tra_mf810_sc9_3" value="<%=  d_tra_mf810_sc9_3 %>" /></td>
				<td class=xl232><input class=box name="d_tra_mf810_sc12_2" value="<%= d_tra_mf810_sc12_2  %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
				<tr height=21 style='mso-height-source:userset;height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>Temp. Rate Adj.-all hours of all days</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="d_tra_all_r" value="<%= d_tra_all_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl227>SC9-II</td>
				<td class=xl227>SC9-III</td>
				<td class=xl228>SC12-II</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231><input class=box name="d_tra_all_sc9_2" value="<%=  d_tra_all_sc9_2 %>" /></td>
				<td class=xl231><input class=box name="d_tra_all_sc9_3" value="<%=  d_tra_all_sc9_3 %>" /></td>
				<td class=xl232><input class=box name="d_tra_all_sc12_2" value="<%= d_tra_all_sc12_2  %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=22 style='mso-height-source:userset;height:16.5pt'>
				<td height=22 class=xl97 style='height:16.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>Statement of Sur-credit</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="d_mf86_r" value="<%= d_mf86_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl227>SC9-I & III</td>
				<td class=xl227>SC9-II</td>
				<td class=xl228>SC12-II</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231><input class=box name="d_sosc_sc9_1_3" value="<%= d_sosc_sc9_1_3 %>" /></td>
				<td class=xl231><input class=box name="d_sosc_sc9_2" value="<%= d_sosc_sc9_2 %>" /></td>
				<td class=xl232><input class=box name="d_sosc_sc12_2" value="<%=d_sosc_sc12_2 %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl140>&nbsp;</td>
				<td class=xl133></td>
				<td class=xl133></td>
				<td class=xl133></td>
				<td class=xl133></td>
				<td class=xl132></td>
				<td></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>			
			<tr height=21 style='mso-height-source:userset;height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>MSC-CAP</td>
				<td rowspan=2 class=xl226 >
					R<input tabindex="999" type="checkbox" name="d_msccap_r" value="<%= d_msccap_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl227>Rider M</td>
				<td class=xl227>SC9R1</td>
				<td class=xl228>SC9R2</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231><input class=box name="d_msccap_m_1" value="<%= d_msccap_m_1 %>" /></td>
				<td class=xl231><input class=box name="d_msccap_sc9r1_1" value="<%= d_msccap_sc9r1_1 %>" /></td>
				<td class=xl232><input class=box name="d_msccap_sc9r2_1" value="<%= d_msccap_sc9r2_1  %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>			
			<tr height=21 style='mso-height-source:userset;height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>Dynamic Load Management Surcharge</td>
				<td rowspan=2 class=xl226 >
					R<input tabindex="999" type="checkbox" name="d_dlms_r" value="<%= d_dlms_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl227>SC9-I</td>
				<td class=xl227>SC9-II & III</td>
				<td class=xl228>SC12-II</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231><input class=box name="d_dlms_sc9_1" value="<%= d_dlms_sc9_1 %>" /></td>
				<td class=xl231><input class=box name="d_dlms_sc9_2_3" value="<%= d_dlms_sc9_2_3 %>" /></td>
				<td class=xl232><input class=box name="d_dlms_sc12_2" value="<%= d_dlms_sc12_2  %>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>				
			<tr height=21 style='mso-height-source:userset;height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146 style='border-bottom:1.0pt solid black'>Reactive-Power demand($/kVar)
</td>
				<td rowspan=2 class=xl226 style='border-bottom:1.0pt solid black'>
					R<input tabindex="999" type="checkbox" name="d_rpd_r" value="<%= d_rpd_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl228>&nbsp;</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=22 style='height:16.5pt'>
				<td height=22 class=xl97 style='height:16.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl235>&nbsp;</td>
				<td class=xl235>&nbsp;</td>
				<td class=xl235>&nbsp;</td>
				<td class=xl237><input class=box name="d_rpd_1" value="<%= d_rpd_1%>" /></td>
				<td class=xl96>&nbsp;</td>
			</tr>
			
			<tr height=22 style='height:16.5pt'>
				<td height=22 class=xl97 style='height:16.5pt'></td>
				<td class=xl103>&nbsp;</td>
				<td class=xl113>&nbsp;</td>
				<td class=xl113>&nbsp;</td>
				<td class=xl113>&nbsp;</td>
				<td class=xl113>&nbsp;</td>
				<td class=xl113>&nbsp;</td>
				<td class=xl113>&nbsp;</td>
				<td class=xl94>&nbsp;</td>
				<td class=xl94>&nbsp;</td>
				<td class=xl94>&nbsp;</td>
				<td class=xl104>&nbsp;</td>
				<td class=xl105>&nbsp;</td>
			</tr>
			<tr height=22 style='height:16.5pt'>
				<td height=22 class=xl97 style='height:16.5pt'></td>
				<td class=xl97></td>
				<td class=xl114></td>
				<td class=xl114></td>
				<td class=xl114></td>
				<td class=xl114></td>
				<td class=xl114></td>
				<td class=xl114></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl97></td>
				<td class=xl97></td>
			</tr>
			<tr height=22 style='height:16.5pt'>
				<td height=22 class=xl97 style='height:16.5pt'></td>
				<td class=xl99>&nbsp;</td>
				<td class=xl112>&nbsp;</td>
				<td class=xl112>&nbsp;</td>
				<td class=xl112>&nbsp;</td>
				<td class=xl112>&nbsp;</td>
				<td class=xl112>&nbsp;</td>
				<td class=xl112>&nbsp;</td>
				<td class=xl106>&nbsp;</td>
				<td class=xl106>&nbsp;</td>
				<td class=xl106>&nbsp;</td>
				<td class=xl106>&nbsp;</td>
				<td class=xl107>&nbsp;</td>
			</tr>
			<tr height=26 style='height:19.5pt'>
				<td height=26 class=xl97 style='height:19.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=10 class=xl151 style='border-right:1.0pt solid black'>Static (in dollars)</td>
				<td class=xl108>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl141>&nbsp;</td>
				<td class=xl134></td>
				<td class=xl134></td>
				<td class=xl134></td>
				<td class=xl134></td>
				<td class=xl134></td>
				<td class=xl134></td>
				<td class=xl134></td>
				<td class=xl134></td>
				<td class=xl108>&nbsp;</td>
				<td class=xl108>&nbsp;</td>
			</tr>
			<tr height=25 style='height:18.75pt'>
				<td height=25 class=xl97 style='height:18.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146>Charges for Metering Services (not for Rider M)</td>
				<td rowspan=2 class=xl226>
					R<input tabindex="999" type="checkbox" name="s_cms_r" value="<%= s_cms_r %>" /></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl227>SC9-I</td>
				<td class=xl227>SC9-I Rider M</td>
				<td class=xl228>SC9-II</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl233><input class=box name="s_cms_sc9_1" value="<%= s_cms_sc9_1 %>" /></td>
				<td class=xl233><input class=box name="s_cms_sc9_m_1" value="<%= s_cms_sc9_m_1 %>" /></td>
				<td class=xl232><input class=box name="s_cms_sc9_2" value="<%= s_cms_sc9_2 %>" /></td>
				<td class=xl93>&nbsp;</td>
			</tr><tr height=25 style='height:18.75pt'>
				<td height=25 class=xl97 style='height:18.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146></td>
				<td rowspan=2 class=xl226></td>
				<td class=xl233>&nbsp;</td>
				<td class=xl227>SC9-III</td>
				<td class=xl227>SC9-III Rider M</td>
				<td class=xl228>SC12-II</td>
				<td class=xl96>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl233>&nbsp;</td>
				<td class=xl233><input class=box name="s_cms_sc9_3" value="<%= s_cms_sc9_3 %>" /></td>
				<td class=xl233><input class=box name="s_cms_sc9_m_3" value="<%= s_cms_sc9_m_3 %>" /></td>
				<td class=xl232><input class=box name="s_cms_sc12_2" value="<%= s_cms_sc12_2 %>" /></td>
				<td class=xl93>&nbsp;</td>
			</tr>
			<tr class=xl97 height=10 style='mso-height-source:userset;height:7.5pt'>
				<td height=10 class=xl97 style='height:7.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl139>&nbsp;</td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl131></td>
				<td class=xl132></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl92></td>
				<td class=xl93>&nbsp;</td>
				<td class=xl93>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl95>&nbsp;</td>
				<td colspan=5 rowspan=2 class=xl146 style='border-bottom:1.0pt solid black'>Billing	and Payment Processing Charges</td>
				<td rowspan=2 class=xl226 style='border-bottom:1.0pt solid black'>
					R<input tabindex="999" type="checkbox" name="s_bppc_r" value="<%= s_bppc_r %>" /></td>
				<td class=xl231>&nbsp;</td>
				<td class=xl231>&nbsp;</td>
				<td class=xl227>Electric</td>
				<td class=xl228>Electric/Gas</td>
				<td class=xl93>&nbsp;</td>
			</tr>
			<tr height=22 style='height:16.5pt'>
				<td height=22 class=xl97 style='height:16.5pt'></td>
				<td class=xl95>&nbsp;</td>
				<td class=xl236>&nbsp;</td>
				<td class=xl236>&nbsp;</td>
				<td class=xl236><input class=box name="s_bppc_el_1" value="<%= s_bppc_el_1 %>" /></td>
				<td class=xl237><input class=box name="s_bppc_elgs_1" value="<%= s_bppc_elgs_1 %>" /></td>
				<td class=xl93>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl103>&nbsp;</td>
				<td class=xl111>&nbsp;</td>
				<td class=xl111>&nbsp;</td>
				<td class=xl111>&nbsp;</td>
				<td class=xl111>&nbsp;</td>
				<td class=xl111>&nbsp;</td>
				<td class=xl111>&nbsp;</td>
				<td class=xl104>&nbsp;</td>
				<td class=xl104>&nbsp;</td>
				<td class=xl104>&nbsp;</td>
				<td class=xl104>&nbsp;</td>
				<td class=xl105>&nbsp;</td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl97></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl97></td>
				<td class=xl97></td>
				<td class=xl97></td>
				<td class=xl97></td>
				<td class=xl97></td>
			</tr>
			<tr height=20 style='height:15.0pt'>
				<td height=20 class=xl97 style='height:15.0pt'></td>
				<td class=xl97></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td colspan=4 rowspan=3 class=xl161 style='border-right:1.0pt solid black;
				border-bottom:1.0pt solid black'><input type="submit" name="action" value="Save & Calculate" class="standard" onclick="saveComponents()"/></td>
				<td class=xl97></td>
			</tr>
			<tr height=20 style='height:15.0pt'>
				<td height=20 class=xl97 style='height:15.0pt'></td>
				<td class=xl97></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl97></td>
			</tr>
			<tr height=21 style='height:15.75pt'>
				<td height=21 class=xl97 style='height:15.75pt'></td>
				<td class=xl97></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl109></td>
				<td class=xl97></td>
			</tr>
			<![if supportMisalignedColumns]>
			<tr height=0 style='display:none'>
			<td width=72 style='width:54pt'></td>
			<td width=72 style='width:54pt'></td>
			<td width=78 style='width:59pt'></td>
			<td width=78 style='width:59pt'></td>
			<td width=78 style='width:59pt'></td>
			<td width=78 style='width:59pt'></td>
			<td width=78 style='width:59pt'></td>
			<td width=46 style='width:35pt'></td>
			<td width=87 style='width:65pt'></td>
			<td width=87 style='width:65pt'></td>
			<td width=87 style='width:65pt'></td>
			<td width=87 style='width:65pt'></td>
			<td width=73 style='width:55pt'></td>
			</tr>
			<![endif]>
		</table>
		
		<input type="hidden" value="<%=today%>" name="rateperiod"/>        
		<input type="hidden" value="<%=rbcid%>" name="rbcid"/>  
	</form>
</body>

</html>
