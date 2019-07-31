<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/nonsecure.inc"-->
<%
'if not(allowGroups("Genergy Users")) then
'%><%
'end if

			dim cnn1, rst1, insertSql, ssql
			set cnn1 = server.createobject("ADODB.connection")
			set rst1 = server.createobject("ADODB.recordset")
			cnn1.open getConnect(0,0,"dbCore")
		%>
		<%
			function toBool(val)
				if val="" then 
					val=0
				else
					val = 1
				end if
				toBool = val
			end function
			function toNumb(val)
				if val="" then
					val = 0
				end if
				toNumb = val
			end function
		%>
		<% 	
			dim rbid,rbcid,rbrid,rateperiod,createdBy,createdOn,modifiedBy,modifiedOn,rp

			dim e_edc_r,e_edc_sc9_1,e_edc_sc9_2,e_edc_sc9_3,e_edc_sc12_2,e_edc_sc1_f250_1,e_edc_sc1_o250_1,e_tra_r,e_tra_sc9_1,e_tra_sc9_2,e_tra_sc9_3,e_tra_sc12_2,e_tra_sc1_f250_1,e_tra_sc1_o250_1,e_mac_r,e_mac_1,e_sbc_r,e_sbc_1,e_rpsp_r,e_rpsp_1,e_psls_r,e_psls_sc9_1,e_psls_sc12_2,e_psls_sc1_1,e_rdm_r,e_rdm_sc9_1,e_rdm_sc12_2,e_rdm_sc1_1,e_drs_r,e_drs_sc9_1,e_drs_sc12_2,e_drs_sc1_1,e_mfc_r,e_mfc_sc912_1,e_mfc_sc1_1
			dim d_mc_r,d_mc_1,d_o5_r,d_o5_1,d_mf86_r,d_mf86_sc9_2,d_mf86_sc9_3,d_mf86_sc12_2,d_mf810_r,d_mf810_sc9_2, d_mf810_sc9_3,d_mf810_sc12_2,d_all_r,d_all_sc9_2,d_all_sc9_3,d_all_sc12_2,d_tra_mc_r,d_tra_mc_1,d_tra_o5_r,d_tra_o5_1,d_tra_mf86_r,d_tra_mf86_sc9_2,d_tra_mf86_sc9_3
			dim d_tra_mf86_sc12_2,d_tra_mf810_r,d_tra_mf810_sc9_2,d_tra_mf810_sc9_3,d_tra_mf810_sc12_2,d_tra_all_r,d_tra_all_sc9_2,d_tra_all_sc9_3,d_tra_all_sc12_2,d_msccap_r,d_msccap_m_1,d_msccap_sc9r1_1,d_msccap_sc9r2_1,d_rpd_r,d_rpd_1
			dim s_cms_r,s_cms_sc9_m_1,s_cms_sc9_1,s_cms_sc9_2,s_cms_sc1_1,s_cms_sc1_low_1,s_bppc_r,s_bppc_el_1,s_bppc_elgs_1, e_cesds_1, e_cesss_1, e_cesds_r, e_cesss_r, s_cms_sc9_3, s_cms_sc9_m_3, s_cms_sc12_2, d_dlms_sc9_1,d_dlms_sc9_2_3, d_dlms_sc12_2, d_dlms_r
			dim e_sosc_sc9_1_3, d_sosc_sc9_1_3, e_sosc_sc9_2, d_sosc_sc9_2, e_sosc_sc12_2, d_sosc_sc12_2
			
			e_edc_r = toBool(request.form("e_edc_r"))
			e_edc_sc9_1 = toNumb(request.form("e_edc_sc9_1"))
			e_edc_sc9_2 = toNumb(request.form("e_edc_sc9_2"))
			e_edc_sc9_3 = toNumb(request.form("e_edc_sc9_3"))
			e_edc_sc12_2 = toNumb(request.form("e_edc_sc12_2"))
			e_edc_sc1_f250_1 = toNumb(request.form("e_edc_sc1_f250_1"))
			e_edc_sc1_o250_1 = toNumb(request.form("e_edc_sc1_o250_1"))
			e_tra_r = toBool(request.form("e_tra_r"))
			e_tra_sc9_1 = toNumb(request.form("e_tra_sc9_1"))
			e_tra_sc9_2 = toNumb(request.form("e_tra_sc9_2"))
			e_tra_sc9_3 = toNumb(request.form("e_tra_sc9_3"))
			e_tra_sc12_2 = toNumb(request.form("e_tra_sc12_2"))
			e_tra_sc1_f250_1 = toNumb(request.form("e_tra_sc1_f250_1"))
			e_tra_sc1_o250_1 = toNumb(request.form("e_tra_sc1_o250_1"))
			e_mac_r = toBool(request.form("e_mac_r"))
			e_mac_1 = toNumb(request.form("e_mac_1"))
			e_sbc_r = toBool(request.form("e_sbc_r"))
			e_sbc_1 = toNumb(request.form("e_sbc_1"))
			e_rpsp_r = toBool(request.form("e_rpsp_r"))
			e_rpsp_1 = toNumb(request.form("e_rpsp_1"))
			e_psls_r = toBool(request.form("e_psls_r"))
			e_psls_sc9_1 = toNumb(request.form("e_psls_sc9_1"))
			e_psls_sc12_2 = toNumb(request.form("e_psls_sc12_2"))
			e_psls_sc1_1 = toNumb(request.form("e_psls_sc1_1"))
			e_rdm_r = toBool(request.form("e_rdm_r"))
			e_rdm_sc9_1 = toNumb(request.form("e_rdm_sc9_1"))
			e_rdm_sc12_2 = toNumb(request.form("e_rdm_sc12_2"))
			e_rdm_sc1_1 = toNumb(request.form("e_rdm_sc1_1"))
			e_cesds_r  = tobool(request.form("e_cesds_r"))
			e_cesds_1 = toNumb(request.form("e_cesds_1"))
			e_drs_r = toBool(request.form("e_drs_r"))
			e_drs_sc9_1 = toNumb(request.form("e_drs_sc9_1"))
			e_drs_sc12_2 = toNumb(request.form("e_drs_sc12_2"))
			e_drs_sc1_1 = toNumb(request.form("e_drs_sc1_1"))
			e_mfc_r = toBool(request.form("e_mfc_r"))
			e_mfc_sc912_1 = toNumb(request.form("e_mfc_sc912_1"))
			e_mfc_sc1_1 = toNumb(request.form("e_mfc_sc1_1"))
			e_cesss_r  = tobool(request.form("e_cesss_r"))
			e_cesss_1 = toNumb(request.form("e_cesss_1"))
			e_sosc_sc9_1_3 = 	toNumb(request.form("e_sosc_sc9_1_3"))
			e_sosc_sc9_2 =		toNumb(request.form("e_sosc_sc9_2"))
			e_sosc_sc12_2 =     toNumb(request.form("e_sosc_sc12_2"))
			
			d_mc_r = toBool(request.form("d_mc_r"))
			d_mc_1 = toNumb(request.form("d_mc_1"))
			d_o5_r = toBool(request.form("d_o5_r"))
			d_o5_1 = toNumb(request.form("d_o5_1"))
			d_mf86_r = toBool(request.form("d_mf86_r"))
			d_mf86_sc9_2 = toNumb(request.form("d_mf86_sc9_2"))
			d_mf86_sc9_3 = toNumb(request.form("d_mf86_sc9_3"))
			d_mf86_sc12_2 = toNumb(request.form("d_mf86_sc12_2"))
			d_mf810_r = toBool(request.form("d_mf810_r"))
			d_mf810_sc9_2 = toNumb(request.form("d_mf810_sc9_2"))
			d_mf810_sc9_3 = toNumb(request.form("d_mf810_sc9_3"))
			d_mf810_sc12_2 = toNumb(request.form("d_mf810_sc12_2"))
			d_all_r = toBool(request.form("d_all_r"))
			d_all_sc9_2 = toNumb(request.form("d_all_sc9_2"))
			d_all_sc9_3 = toNumb(request.form("d_all_sc9_3"))
			d_all_sc12_2 = toNumb(request.form("d_all_sc12_2"))
			d_tra_mc_r = toBool(request.form("d_tra_mc_r"))
			d_tra_mc_1 = toNumb(request.form("d_tra_mc_1"))
			d_tra_o5_r = toBool(request.form("d_tra_o5_r"))
			d_tra_o5_1 = toNumb(request.form("d_tra_o5_1"))
			d_tra_mf86_r = toBool(request.form("d_tra_mf86_r"))
			d_tra_mf86_sc9_2 = toNumb(request.form("d_tra_mf86_sc9_2"))
			d_tra_mf86_sc9_3 = toNumb(request.form("d_tra_mf86_sc9_3"))
			d_tra_mf86_sc12_2 = toNumb(request.form("d_tra_mf86_sc12_2"))
			d_tra_mf810_r = toBool(request.form("d_tra_mf810_r"))
			d_tra_mf810_sc9_2 = toNumb(request.form("d_tra_mf810_sc9_2"))
			d_tra_mf810_sc9_3 = toNumb(request.form("d_tra_mf810_sc9_3"))
			d_tra_mf810_sc12_2 = toNumb(request.form("d_tra_mf810_sc12_2"))
			d_tra_all_r = toBool(request.form("d_tra_all_r"))
			d_tra_all_sc9_2 = toNumb(request.form("d_tra_all_sc9_2"))
			d_tra_all_sc9_3 = toNumb(request.form("d_tra_all_sc9_3"))
			d_tra_all_sc12_2 = toNumb(request.form("d_tra_all_sc12_2"))
			
			d_msccap_r = toBool(request.form("d_msccap_r"))
			d_msccap_sc9r1_1 = toNumb(request.form("d_msccap_sc9r1_1"))
			d_msccap_sc9r2_1 = toNumb(request.form("d_msccap_sc9r2_1"))
			d_msccap_m_1 = toNumb(request.form("d_msccap_m_1"))
			d_dlms_r = toBool(request.form("d_dlms_r"))
			d_dlms_sc9_1 = toNumb(request.form("d_dlms_sc9_1"))
			d_dlms_sc9_2_3 = toNumb(request.form("d_dlms_sc9_2_3"))
			d_dlms_sc12_2 = toNumb(request.form("d_dlms_sc12_2"))
			d_rpd_r = toBool(request.form("d_rpd_r"))
			d_rpd_1 = toNumb(request.form("d_rpd_1"))
			d_sosc_sc9_1_3 = 	toNumb(request.form("d_sosc_sc9_1_3"))
			d_sosc_sc9_2 =		toNumb(request.form("d_sosc_sc9_2"))
			d_sosc_sc12_2 =     toNumb(request.form("d_sosc_sc12_2"))
			
			s_cms_r = toBool(request.form("s_cms_r"))
			s_cms_sc9_m_1 = toNumb(request.form("s_cms_sc9_m_1"))
			s_cms_sc9_1 = toNumb(request.form("s_cms_sc9_1"))
			s_cms_sc9_2 = toNumb(request.form("s_cms_sc9_2"))
			s_cms_sc1_1 = toNumb(request.form("s_cms_sc1_1"))
			s_cms_sc1_low_1 = toNumb(request.form("s_cms_sc1_low_1"))
			s_cms_sc9_3		=	toNumb(request.form("s_cms_sc9_3"))
			s_cms_sc9_m_3	=   toNumb(request.form("s_cms_sc9_m_3"))
			s_cms_sc12_2	=	toNumb(request.form("s_cms_sc12_2"))
			s_bppc_r = toBool(request.form("s_bppc_r"))
			s_bppc_el_1 = toNumb(request.form("s_bppc_el_1"))
			s_bppc_elgs_1 = toNumb(request.form("s_bppc_elgs_1"))
			
			rp = request.form("rateperiod")
			rateperiod = year(rp)&","&month(rp)&","&day(rp)
			rbcid = request.form("rbcid")
			createdby = getXmlUserName()
			modifiedby = getXmlUserName()
		%>

		<% 
			if rbcid <> "" then
				insertsql = "UPDATE [dbo].[RateBuilderComponents] SET [d_dlms_sc9_1] = '"&d_dlms_sc9_1&"',[d_dlms_sc9_2_3] = '"&d_dlms_sc9_2_3&"',[d_dlms_sc12_2] = '"&d_dlms_sc12_2&"',[d_dlms_r] = '"&d_dlms_r&"',[e_cesds_1] = '"&e_cesds_1&"',[e_cesds_r] = '"&e_cesds_r&"', [e_edc_r] = '"&e_edc_r&"',[e_edc_sc9_1] = '"&e_edc_sc9_1&"',[e_edc_sc9_2] = '"&e_edc_sc9_2&"',[e_edc_sc9_3] = '"&e_edc_sc9_3&"',[e_edc_sc12_2] = '"&e_edc_sc12_2&"',[e_edc_sc1_f250_1] = '"&e_edc_sc1_f250_1&"',[e_edc_sc1_o250_1] = '"&e_edc_sc1_o250_1&"',[e_tra_r] = '"&e_tra_r&"',[e_tra_sc9_1] = '"&e_tra_sc9_1&"',[e_tra_sc9_2] = '"&e_tra_sc9_2&"',[e_tra_sc9_3] = '"&e_tra_sc9_3&"',[e_tra_sc12_2] = '"&e_tra_sc12_2&"',[e_tra_sc1_f250_1] = '"&e_tra_sc1_f250_1&"',[e_tra_sc1_o250_1] = '"&e_tra_sc1_o250_1&"',[e_mac_r] = '"&e_mac_r&"',[e_mac_1] = '"&e_mac_1&"',[e_sbc_r] = '"&e_sbc_r&"',[e_sbc_1] = '"&e_sbc_1&"',[e_rpsp_r] = '"&e_rpsp_r&"',[e_rpsp_1] = '"&e_rpsp_1&"',[e_psls_r] = '"&e_psls_r&"',[e_psls_sc9_1] = '"&e_psls_sc9_1&"',[e_psls_sc12_2] = '"&e_psls_sc12_2&"',[e_psls_sc1_1] = '"&e_psls_sc1_1&"',[e_rdm_r] = '"&e_rdm_r&"',[e_rdm_sc9_1] = '"&e_rdm_sc9_1&"',[e_rdm_sc12_2] = '"&e_rdm_sc12_2&"',[e_rdm_sc1_1] = '"&e_rdm_sc1_1&"',[e_drs_r] = '"&e_drs_r&"',[e_drs_sc9_1] = '"&e_drs_sc9_1&"',[e_drs_sc12_2] = '"&e_drs_sc12_2&"',[e_drs_sc1_1] = '"&e_drs_sc1_1&"',[e_mfc_r] = '"&e_mfc_r&"',[e_mfc_sc912_1] = '"&e_mfc_sc912_1&"',[e_mfc_sc1_1] = '"&e_mfc_sc1_1&"',[d_mc_r] = '"&d_mc_r&"',[d_mc_1] = '"&d_mc_1&"',[d_o5_r] = '"&d_o5_r&"',[d_o5_1] = '"&d_o5_1&"',[d_mf86_r] = '"&d_mf86_r&"',[d_mf86_sc9_2] = '"&d_mf86_sc9_2&"',[d_mf86_sc9_3] = '"&d_mf86_sc9_3&"',[d_mf86_sc12_2] = '"&d_mf86_sc12_2&"',[d_mf810_r] = '"&d_mf810_r&"',[d_mf810_sc9_2] = '"&d_mf810_sc9_2&"',[d_mf810_sc9_3] = '"&d_mf810_sc9_3&"',[d_mf810_sc12_2] = '"&d_mf810_sc12_2&"',[d_all_r] = '"&d_all_r&"',[d_all_sc9_2] = '"&d_all_sc9_2&"',[d_all_sc9_3] = '"&d_all_sc9_3&"',[d_all_sc12_2] = '"&d_all_sc12_2&"',[d_tra_mc_r] = '"&d_tra_mc_r&"',[d_tra_mc_1] = '"&d_tra_mc_1&"',[d_tra_o5_r] = '"&d_tra_o5_r&"',[d_tra_o5_1] = '"&d_tra_o5_1&"',[d_tra_mf86_r] = '"&d_tra_mf86_r&"',[d_tra_mf86_sc9_2] = '"&d_tra_mf86_sc9_2&"',[d_tra_mf86_sc9_3] = '"&d_tra_mf86_sc9_3&"',[d_tra_mf86_sc12_2] = '"&d_tra_mf86_sc12_2&"',[d_tra_mf810_r] = '"&d_tra_mf810_r&"',[d_tra_mf810_sc9_2] = '"&d_tra_mf810_sc9_2&"',[d_tra_mf810_sc9_3] = '"&d_tra_mf810_sc9_3&"',[d_tra_mf810_sc12_2] = '"&d_tra_mf810_sc12_2&"',[d_tra_all_r] = '"&d_tra_all_r&"',[d_tra_all_sc9_2] = '"&d_tra_all_sc9_2&"',[d_tra_all_sc9_3] = '"&d_tra_all_sc9_3&"',[d_tra_all_sc12_2] = '"&d_tra_all_sc12_2&"',[d_msccap_r] = '"&d_msccap_r&"',[d_msccap_m_1] = '"&d_msccap_m_1&"',[d_msccap_sc9r1_1] = '"&d_msccap_sc9r1_1&"',[d_msccap_sc9r2_1] = '"&d_msccap_sc9r2_1&"',[d_rpd_r] = '"&d_rpd_r&"',[d_rpd_1] = '"&d_rpd_1&"',[s_cms_r] = '"&s_cms_r&"',[s_cms_sc9_m_1] = '"&s_cms_sc9_m_1&"',[s_cms_sc9_1] = '"&s_cms_sc9_1&"',[s_cms_sc9_2] = '"&s_cms_sc9_2&"',[s_cms_sc1_1] = '"&s_cms_sc1_1&"',[s_cms_sc1_low_1] = '"&s_cms_sc1_low_1&"',[s_bppc_r] = '"&s_bppc_r&"',[s_bppc_el_1] = '"&s_bppc_el_1&"',[s_bppc_elgs_1] = '"&s_bppc_elgs_1&"',[s_cms_sc9_3] = '"&s_cms_sc9_3&"',[s_cms_sc9_m_3] = '"&s_cms_sc9_m_3&"',[s_cms_sc12_2] = '"&s_cms_sc12_2&"',[e_cesss_1] = '"&e_cesss_1&"',[e_sosc_sc9_1_3] = '"&e_sosc_sc9_1_3&"',[e_sosc_sc9_2] = '"&e_sosc_sc9_2&"',[e_sosc_sc12_2] = '"&e_sosc_sc12_2&"',[d_sosc_sc9_1_3] = '"&d_sosc_sc9_1_3&"',[d_sosc_sc9_2] = '"&d_sosc_sc9_2&"',[d_sosc_sc12_2] = '"&d_sosc_sc12_2&"' output Inserted.rbcid WHERE rbcid ='" &rbcid& "'"
			else
				insertSql = "insert into ratebuildercomponents ( e_cesds_r,e_cesds_1,e_cesss_r,e_cesss_1,e_edc_r,e_edc_sc9_1,e_edc_sc9_2,e_edc_sc9_3,e_edc_sc12_2,e_edc_sc1_f250_1,e_edc_sc1_o250_1,e_tra_r,e_tra_sc9_1,e_tra_sc9_2,e_tra_sc9_3,e_tra_sc12_2,e_tra_sc1_f250_1,e_tra_sc1_o250_1,e_mac_r,e_mac_1,e_sbc_r,e_sbc_1,e_rpsp_r,e_rpsp_1,e_psls_r,e_psls_sc9_1,e_psls_sc12_2,e_psls_sc1_1,e_rdm_r,e_rdm_sc9_1,e_rdm_sc12_2,e_rdm_sc1_1,e_drs_r,e_drs_sc9_1,e_drs_sc12_2,e_drs_sc1_1,e_mfc_r,e_mfc_sc912_1,e_mfc_sc1_1,d_mc_r,d_mc_1,d_o5_r,d_o5_1,d_mf86_r,d_mf86_sc9_2,d_mf86_sc9_3,d_mf86_sc12_2,d_mf810_r,d_mf810_sc9_2,d_mf810_sc9_3,d_mf810_sc12_2,d_all_r,d_all_sc9_2,d_all_sc9_3,d_all_sc12_2,d_tra_mc_r,d_tra_mc_1,d_tra_o5_r,d_tra_o5_1,d_tra_mf86_r,d_tra_mf86_sc9_2,d_tra_mf86_sc9_3,d_tra_mf86_sc12_2,d_tra_mf810_r,d_tra_mf810_sc9_2,d_tra_mf810_sc9_3,d_tra_mf810_sc12_2,d_tra_all_r,d_tra_all_sc9_2,d_tra_all_sc9_3,d_tra_all_sc12_2,d_msccap_r,d_msccap_m_1,d_msccap_sc9r1_1,d_msccap_sc9r2_1,d_rpd_r,d_rpd_1,s_cms_r,s_cms_sc9_m_1,s_cms_sc9_1,s_cms_sc9_2,s_cms_sc9_3,s_cms_sc9_m_3,s_cms_sc12_2,s_cms_sc1_1,s_cms_sc1_low_1,s_bppc_r,s_bppc_el_1,s_bppc_elgs_1, d_dlms_r, d_dlms_sc9_1, d_dlms_sc9_2_3, d_dlms_sc12_2, e_sosc_sc9_1_3, d_sosc_sc9_1_3, e_sosc_sc9_2, d_sosc_sc9_2, e_sosc_sc12_2, d_sosc_sc12_2) output Inserted.rbcid values (" & _
				"'"&e_cesds_r&"','"&e_cesds_1&"','"&e_cesss_r&"','"&e_cesss_1&"','"&e_edc_r&"','"&e_edc_sc9_1&"','"&e_edc_sc9_2&"','"&e_edc_sc9_3&"','"&e_edc_sc12_2&"','"&e_edc_sc1_f250_1&"','"&e_edc_sc1_o250_1&"','"&e_tra_r&"','"&e_tra_sc9_1&"','"&e_tra_sc9_2&"','"&e_tra_sc9_3&"','"&e_tra_sc12_2&"','"&e_tra_sc1_f250_1&"','"&e_tra_sc1_o250_1&"','"&e_mac_r&"','"&e_mac_1&"','"&e_sbc_r&_
				"','"&e_sbc_1&"','"&e_rpsp_r&"','"&e_rpsp_1&"','"&e_psls_r&"','"&e_psls_sc9_1&"','"&e_psls_sc12_2&"','"&e_psls_sc1_1&"','"&e_rdm_r&"','"&e_rdm_sc9_1&"','"&e_rdm_sc12_2&"','"&e_rdm_sc1_1&"','"&e_drs_r&"','"&e_drs_sc9_1&"','"&e_drs_sc1_1&"','"&e_drs_sc12_2&"','"&e_mfc_r&_
				"','"&e_mfc_sc912_1&"','"&e_mfc_sc1_1&"','"&d_mc_r&"','"&d_mc_1&"','"&d_o5_r&"','"&d_o5_1&"','"&d_mf86_r&"','"&d_mf86_sc9_2&"','"&d_mf86_sc9_3&"','"&d_mf86_sc12_2&"','"&d_mf810_r&"','"&d_mf810_sc9_2&"','"&d_mf810_sc9_3&"','"&d_mf810_sc12_2&_
				"','"&d_all_r&"','"&d_all_sc9_2&"','"&d_all_sc9_3&"','"&d_all_sc12_2&"','"&d_tra_mc_r&"','"&d_tra_mc_1&"','"&d_tra_o5_r&"','"&d_tra_o5_1&"','"&d_tra_mf86_r&"','"&d_tra_mf86_sc9_2&"','"&d_tra_mf86_sc9_3&_
				"','"&d_tra_mf86_sc12_2&"','"&d_tra_mf810_r&"','"&d_tra_mf810_sc9_2&"','"&d_tra_mf810_sc9_3&"','"&d_tra_mf810_sc12_2&"','"&d_tra_all_r&"','"&d_tra_all_sc9_2&"','"&d_tra_all_sc9_3&"','"&d_tra_all_sc12_2&_
				"','"&d_msccap_r&"','"&d_msccap_m_1&"','"&d_msccap_sc9r1_1&"','"&d_msccap_sc9r2_1&"','"&d_rpd_r&"','"&d_rpd_1&"','"&s_cms_r&"','"&s_cms_sc9_m_1&"','"&s_cms_sc9_1&"','"&s_cms_sc9_2&"','"&s_cms_sc9_3&"','"&s_cms_sc9_m_3&"','"&s_cms_sc12_2&"','"&s_cms_sc1_1&"','"&s_cms_sc1_low_1&"','"&s_bppc_r&"','"&s_bppc_el_1&"','"&s_bppc_elgs_1&"','"&d_dlms_r&"','"&d_dlms_sc9_1&"','"&d_dlms_sc9_2_3&"','"&d_dlms_sc12_2&"','"&e_sosc_sc9_1_3&"','"&d_sosc_sc9_1_3&"','"&e_sosc_sc9_2&"','"&d_sosc_sc9_2&"','"&e_sosc_sc12_2&"','"&d_sosc_sc12_2&"' "&_
				")"
			end if
			response.write(insertSql) & "</br>"
			'response.end
			cnn1.Execute insertSql
			rst1.Open "Select rbcid from ratebuildercomponents where rbcid = SCOPE_IDENTITY()", cnn1
			if not rst1.eof then rbcid = rst1("rbcid")
			response.write "rbcid: " & rbcid & "</br>"
			rst1.close
			
			'insertSql = "if not exists(select top 1 rbid from ratebuilder where rbcid = '"&rbcid&"') insert into ratebuilder ( rbcid, rateperiod, createdBy, createdOn) output Inserted.rbid values ('"&rbcid&"', datefromparts("&rateperiod&"),'"&createdBy&"', getdate()" &") else update ratebuilder set rateperiod=datefromparts("&rateperiod&"), modifiedOn = getdate(), modifiedBy = '"&modifiedby&"' output Inserted.rbid where rbcid='"&rbcid&"'"
			
			ssql = "select top 1 rbid from ratebuilder where rbcid = '"&rbcid&"'"
			rst1.open ssql, cnn1
			if not rst1.eof then rbid = rst1("rbid")
			rst1.close
			if rbid <> "" then
				insertsql = "update ratebuilder set rateperiod=datefromparts("&rateperiod&"), modifiedOn = getdate(), modifiedBy = '"&modifiedby&"' output Inserted.rbid where rbid='"&rbid&"'"
			else
				insertsql = "insert into ratebuilder ( rbcid, rateperiod, createdBy, createdOn) output Inserted.rbid values ('"&rbcid&"', datefromparts("&rateperiod&"),'"&createdBy&"', getdate()" &")"
			end if
			response.write(insertSql) & "</br>"
			cnn1.Execute insertSql
			rst1.Open "Select rbid from ratebuilder where rbid = SCOPE_IDENTITY()", cnn1
			if not rst1.eof then rbid = rst1("rbid")
			response.write "rbcid:"&rbcid & " | rbrid:"&rbrid & " | rbid:"&rbid&"</br>"
			rst1.Close
			response.write "rbid: " & rbid & "</br>"
			'response.end
			
			Response.redirect "ratevalues.asp?rbid="&rbid
			%>
