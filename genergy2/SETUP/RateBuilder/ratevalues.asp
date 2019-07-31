<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/nonsecure.inc"-->
<%


dim cnn1, rst1, strsql, insertSql, cmd
set cnn1 = server.createobject("ADODB.connection")
set cmd = server.createobject("ADODB.command")
set rst1 = server.createobject("ADODB.recordset")
cnn1.open getConnect(0,0,"dbCore")
%>
<%
	dim rbid,rbcid,rbrid,rateperiod,createdBy,createdOn,modifiedBy,modifiedOn, year, month
	
	dim e_edc_r,e_edc_sc9_1,e_edc_sc9_2,e_edc_sc9_3,e_edc_sc12_2,e_tra_r,e_tra_sc9_1,e_tra_sc9_2,e_tra_sc9_3,e_tra_sc12_2,e_mac_r,e_mac_1,e_sbc_r,e_sbc_1,e_rpsp_r,e_rpsp_1,e_psls_r,e_psls_sc9_1
	dim e_psls_sc12_2,e_rdm_r,e_rdm_sc9_1,e_rdm_sc12_2,e_drs_r,e_drs_sc9_1,e_drs_sc12_2,e_mfc_r,e_mfc_sc912_1,d_mc_r,d_mc_1,d_o5_r,d_o5_1,d_mf86_r,d_mf86_sc9_2,d_mf86_sc9_3,d_mf86_sc12_2,d_mf810_r
	dim d_mf810_sc9_2, d_mf810_sc9_3,d_mf810_sc12_2,d_all_r,d_all_sc9_2,d_all_sc9_3,d_all_sc12_2,d_tra_mc_r,d_tra_mc_1,d_tra_o5_r,d_tra_o5_1,d_tra_mf86_r,d_tra_mf86_sc9_2,d_tra_mf86_sc9_3
	dim d_tra_mf86_sc12_2,d_tra_mf810_r,d_tra_mf810_sc9_2,d_tra_mf810_sc9_3,d_tra_mf810_sc12_2,d_tra_all_r,d_tra_all_sc9_2,d_tra_all_sc9_3,d_tra_all_sc12_2,d_msccap_r,d_msccap_sc9r1_1,d_msccap_sc9r2_1,d_msccap_m_1
	dim s_cms_r,s_cms_sc9_1,s_cms_sc9_m_1,s_cms_sc9_2,s_cms_sc1_1,s_cms_sc1_low_1,s_bppc_r,s_bppc_el_1,s_bppc_elgs_1, e_cesds_r, e_cesds_1, e_cesss_r, e_cesss_1, s_cms_sc9_3, s_cms_sc9_m_3, s_cms_sc12_2, d_dlms_sc9_1,d_dlms_sc9_2_3,d_dlms_sc12_2, d_rpd_1
	dim e_sosc_sc9_1_3, d_sosc_sc9_1_3, e_sosc_sc9_2, d_sosc_sc9_2, e_sosc_sc12_2, d_sosc_sc12_2
	dim sc9r1_d_dlms,sc9r1_e_er, sc9r1_e_macadj, sc9r1_d_mscadj, sc9r1_d_dr_l05, sc9r1_d_dr_l100, sc9r1_d_dr_l999, conedsc9r1_s_bppc, conedsc9r1_s_cmc, conedsc9r1_e_er, conedsc9r1_e_macadj, conedsc9r1_e_mfc, conedsc9r1_d_mscadj, conedsc9r1_d_dr_l05, conedsc9r1_d_dr_l999, conedsc9r1_e_cesss, sc9r2_e_er_p759 ,sc9r2_e_er_p1800,sc9r2_e_er_p2200,sc9r2_e_er_p2359,sc9r2_e_er,sc9r2_e_macadj_p759 ,sc9r2_e_macadj_p1800,sc9r2_e_macadj_p2200,sc9r2_e_macadj_p2359,sc9r2_e_macadj,sc9r2_d_mscadj_p2200, sc9r2_d_dr_p1800, sc9r2_d_dr_p2200, sc9r2_d_dr_p2359, sc9ra1_e_er, sc9ra1_e_macadj, sc9ra1_d_dr_l05, sc9ra1_d_dr_l100, sc9ra1_d_dr_l999,sc9ra2_e_er_p759 ,sc9ra2_e_er_p1800,sc9ra2_e_er_p2200,sc9ra2_e_er_p2359,sc9ra2_e_er,sc9ra2_e_macadj_p759 ,sc9ra2_e_macadj_p1800,sc9ra2_e_macadj_p2200,sc9ra2_e_macadj_p2359,sc9ra2_e_macadj,sc9ra2_d_dr_p1800, sc9ra2_d_dr_p2200, sc9ra2_d_dr_p2359,sc9ra3_e_er_p759 ,sc9ra3_e_er_p1800,sc9ra3_e_er_p2200,sc9ra3_e_er_p2359,sc9ra3_e_er,sc9ra3_e_macadj_p759 ,sc9ra3_e_macadj_p1800,sc9ra3_e_macadj_p2200,sc9ra3_e_macadj_p2359,sc9ra3_e_macadj,sc9ra3_d_dr_p1800, sc9ra3_d_dr_p2200, sc9ra3_d_dr_p2359,sc12ra2_e_er_p759 ,sc12ra2_e_er_p1800,sc12ra2_e_er_p2200,sc12ra2_e_er_p2359,sc12ra2_e_er,sc12ra2_e_macadj_p759 ,sc12ra2_e_macadj_p1800,sc12ra2_e_macadj_p2200,sc12ra2_e_macadj_p2359,sc12ra2_e_macadj,sc12ra2_d_dr_p1800, sc12ra2_d_dr_p2200, sc12ra2_d_dr_p2359, sc9ra1_d_dlms, conedsc9r1_d_dlms
	
	dim conedsc9r1m_s_bppc, conedsc9r1m_s_cmc, conedsc9r1m_e_er, conedsc9r1m_e_macadj, conedsc9r1m_e_mfc, conedsc9r1m_d_mscadj, conedsc9r1m_d_dr_l05, conedsc9r1m_d_dr_l999, conedsc9r1m_e_cesss, conedsc9r1m_d_dlms, conedsc9ra1m_d_msccap
	
	dim conedsc9ra1_s_bppc, conedsc9ra1_s_cmc, conedsc9ra1_e_er, conedsc9ra1_e_macadj, conedsc9ra1_e_mfc, conedsc9ra1_d_mscadj, conedsc9ra1_d_dr_l05, conedsc9ra1_d_dr_l999, conedsc9ra1_e_cesss, conedsc9ra1_d_dlms
	
	dim conedsc9ra1m_s_bppc, conedsc9ra1m_s_cmc, conedsc9ra1m_e_er, conedsc9ra1m_e_macadj, conedsc9ra1m_e_mfc, conedsc9ra1m_d_mscadj, conedsc9ra1m_d_dr_l05, conedsc9ra1m_d_dr_l999, conedsc9ra1m_e_cesss, conedsc9r2_s_bppc, conedsc9r2_s_cmc, conedsc9r2_e_er, conedsc9r2_e_macadj, conedsc9r2_e_mfc, conedsc9r2_d_mscadj,  conedsc9r2_e_cesss, conedsc9r2_d_dr_p1800, conedsc9r2_d_dr_p2200, conedsc9r2_d_dr_p2359, conedsc9r2_d_mscadj_p2200, conedsc9r2_e_macadj_p759, conedsc9r2_e_macadj_p1800, conedsc9r2_e_macadj_p2200, conedsc9r2_e_macadj_p2359, conedsc9r2_e_er_p759, conedsc9r2_e_er_p1800, conedsc9r2_e_er_p2200, conedsc9r2_e_er_p2359, conedsc9ra1m_d_dlms
	
	dim conedsc9r2m_s_bppc, conedsc9r2m_s_cmc, conedsc9r2m_e_er, conedsc9r2m_e_macadj, conedsc9r2m_e_mfc, conedsc9r2m_d_mscadj,  conedsc9r2m_e_cesss, conedsc9r2m_d_dr_p1800, conedsc9r2m_d_dr_p2200, conedsc9r2m_d_dr_p2359, conedsc9r2m_d_mscadj_p2200, conedsc9r2m_e_macadj_p759, conedsc9r2m_e_macadj_p1800, conedsc9r2m_e_macadj_p2200, conedsc9r2m_e_macadj_p2359, conedsc9r2m_e_er_p759, conedsc9r2m_e_er_p1800, conedsc9r2m_e_er_p2200, conedsc9r2m_e_er_p2359
	
	dim conedsc9ra2_s_bppc, conedsc9ra2_s_cmc, conedsc9ra2_e_er, conedsc9ra2_e_macadj, conedsc9ra2_e_mfc, conedsc9ra2_d_mscadj,  conedsc9ra2_e_cesss, conedsc9ra2_d_dr_p1800, conedsc9ra2_d_dr_p2200, conedsc9ra2_d_dr_p2359, conedsc9ra2_d_mscadj_p2200, conedsc9ra2_e_macadj_p759, conedsc9ra2_e_macadj_p1800, conedsc9ra2_e_macadj_p2200, conedsc9ra2_e_macadj_p2359, conedsc9ra2_e_er_p759, conedsc9ra2_e_er_p1800, conedsc9ra2_e_er_p2200, conedsc9ra2_e_er_p2359
	
	dim conedsc9ra3_s_bppc, conedsc9ra3_s_cmc, conedsc9ra3_e_er, conedsc9ra3_e_macadj, conedsc9ra3_e_mfc, conedsc9ra3_d_mscadj,  conedsc9ra3_e_cesss, conedsc9ra3_d_dr_p1800, conedsc9ra3_d_dr_p2200, conedsc9ra3_d_dr_p2359, conedsc9ra3_d_mscadj_p2200, conedsc9ra3_e_macadj_p759, conedsc9ra3_e_macadj_p1800, conedsc9ra3_e_macadj_p2200, conedsc9ra3_e_macadj_p2359, conedsc9ra3_e_er_p759, conedsc9ra3_e_er_p1800, conedsc9ra3_e_er_p2200, conedsc9ra3_e_er_p2359
	
	dim conedsc9ra3m_s_bppc, conedsc9ra3m_s_cmc, conedsc9ra3m_e_er, conedsc9ra3m_e_macadj, conedsc9ra3m_e_mfc, conedsc9ra3m_d_mscadj,  conedsc9ra3m_e_cesss, conedsc9ra3m_d_dr_p1800, conedsc9ra3m_d_dr_p2200, conedsc9ra3m_d_dr_p2359, conedsc9ra3m_d_mscadj_p2200, conedsc9ra3m_e_macadj_p759, conedsc9ra3m_e_macadj_p1800, conedsc9ra3m_e_macadj_p2200, conedsc9ra3m_e_macadj_p2359, conedsc9ra3m_e_er_p759, conedsc9ra3m_e_er_p1800, conedsc9ra3m_e_er_p2200, conedsc9ra3m_e_er_p2359
	
	dim conedsc12ra2_s_bppc, conedsc12ra2_s_cmc, conedsc12ra2_e_er, conedsc12ra2_e_macadj, conedsc12ra2_e_mfc, conedsc12ra2_d_mscadj,  conedsc12ra2_e_cesss, conedsc12ra2_d_dr_p1800, conedsc12ra2_d_dr_p2200, conedsc12ra2_d_dr_p2359, conedsc12ra2_d_mscadj_p2200, conedsc12ra2_e_macadj_p759, conedsc12ra2_e_macadj_p1800, conedsc12ra2_e_macadj_p2200, conedsc12ra2_e_macadj_p2359, conedsc12ra2_e_er_p759, conedsc12ra2_e_er_p1800, conedsc12ra2_e_er_p2200, conedsc12ra2_e_er_p2359
	
	dim conedsc12ra2_e_sbc_p1800,  conedsc12ra2_e_sbc_p2200,  conedsc12ra2_e_sbc_p2359,  conedsc9r2_e_sbc_p1800,  conedsc9r2_e_sbc_p2200,  conedsc9r2_e_sbc_p2359,  conedsc9r2_d_msccap,  conedsc9r1_d_msccap,  conedsc9r1_e_sbc,  conedsc9r1m_e_sbc,  conedsc9r1m_d_msccap,  conedsc9r2m_e_sbc_p1800,  conedsc9r2m_e_sbc_p2200,  conedsc9r2m_e_sbc_p2359,  conedsc9r2m_d_msccap_p2200,  conedsc9ra1_e_sbc,  conedsc9ra1m_e_sbc,  conedsc9ra2_e_sbc_p1800,  conedsc9ra2_e_sbc_p2200,  conedsc9ra2_e_sbc_p2359,  conedsc9ra3_e_sbc_p1800,  conedsc9ra3_e_sbc_p2200,  conedsc9ra3_e_sbc_p2359,  conedsc9ra3m_e_sbc_p1800,  conedsc9ra3m_e_sbc_p2200,  conedsc9ra3m_e_sbc_p2359
	
	dim conedsc9r1_e_tsc, conedsc9r1_d_tsc,  conedsc9r1m_e_tsc, conedsc9r1m_d_tsc, conedsc9ra1_e_tsc, conedsc9ra1_d_tsc, conedsc9ra1m_e_tsc, conedsc9ra1m_d_tsc, conedsc9r2_e_tsc, conedsc9r2_d_tsc, conedsc9r2m_e_tsc, conedsc9r2m_d_tsc, conedsc9ra2_e_tsc, conedsc9ra2_d_tsc, conedsc9ra3_e_tsc, conedsc9ra3_d_tsc, conedsc9ra3m_e_tsc, conedsc9ra3m_d_tsc, conedsc12ra2_e_tsc, conedsc12ra2_d_tsc
	
	rbid = trim(secureRequest("rbid"))
	strsql ="select * from ratebuilder where rbid='" & rbid &"'"
	rst1.Open strsql, cnn1
	rbcid = rst1("rbcid")
	rateperiod = rst1("rateperiod")
	year = datepart("yyyy", rateperiod)
	month = datepart("m", rateperiod)
	rst1.close
	
	strsql ="select * from ratebuildercomponents where rbcid='" & rbcid &"'"'strsql ="select * from ratebuildercomponents rbc join ratebuilder rb on rb.rbcid = rbc.rbcid where rb.rbid ="&rbid
	rst1.Open strsql, cnn1
	if not rst1.eof then 
		e_edc_r = (rst1("e_edc_r"))
		e_edc_sc9_1 = (rst1("e_edc_sc9_1"))
		e_edc_sc9_2 = (rst1("e_edc_sc9_2"))
		e_edc_sc9_3 = (rst1("e_edc_sc9_3"))
		e_edc_sc12_2 = (rst1("e_edc_sc12_2"))
		e_tra_r = (rst1("e_tra_r"))
		e_tra_sc9_1 = (rst1("e_tra_sc9_1"))
		e_tra_sc9_2 = (rst1("e_tra_sc9_2"))
		e_tra_sc9_3 = (rst1("e_tra_sc9_3"))
		e_tra_sc12_2 = (rst1("e_tra_sc12_2"))
		e_mac_r = (rst1("e_mac_r"))
		e_mac_1 = (rst1("e_mac_1"))
		e_sbc_r = (rst1("e_sbc_r"))
		e_sbc_1 = (rst1("e_sbc_1"))
		e_rpsp_r = (rst1("e_rpsp_r"))
		e_rpsp_1 = (rst1("e_rpsp_1"))
		e_psls_r = (rst1("e_psls_r"))
		e_psls_sc9_1 = (rst1("e_psls_sc9_1"))
		e_psls_sc12_2 = (rst1("e_psls_sc12_2"))	
		e_rdm_r = (rst1("e_rdm_r"))
		e_rdm_sc9_1 = (rst1("e_rdm_sc9_1"))
		e_rdm_sc12_2 = (rst1("e_rdm_sc12_2"))
		e_drs_r = (rst1("e_drs_r"))
		e_drs_sc9_1 = (rst1("e_drs_sc9_1"))
		e_drs_sc12_2 = (rst1("e_drs_sc12_2"))
		e_mfc_r = (rst1("e_mfc_r"))
		e_mfc_sc912_1 = (rst1("e_mfc_sc912_1"))
		e_cesds_r = (rst1("e_cesds_r"))
		e_cesds_1 = (rst1("e_cesds_1"))
		e_cesss_r = (rst1("e_cesss_r"))
		e_cesss_1 = (rst1("e_cesss_1"))
		e_sosc_sc9_1_3 = rst1("e_sosc_sc9_1_3")
		e_sosc_sc9_2 = rst1("e_sosc_sc9_2")
		e_sosc_sc12_2 = rst1("e_sosc_sc12_2")
			
		d_mc_r = (rst1("d_mc_r"))
		d_mc_1 = (rst1("d_mc_1"))
		d_o5_r = (rst1("d_o5_r"))
		d_o5_1 = (rst1("d_o5_1"))
		d_mf86_r = (rst1("d_mf86_r"))
		d_mf86_sc9_2 = (rst1("d_mf86_sc9_2"))
		d_mf86_sc9_3 = (rst1("d_mf86_sc9_3"))
		d_mf86_sc12_2 = (rst1("d_mf86_sc12_2"))
		d_mf810_r = (rst1("d_mf810_r"))
		d_mf810_sc9_2 = (rst1("d_mf810_sc9_2"))
		d_mf810_sc9_3 = (rst1("d_mf810_sc9_3"))
		d_mf810_sc12_2 = (rst1("d_mf810_sc12_2"))
		d_all_r = (rst1("d_all_r"))
		d_all_sc9_2 = (rst1("d_all_sc9_2"))
		d_all_sc9_3 = (rst1("d_all_sc9_3"))
		d_all_sc12_2 = (rst1("d_all_sc12_2"))
		d_tra_mc_r = (rst1("d_tra_mc_r"))
		d_tra_mc_1 = (rst1("d_tra_mc_1"))
		d_tra_o5_r = (rst1("d_tra_o5_r"))
		d_tra_o5_1 = (rst1("d_tra_o5_1"))
		d_tra_mf86_r = (rst1("d_tra_mf86_r"))
		d_tra_mf86_sc9_2 = (rst1("d_tra_mf86_sc9_2"))
		d_tra_mf86_sc9_3 = (rst1("d_tra_mf86_sc9_3"))
		d_tra_mf86_sc12_2 = (rst1("d_tra_mf86_sc12_2"))
		d_tra_mf810_r = (rst1("d_tra_mf810_r"))
		d_tra_mf810_sc9_2 = (rst1("d_tra_mf810_sc9_2"))
		d_tra_mf810_sc9_3 = (rst1("d_tra_mf810_sc9_3"))
		d_tra_mf810_sc12_2 = (rst1("d_tra_mf810_sc12_2"))
		d_tra_all_r = (rst1("d_tra_all_r"))
		d_tra_all_sc9_2 = (rst1("d_tra_all_sc9_2"))
		d_tra_all_sc9_3 = (rst1("d_tra_all_sc9_3"))
		d_tra_all_sc12_2 = (rst1("d_tra_all_sc12_2"))
		d_msccap_r = (rst1("d_msccap_r"))
		d_msccap_m_1 = (rst1("d_msccap_m_1"))
		d_msccap_sc9r1_1 = (rst1("d_msccap_sc9r1_1"))
		d_msccap_sc9r2_1 = (rst1("d_msccap_sc9r2_1"))
		d_sosc_sc9_1_3 = rst1("d_sosc_sc9_1_3")
		d_sosc_sc9_2 = rst1("d_sosc_sc9_2")
		d_sosc_sc12_2 = rst1("d_sosc_sc12_2")		
		
		d_dlms_sc9_1 = (rst1("d_dlms_sc9_1"))
		d_dlms_sc9_2_3 = (rst1("d_dlms_sc9_2_3"))
		d_dlms_sc12_2 = (rst1("d_dlms_sc12_2"))
		d_rpd_1 = (rst1("d_rpd_1"))
		
		s_cms_r = (rst1("s_cms_r"))
		s_cms_sc9_1 = (rst1("s_cms_sc9_1"))
		s_cms_sc9_m_1	=(rst1("s_cms_sc9_m_1"))
		s_cms_sc9_2		=(rst1("s_cms_sc9_2"))
		s_cms_sc9_3		=(rst1("s_cms_sc9_3"))
		s_cms_sc9_m_3	=(rst1("s_cms_sc9_m_3"))
		s_cms_sc12_2	=(rst1("s_cms_sc12_2"))
		
		s_bppc_r = (rst1("s_bppc_r"))
		s_bppc_el_1 = (rst1("s_bppc_el_1"))
		
		sc9r1_e_er = ( e_edc_sc9_1 + e_tra_sc9_1 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_mfc_sc912_1 + e_cesds_1 + e_cesss_1 + e_sosc_sc9_1_3 ) / 100
		sc9r1_e_macadj = e_mac_1 / 100
		sc9r1_d_mscadj = d_msccap_sc9r1_1
		sc9r1_d_dr_l05 = d_mc_1 + d_tra_mc_1
		sc9r1_d_dr_l100 = d_o5_1 + d_tra_o5_1
		sc9r1_d_dr_l999 = d_o5_1 + d_tra_o5_1
		sc9r1_d_dlms = d_dlms_sc9_1 + d_sosc_sc9_1_3
		
		conedsc9r1_s_bppc = s_bppc_el_1
		conedsc9r1_s_cmc = s_cms_sc9_1
		conedsc9r1_e_er = ( e_edc_sc9_1 + e_mac_1 + e_rdm_sc9_1 + e_cesds_1 + e_drs_sc9_1 ) / 100
		'conedsc9r1_e_macadj = e_mac_1 / 100
		conedsc9r1_e_sbc = e_sbc_1 / 100
		conedsc9r1_e_mfc = e_mfc_sc912_1 / 100
		conedsc9r1_e_cesss = e_cesss_1 / 100
		conedsc9r1_d_msccap = d_msccap_sc9r1_1
		conedsc9r1_d_dr_l05 = d_mc_1 '+ d_tra_mc_1
		conedsc9r1_d_dr_l999 = d_o5_1 '+ d_tra_o5_1
		conedsc9r1_d_dlms = d_dlms_sc9_1
		conedsc9r1_e_tsc = e_sosc_sc9_1_3 / 100		
		conedsc9r1_d_tsc = d_sosc_sc9_1_3

		conedsc9ra1_s_bppc = s_bppc_el_1
		conedsc9ra1_s_cmc = s_cms_sc9_1
		conedsc9ra1_e_er = ( e_edc_sc9_1 + e_mac_1 + e_rdm_sc9_1 + e_cesds_1 + e_drs_sc9_1 ) / 100
		'conedsc9ra1_e_macadj = e_mac_1 / 100
		conedsc9ra1_e_sbc = e_sbc_1 / 100
		conedsc9ra1_d_dr_l05 = d_mc_1 '+ d_tra_mc_1
		conedsc9ra1_d_dr_l999 = d_o5_1 '+ d_tra_o5_1
		conedsc9ra1_d_dlms = d_dlms_sc9_1
		conedsc9ra1_e_tsc = e_sosc_sc9_1_3 / 100		
		conedsc9ra1_d_tsc = d_sosc_sc9_1_3
		
		conedsc9r1m_s_bppc = s_bppc_el_1
		conedsc9r1m_s_cmc = s_cms_sc9_m_1
		conedsc9r1m_e_er = ( e_edc_sc9_1 + e_tra_sc9_1 + e_mac_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
		conedsc9r1m_e_macadj = e_mac_1 / 100
		conedsc9r1m_e_mfc = e_mfc_sc912_1 / 100
		conedsc9r1m_e_cesss = e_cesss_1 / 100
		conedsc9r1m_d_msccap = d_msccap_m_1
		conedsc9r1m_d_dr_l05 = d_mc_1 + d_tra_mc_1
		conedsc9r1m_d_dr_l999 = d_o5_1 + d_tra_o5_1
		conedsc9r1m_d_dlms = d_dlms_sc9_1
		conedsc9r1m_e_sbc = e_sbc_1 / 100
		conedsc9r1m_e_tsc = e_sosc_sc9_1_3 / 100		
		conedsc9r1m_d_tsc = d_sosc_sc9_1_3
		

		conedsc9ra1m_s_bppc = s_bppc_el_1
		conedsc9ra1m_s_cmc = s_cms_sc9_m_1
		conedsc9ra1m_e_er = ( e_edc_sc9_1 + e_tra_sc9_1 + e_mac_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
		conedsc9ra1m_e_macadj = e_mac_1 / 100
		conedsc9ra1m_d_dr_l05 = d_mc_1 + d_tra_mc_1
		conedsc9ra1m_d_dr_l999 = d_o5_1 + d_tra_o5_1
		conedsc9ra1m_d_dlms = d_dlms_sc9_1
		conedsc9ra1m_e_sbc = e_sbc_1 / 100
		conedsc9ra1m_e_tsc = e_sosc_sc9_1_3 / 100		
		conedsc9ra1m_d_tsc = d_sosc_sc9_1_3
		
		conedsc9r2_s_bppc = s_bppc_el_1
		conedsc9r2_s_cmc = s_cms_sc9_2
		conedsc9r2_e_er = ( e_edc_sc9_2 + e_tra_sc9_2 + e_mac_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
		conedsc9r2_e_er_p759  = conedsc9r2_e_er
		conedsc9r2_e_er_p1800 = conedsc9r2_e_er
		conedsc9r2_e_er_p2200 = conedsc9r2_e_er
		conedsc9r2_e_er_p2359 = conedsc9r2_e_er
		conedsc9r2_e_macadj = e_mac_1 / 100
		conedsc9r2_e_macadj_p759  = conedsc9r2_e_macadj
		conedsc9r2_e_macadj_p1800 = conedsc9r2_e_macadj
		conedsc9r2_e_macadj_p2200 = conedsc9r2_e_macadj
		conedsc9r2_e_macadj_p2359 = conedsc9r2_e_macadj		
		conedsc9r2_e_mfc = e_mfc_sc912_1 / 100
		conedsc9r2_e_cesss = e_cesss_1 / 100
		conedsc9r2_d_mscadj_p2200 = d_msccap_sc9r2_1
		conedsc9r2_d_msccap = d_msccap_sc9r2_1
		conedsc9r2_d_dr_p1800 = d_mf86_sc9_2 + d_tra_mf86_sc9_2
		conedsc9r2_d_dr_p2200 = d_mf810_sc9_2 + d_tra_mf810_sc9_2 + d_dlms_sc9_2_3
		conedsc9r2_d_dr_p2359 = d_all_sc9_2 + d_tra_all_sc9_2
		conedsc9r2_e_sbc_p1800 = e_sbc_1 / 100
		conedsc9r2_e_sbc_p2200 = e_sbc_1 / 100
		conedsc9r2_e_sbc_p2359 = e_sbc_1 / 100
		conedsc9r2_e_tsc = e_sosc_sc9_2 / 100
		conedsc9r2_d_tsc = d_sosc_sc9_2
		
		conedsc9r2m_s_bppc = s_bppc_el_1
		conedsc9r2m_s_cmc = s_cms_sc9_2
		conedsc9r2m_e_er = ( e_edc_sc9_2 + e_tra_sc9_2 + e_mac_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
		conedsc9r2m_e_er_p759  = conedsc9r2m_e_er
		conedsc9r2m_e_er_p1800 = conedsc9r2m_e_er
		conedsc9r2m_e_er_p2200 = conedsc9r2m_e_er
		conedsc9r2m_e_er_p2359 = conedsc9r2m_e_er
		conedsc9r2m_e_macadj = e_mac_1 / 100
		conedsc9r2m_e_macadj_p759  = conedsc9r2m_e_macadj
		conedsc9r2m_e_macadj_p1800 = conedsc9r2m_e_macadj
		conedsc9r2m_e_macadj_p2200 = conedsc9r2m_e_macadj
		conedsc9r2m_e_macadj_p2359 = conedsc9r2m_e_macadj		
		conedsc9r2m_e_mfc = e_mfc_sc912_1 / 100
		conedsc9r2m_e_cesss = e_cesss_1 / 100
		conedsc9r2m_d_mscadj_p2200 = d_msccap_m_1
		conedsc9r2m_d_msccap_p2200 = d_msccap_m_1
		conedsc9r2m_d_dr_p1800 = d_mf86_sc9_2 + d_tra_mf86_sc9_2
		conedsc9r2m_d_dr_p2200 = d_mf810_sc9_2 + d_tra_mf810_sc9_2 + d_dlms_sc9_2_3
		conedsc9r2m_d_dr_p2359 = d_all_sc9_2 + d_tra_all_sc9_2
		conedsc9r2m_e_sbc_p1800 = e_sbc_1 / 100
		conedsc9r2m_e_sbc_p2200 = e_sbc_1 / 100
		conedsc9r2m_e_sbc_p2359 = e_sbc_1 / 100
		conedsc9r2m_e_tsc = e_sosc_sc9_2 / 100
	    conedsc9r2m_d_tsc = d_sosc_sc9_2
		
		conedsc9ra2_s_bppc = s_bppc_el_1
		conedsc9ra2_s_cmc = s_cms_sc9_2
		conedsc9ra2_e_er = ( e_edc_sc9_2 + e_tra_sc9_3 + e_mac_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
		conedsc9ra2_e_er_p759  = conedsc9ra2_e_er
		conedsc9ra2_e_er_p1800 = conedsc9ra2_e_er
		conedsc9ra2_e_er_p2200 = conedsc9ra2_e_er
		conedsc9ra2_e_er_p2359 = conedsc9ra2_e_er
		conedsc9ra2_e_macadj = e_mac_1 / 100
		conedsc9ra2_e_macadj_p759  = conedsc9ra2_e_macadj
		conedsc9ra2_e_macadj_p1800 = conedsc9ra2_e_macadj
		conedsc9ra2_e_macadj_p2200 = conedsc9ra2_e_macadj
		conedsc9ra2_e_macadj_p2359 = conedsc9ra2_e_macadj		
		conedsc9ra2_d_dr_p1800 = d_mf86_sc9_2 + d_tra_mf86_sc9_2
		conedsc9ra2_d_dr_p2200 = d_mf810_sc9_2 + d_tra_mf810_sc9_2 + d_dlms_sc9_2_3
		conedsc9ra2_d_dr_p2359 = d_all_sc9_2 + d_tra_all_sc9_2
		conedsc9ra2_e_sbc_p1800 = e_sbc_1 / 100
		conedsc9ra2_e_sbc_p2200 = e_sbc_1 / 100
		conedsc9ra2_e_sbc_p2359 = e_sbc_1 / 100
		conedsc9ra2_e_tsc = e_sosc_sc9_2 / 100
        conedsc9ra2_d_tsc = d_sosc_sc9_2
		
		conedsc9ra3_s_bppc = s_bppc_el_1
		conedsc9ra3_s_cmc = s_cms_sc9_3
		conedsc9ra3_e_er = ( e_edc_sc9_3 + e_tra_sc9_3 + e_mac_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
		conedsc9ra3_e_er_p759  = conedsc9ra3_e_er
		conedsc9ra3_e_er_p1800 = conedsc9ra3_e_er
		conedsc9ra3_e_er_p2200 = conedsc9ra3_e_er
		conedsc9ra3_e_er_p2359 = conedsc9ra3_e_er
		conedsc9ra3_e_macadj = e_mac_1 / 100
		conedsc9ra3_e_macadj_p759  = conedsc9ra3_e_macadj
		conedsc9ra3_e_macadj_p1800 = conedsc9ra3_e_macadj
		conedsc9ra3_e_macadj_p2200 = conedsc9ra3_e_macadj
		conedsc9ra3_e_macadj_p2359 = conedsc9ra3_e_macadj		
		conedsc9ra3_d_dr_p1800 = d_mf86_sc9_3 + d_tra_mf86_sc9_3
		conedsc9ra3_d_dr_p2200 = d_mf810_sc9_3 + d_tra_mf810_sc9_3 + d_dlms_sc9_2_3
		conedsc9ra3_d_dr_p2359 = d_all_sc9_3 + d_tra_all_sc9_3
		conedsc9ra3_e_sbc_p1800 = e_sbc_1 / 100
		conedsc9ra3_e_sbc_p2200 = e_sbc_1 / 100
		conedsc9ra3_e_sbc_p2359 = e_sbc_1 / 100
		conedsc9ra3_e_tsc = e_sosc_sc9_1_3 / 100		
		conedsc9ra3_d_tsc = d_sosc_sc9_1_3

		conedsc9ra3m_s_bppc = s_bppc_el_1
		conedsc9ra3m_s_cmc = s_cms_sc9_m_3
		conedsc9ra3m_e_er = ( e_edc_sc9_3 + e_tra_sc9_3 + e_mac_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
		conedsc9ra3m_e_er_p759  = conedsc9ra3m_e_er
		conedsc9ra3m_e_er_p1800 = conedsc9ra3m_e_er
		conedsc9ra3m_e_er_p2200 = conedsc9ra3m_e_er
		conedsc9ra3m_e_er_p2359 = conedsc9ra3m_e_er
		conedsc9ra3m_e_macadj = e_mac_1 / 100
		conedsc9ra3m_e_macadj_p759  = conedsc9ra3m_e_macadj
		conedsc9ra3m_e_macadj_p1800 = conedsc9ra3m_e_macadj
		conedsc9ra3m_e_macadj_p2200 = conedsc9ra3m_e_macadj
		conedsc9ra3m_e_macadj_p2359 = conedsc9ra3m_e_macadj		
		conedsc9ra3m_d_dr_p1800 = d_mf86_sc9_3 + d_tra_mf86_sc9_3
		conedsc9ra3m_d_dr_p2200 = d_mf810_sc9_3 + d_tra_mf810_sc9_3 + d_dlms_sc9_2_3
		conedsc9ra3m_d_dr_p2359 = d_all_sc9_3 + d_tra_all_sc9_3
		conedsc9ra3m_e_sbc_p1800 = e_sbc_1 / 100
		conedsc9ra3m_e_sbc_p2200 = e_sbc_1 / 100
		conedsc9ra3m_e_sbc_p2359 = e_sbc_1 / 100
		conedsc9ra3m_e_tsc = e_sosc_sc9_1_3 / 100		
		conedsc9ra3m_d_tsc = d_sosc_sc9_1_3

		conedsc12ra2_s_bppc = s_bppc_el_1
		conedsc12ra2_s_cmc = s_cms_sc12_2
		conedsc12ra2_e_er = ( e_edc_sc12_2 + e_tra_sc12_2 + e_mac_1 + e_rpsp_1 + e_psls_sc12_2 + e_rdm_sc12_2 + e_drs_sc12_2 + e_cesds_1) / 100
		conedsc12ra2_e_er_p759  = conedsc12ra2_e_er
		conedsc12ra2_e_er_p1800 = conedsc12ra2_e_er
		conedsc12ra2_e_er_p2200 = conedsc12ra2_e_er
		conedsc12ra2_e_er_p2359 = conedsc12ra2_e_er
		conedsc12ra2_e_macadj = e_mac_1 / 100
		conedsc12ra2_e_macadj_p759  = conedsc12ra2_e_macadj
		conedsc12ra2_e_macadj_p1800 = conedsc12ra2_e_macadj
		conedsc12ra2_e_macadj_p2200 = conedsc12ra2_e_macadj
		conedsc12ra2_e_macadj_p2359 = conedsc12ra2_e_macadj		
		conedsc12ra2_d_dr_p1800 = d_mf86_sc12_2 + d_tra_mf86_sc12_2
		conedsc12ra2_d_dr_p2200 = d_mf810_sc12_2 + d_tra_mf810_sc12_2 + d_dlms_sc12_2
		conedsc12ra2_d_dr_p2359 = d_all_sc12_2 + d_tra_all_sc12_2
		conedsc12ra2_e_sbc_p1800 = e_sbc_1 / 100
		conedsc12ra2_e_sbc_p2200 = e_sbc_1 / 100
		conedsc12ra2_e_sbc_p2359 = e_sbc_1 / 100
		conedsc12ra2_e_tsc = e_sosc_sc12_2 / 100		
		conedsc12ra2_d_tsc = d_sosc_sc12_2
		
		sc9r2_e_er = ( e_edc_sc9_2 + e_tra_sc9_2 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_mfc_sc912_1 + e_cesds_1 + e_cesss_1 + e_sosc_sc9_2 ) / 100
		sc9r2_e_er_p759  = sc9r2_e_er
		sc9r2_e_er_p1800 = sc9r2_e_er
		sc9r2_e_er_p2200 = sc9r2_e_er
		sc9r2_e_er_p2359 = sc9r2_e_er
		sc9r2_e_macadj = e_mac_1 / 100
		sc9r2_e_macadj_p759  = sc9r2_e_macadj
		sc9r2_e_macadj_p1800 = sc9r2_e_macadj
		sc9r2_e_macadj_p2200 = sc9r2_e_macadj
		sc9r2_e_macadj_p2359 = sc9r2_e_macadj
		sc9r2_d_mscadj_p2200 = d_msccap_sc9r2_1
		sc9r2_d_dr_p1800 = d_mf86_sc9_2 + d_tra_mf86_sc9_2
		sc9r2_d_dr_p2200 = d_mf810_sc9_2 + d_tra_mf810_sc9_2 + d_dlms_sc9_2_3 + d_sosc_sc9_2
		sc9r2_d_dr_p2359 = d_all_sc9_2 + d_tra_all_sc9_2
		
		sc9ra1_e_er = ( e_edc_sc9_1 + e_tra_sc9_1 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1 + e_sosc_sc9_1_3 ) / 100
		sc9ra1_e_macadj = e_mac_1 / 100
		sc9ra1_d_dr_l05 = d_mc_1 + d_tra_mc_1
		sc9ra1_d_dr_l100 = d_o5_1 + d_tra_o5_1
		sc9ra1_d_dr_l999 = d_o5_1 + d_tra_o5_1
		sc9ra1_d_dlms = d_dlms_sc9_1 + d_sosc_sc9_1_3
		
		sc9ra2_e_er = ( e_edc_sc9_2 + e_tra_sc9_2 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1 + e_sosc_sc9_2 ) / 100
		sc9ra2_e_er_p759  = sc9ra2_e_er
		sc9ra2_e_er_p1800 = sc9ra2_e_er
		sc9ra2_e_er_p2200 = sc9ra2_e_er
		sc9ra2_e_er_p2359 = sc9ra2_e_er
		sc9ra2_e_macadj = e_mac_1 / 100
		sc9ra2_e_macadj_p759  = sc9ra2_e_macadj
		sc9ra2_e_macadj_p1800 = sc9ra2_e_macadj
		sc9ra2_e_macadj_p2200 = sc9ra2_e_macadj
		sc9ra2_e_macadj_p2359 = sc9ra2_e_macadj
		sc9ra2_d_dr_p1800 = d_mf86_sc9_2 + d_tra_mf86_sc9_2
		sc9ra2_d_dr_p2200 = d_mf810_sc9_2 + d_tra_mf810_sc9_2 + d_dlms_sc9_2_3 + d_sosc_sc9_2
		sc9ra2_d_dr_p2359 = d_all_sc9_2 + d_tra_all_sc9_2
		
		sc9ra3_e_er = ( e_edc_sc9_3 + e_tra_sc9_3 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1 + e_sosc_sc9_1_3 ) / 100
		sc9ra3_e_er_p759  = sc9ra3_e_er
		sc9ra3_e_er_p1800 = sc9ra3_e_er
		sc9ra3_e_er_p2200 = sc9ra3_e_er
		sc9ra3_e_er_p2359 = sc9ra3_e_er
		sc9ra3_e_macadj = e_mac_1 / 100
		sc9ra3_e_macadj_p759  = sc9ra3_e_macadj
		sc9ra3_e_macadj_p1800 = sc9ra3_e_macadj
		sc9ra3_e_macadj_p2200 = sc9ra3_e_macadj
		sc9ra3_e_macadj_p2359 = sc9ra3_e_macadj
		sc9ra3_d_dr_p1800 = d_mf86_sc9_3 + d_tra_mf86_sc9_3
		sc9ra3_d_dr_p2200 = d_mf810_sc9_3 + d_tra_mf810_sc9_3 + d_dlms_sc9_2_3 + d_sosc_sc9_1_3
		sc9ra3_d_dr_p2359 = d_all_sc9_3 + d_tra_all_sc9_3
		
		sc12ra2_e_er = ( e_edc_sc12_2 + e_tra_sc12_2 + e_sbc_1 + e_rpsp_1 + e_psls_sc12_2 + e_rdm_sc12_2 + e_drs_sc12_2 + e_cesds_1 + e_sosc_sc12_2 ) / 100
		sc12ra2_e_er_p759  = sc12ra2_e_er
		sc12ra2_e_er_p1800 = sc12ra2_e_er
		sc12ra2_e_er_p2200 = sc12ra2_e_er
		sc12ra2_e_er_p2359 = sc12ra2_e_er
		sc12ra2_e_macadj = e_mac_1 / 100
		sc12ra2_e_macadj_p759  = sc12ra2_e_macadj
		sc12ra2_e_macadj_p1800 = sc12ra2_e_macadj
		sc12ra2_e_macadj_p2200 = sc12ra2_e_macadj
		sc12ra2_e_macadj_p2359 = sc12ra2_e_macadj
		sc12ra2_d_dr_p1800 = d_mf86_sc12_2 + d_tra_mf86_sc12_2
		sc12ra2_d_dr_p2200 = d_mf810_sc12_2 + d_tra_mf810_sc12_2 + d_dlms_sc12_2 + d_sosc_sc12_2
		sc12ra2_d_dr_p2359 = d_all_sc12_2 + d_tra_all_sc12_2
	end if
%>
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 15">
<meta name=Originator content="Microsoft Word 15">
<link rel=File-List href="Rate%20Builder1_files/filelist.xml">
<link rel=Preview href="Rate%20Builder1_files/preview.wmf">
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
  <o:Author>LeRoi Isaacs</o:Author>
  <o:LastAuthor>LeRoi Isaacs</o:LastAuthor>
  <o:Revision>5</o:Revision>
  <o:TotalTime>6</o:TotalTime>
  <o:Created>2018-04-17T18:27:00Z</o:Created>
  <o:LastSaved>2018-04-17T18:32:00Z</o:LastSaved>
  <o:Pages>1</o:Pages>
  <o:Words>925</o:Words>
  <o:Characters>5273</o:Characters>
  <o:Lines>43</o:Lines>
  <o:Paragraphs>12</o:Paragraphs>
  <o:CharactersWithSpaces>6186</o:CharactersWithSpaces>
  <o:Version>16.00</o:Version>
 </o:DocumentProperties>
 <o:OfficeDocumentSettings>
  <o:AllowPNG/>
 </o:OfficeDocumentSettings>
</xml><![endif]-->
<link rel=themeData href="Rate%20Builder1_files/themedata.thmx">
<link rel=colorSchemeMapping href="Rate%20Builder1_files/colorschememapping.xml">
<!--[if gte mso 9]><xml>
 <w:WordDocument>
  <w:Zoom>110</w:Zoom>
  <w:SpellingState>Clean</w:SpellingState>
  <w:GrammarState>Clean</w:GrammarState>
  <w:TrackMoves>false</w:TrackMoves>
  <w:TrackFormatting/>
  <w:PunctuationKerning/>
  <w:ValidateAgainstSchemas/>
  <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>
  <w:IgnoreMixedContent>false</w:IgnoreMixedContent>
  <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>
  <w:DoNotPromoteQF/>
  <w:LidThemeOther>EN-US</w:LidThemeOther>
  <w:LidThemeAsian>X-NONE</w:LidThemeAsian>
  <w:LidThemeComplexScript>X-NONE</w:LidThemeComplexScript>
  <w:Compatibility>
   <w:BreakWrappedTables/>
   <w:SnapToGridInCell/>
   <w:WrapTextWithPunct/>
   <w:UseAsianBreakRules/>
   <w:DontGrowAutofit/>
   <w:SplitPgBreakAndParaMark/>
   <w:EnableOpenTypeKerning/>
   <w:DontFlipMirrorIndents/>
   <w:OverrideTableStyleHps/>
  </w:Compatibility>
  <m:mathPr>
   <m:mathFont m:val="Cambria Math"/>
   <m:brkBin m:val="before"/>
   <m:brkBinSub m:val="&#45;-"/>
   <m:smallFrac m:val="off"/>
   <m:dispDef/>
   <m:lMargin m:val="0"/>
   <m:rMargin m:val="0"/>
   <m:defJc m:val="centerGroup"/>
   <m:wrapIndent m:val="1440"/>
   <m:intLim m:val="subSup"/>
   <m:naryLim m:val="undOvr"/>
  </m:mathPr></w:WordDocument>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <w:LatentStyles DefLockedState="false" DefUnhideWhenUsed="false"
  DefSemiHidden="false" DefQFormat="false" DefPriority="99"
  LatentStyleCount="371">
  <w:LsdException Locked="false" Priority="0" QFormat="true" Name="Normal"/>
  <w:LsdException Locked="false" Priority="9" QFormat="true" Name="heading 1"/>
  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="heading 2"/>
  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="heading 3"/>
  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="heading 4"/>
  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="heading 5"/>
  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="heading 6"/>
  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="heading 7"/>
  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="heading 8"/>
  <w:LsdException Locked="false" Priority="9" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="heading 9"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 5"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 6"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 7"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 8"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index 9"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 1"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 2"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 3"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 4"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 5"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 6"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 7"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 8"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" Name="toc 9"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Normal Indent"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="footnote text"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="annotation text"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="header"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="footer"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="index heading"/>
  <w:LsdException Locked="false" Priority="35" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="caption"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="table of figures"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="envelope address"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="envelope return"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="footnote reference"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="annotation reference"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="line number"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="page number"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="endnote reference"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="endnote text"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="table of authorities"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="macro"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="toa heading"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Bullet"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Number"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List 5"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Bullet 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Bullet 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Bullet 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Bullet 5"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Number 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Number 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Number 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Number 5"/>
  <w:LsdException Locked="false" Priority="10" QFormat="true" Name="Title"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Closing"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Signature"/>
  <w:LsdException Locked="false" Priority="1" SemiHidden="true"
   UnhideWhenUsed="true" Name="Default Paragraph Font"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Body Text"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Body Text Indent"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Continue"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Continue 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Continue 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Continue 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="List Continue 5"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Message Header"/>
  <w:LsdException Locked="false" Priority="11" QFormat="true" Name="Subtitle"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Salutation"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Date"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Body Text First Indent"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Body Text First Indent 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Note Heading"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Body Text 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Body Text 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Body Text Indent 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Body Text Indent 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Block Text"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Hyperlink"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="FollowedHyperlink"/>
  <w:LsdException Locked="false" Priority="22" QFormat="true" Name="Strong"/>
  <w:LsdException Locked="false" Priority="20" QFormat="true" Name="Emphasis"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Document Map"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Plain Text"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="E-mail Signature"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Top of Form"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Bottom of Form"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Normal (Web)"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Acronym"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Address"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Cite"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Code"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Definition"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Keyboard"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Preformatted"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Sample"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Typewriter"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="HTML Variable"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Normal Table"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="annotation subject"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="No List"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Outline List 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Outline List 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Outline List 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Simple 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Simple 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Simple 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Classic 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Classic 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Classic 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Classic 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Colorful 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Colorful 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Colorful 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Columns 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Columns 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Columns 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Columns 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Columns 5"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Grid 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Grid 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Grid 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Grid 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Grid 5"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Grid 6"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Grid 7"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Grid 8"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table List 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table List 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table List 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table List 4"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table List 5"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table List 6"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table List 7"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table List 8"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table 3D effects 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table 3D effects 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table 3D effects 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Contemporary"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Elegant"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Professional"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Subtle 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Subtle 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Web 1"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Web 2"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Web 3"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Balloon Text"/>
  <w:LsdException Locked="false" Priority="39" Name="Table Grid"/>
  <w:LsdException Locked="false" SemiHidden="true" UnhideWhenUsed="true"
   Name="Table Theme"/>
  <w:LsdException Locked="false" SemiHidden="true" Name="Placeholder Text"/>
  <w:LsdException Locked="false" Priority="1" QFormat="true" Name="No Spacing"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 1"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 1"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 1"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 1"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 1"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 1"/>
  <w:LsdException Locked="false" SemiHidden="true" Name="Revision"/>
  <w:LsdException Locked="false" Priority="34" QFormat="true"
   Name="List Paragraph"/>
  <w:LsdException Locked="false" Priority="29" QFormat="true" Name="Quote"/>
  <w:LsdException Locked="false" Priority="30" QFormat="true"
   Name="Intense Quote"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 1"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 1"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 1"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 1"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 1"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 1"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 1"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 1"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 2"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 2"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 2"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 2"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 2"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 2"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 2"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 2"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 2"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 2"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 2"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 2"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 2"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 2"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 3"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 3"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 3"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 3"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 3"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 3"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 3"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 3"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 3"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 3"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 3"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 3"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 3"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 3"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 4"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 4"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 4"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 4"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 4"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 4"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 4"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 4"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 4"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 4"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 4"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 4"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 4"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 4"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 5"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 5"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 5"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 5"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 5"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 5"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 5"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 5"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 5"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 5"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 5"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 5"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 5"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 5"/>
  <w:LsdException Locked="false" Priority="60" Name="Light Shading Accent 6"/>
  <w:LsdException Locked="false" Priority="61" Name="Light List Accent 6"/>
  <w:LsdException Locked="false" Priority="62" Name="Light Grid Accent 6"/>
  <w:LsdException Locked="false" Priority="63" Name="Medium Shading 1 Accent 6"/>
  <w:LsdException Locked="false" Priority="64" Name="Medium Shading 2 Accent 6"/>
  <w:LsdException Locked="false" Priority="65" Name="Medium List 1 Accent 6"/>
  <w:LsdException Locked="false" Priority="66" Name="Medium List 2 Accent 6"/>
  <w:LsdException Locked="false" Priority="67" Name="Medium Grid 1 Accent 6"/>
  <w:LsdException Locked="false" Priority="68" Name="Medium Grid 2 Accent 6"/>
  <w:LsdException Locked="false" Priority="69" Name="Medium Grid 3 Accent 6"/>
  <w:LsdException Locked="false" Priority="70" Name="Dark List Accent 6"/>
  <w:LsdException Locked="false" Priority="71" Name="Colorful Shading Accent 6"/>
  <w:LsdException Locked="false" Priority="72" Name="Colorful List Accent 6"/>
  <w:LsdException Locked="false" Priority="73" Name="Colorful Grid Accent 6"/>
  <w:LsdException Locked="false" Priority="19" QFormat="true"
   Name="Subtle Emphasis"/>
  <w:LsdException Locked="false" Priority="21" QFormat="true"
   Name="Intense Emphasis"/>
  <w:LsdException Locked="false" Priority="31" QFormat="true"
   Name="Subtle Reference"/>
  <w:LsdException Locked="false" Priority="32" QFormat="true"
   Name="Intense Reference"/>
  <w:LsdException Locked="false" Priority="33" QFormat="true" Name="Book Title"/>
  <w:LsdException Locked="false" Priority="37" SemiHidden="true"
   UnhideWhenUsed="true" Name="Bibliography"/>
  <w:LsdException Locked="false" Priority="39" SemiHidden="true"
   UnhideWhenUsed="true" QFormat="true" Name="TOC Heading"/>
  <w:LsdException Locked="false" Priority="41" Name="Plain Table 1"/>
  <w:LsdException Locked="false" Priority="42" Name="Plain Table 2"/>
  <w:LsdException Locked="false" Priority="43" Name="Plain Table 3"/>
  <w:LsdException Locked="false" Priority="44" Name="Plain Table 4"/>
  <w:LsdException Locked="false" Priority="45" Name="Plain Table 5"/>
  <w:LsdException Locked="false" Priority="40" Name="Grid Table Light"/>
  <w:LsdException Locked="false" Priority="46" Name="Grid Table 1 Light"/>
  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2"/>
  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3"/>
  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4"/>
  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark"/>
  <w:LsdException Locked="false" Priority="51" Name="Grid Table 6 Colorful"/>
  <w:LsdException Locked="false" Priority="52" Name="Grid Table 7 Colorful"/>
  <w:LsdException Locked="false" Priority="46"
   Name="Grid Table 1 Light Accent 1"/>
  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 1"/>
  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 1"/>
  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 1"/>
  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 1"/>
  <w:LsdException Locked="false" Priority="51"
   Name="Grid Table 6 Colorful Accent 1"/>
  <w:LsdException Locked="false" Priority="52"
   Name="Grid Table 7 Colorful Accent 1"/>
  <w:LsdException Locked="false" Priority="46"
   Name="Grid Table 1 Light Accent 2"/>
  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 2"/>
  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 2"/>
  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 2"/>
  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 2"/>
  <w:LsdException Locked="false" Priority="51"
   Name="Grid Table 6 Colorful Accent 2"/>
  <w:LsdException Locked="false" Priority="52"
   Name="Grid Table 7 Colorful Accent 2"/>
  <w:LsdException Locked="false" Priority="46"
   Name="Grid Table 1 Light Accent 3"/>
  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 3"/>
  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 3"/>
  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 3"/>
  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 3"/>
  <w:LsdException Locked="false" Priority="51"
   Name="Grid Table 6 Colorful Accent 3"/>
  <w:LsdException Locked="false" Priority="52"
   Name="Grid Table 7 Colorful Accent 3"/>
  <w:LsdException Locked="false" Priority="46"
   Name="Grid Table 1 Light Accent 4"/>
  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 4"/>
  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 4"/>
  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 4"/>
  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 4"/>
  <w:LsdException Locked="false" Priority="51"
   Name="Grid Table 6 Colorful Accent 4"/>
  <w:LsdException Locked="false" Priority="52"
   Name="Grid Table 7 Colorful Accent 4"/>
  <w:LsdException Locked="false" Priority="46"
   Name="Grid Table 1 Light Accent 5"/>
  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 5"/>
  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 5"/>
  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 5"/>
  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 5"/>
  <w:LsdException Locked="false" Priority="51"
   Name="Grid Table 6 Colorful Accent 5"/>
  <w:LsdException Locked="false" Priority="52"
   Name="Grid Table 7 Colorful Accent 5"/>
  <w:LsdException Locked="false" Priority="46"
   Name="Grid Table 1 Light Accent 6"/>
  <w:LsdException Locked="false" Priority="47" Name="Grid Table 2 Accent 6"/>
  <w:LsdException Locked="false" Priority="48" Name="Grid Table 3 Accent 6"/>
  <w:LsdException Locked="false" Priority="49" Name="Grid Table 4 Accent 6"/>
  <w:LsdException Locked="false" Priority="50" Name="Grid Table 5 Dark Accent 6"/>
  <w:LsdException Locked="false" Priority="51"
   Name="Grid Table 6 Colorful Accent 6"/>
  <w:LsdException Locked="false" Priority="52"
   Name="Grid Table 7 Colorful Accent 6"/>
  <w:LsdException Locked="false" Priority="46" Name="List Table 1 Light"/>
  <w:LsdException Locked="false" Priority="47" Name="List Table 2"/>
  <w:LsdException Locked="false" Priority="48" Name="List Table 3"/>
  <w:LsdException Locked="false" Priority="49" Name="List Table 4"/>
  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark"/>
  <w:LsdException Locked="false" Priority="51" Name="List Table 6 Colorful"/>
  <w:LsdException Locked="false" Priority="52" Name="List Table 7 Colorful"/>
  <w:LsdException Locked="false" Priority="46"
   Name="List Table 1 Light Accent 1"/>
  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 1"/>
  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 1"/>
  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 1"/>
  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 1"/>
  <w:LsdException Locked="false" Priority="51"
   Name="List Table 6 Colorful Accent 1"/>
  <w:LsdException Locked="false" Priority="52"
   Name="List Table 7 Colorful Accent 1"/>
  <w:LsdException Locked="false" Priority="46"
   Name="List Table 1 Light Accent 2"/>
  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 2"/>
  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 2"/>
  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 2"/>
  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 2"/>
  <w:LsdException Locked="false" Priority="51"
   Name="List Table 6 Colorful Accent 2"/>
  <w:LsdException Locked="false" Priority="52"
   Name="List Table 7 Colorful Accent 2"/>
  <w:LsdException Locked="false" Priority="46"
   Name="List Table 1 Light Accent 3"/>
  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 3"/>
  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 3"/>
  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 3"/>
  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 3"/>
  <w:LsdException Locked="false" Priority="51"
   Name="List Table 6 Colorful Accent 3"/>
  <w:LsdException Locked="false" Priority="52"
   Name="List Table 7 Colorful Accent 3"/>
  <w:LsdException Locked="false" Priority="46"
   Name="List Table 1 Light Accent 4"/>
  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 4"/>
  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 4"/>
  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 4"/>
  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 4"/>
  <w:LsdException Locked="false" Priority="51"
   Name="List Table 6 Colorful Accent 4"/>
  <w:LsdException Locked="false" Priority="52"
   Name="List Table 7 Colorful Accent 4"/>
  <w:LsdException Locked="false" Priority="46"
   Name="List Table 1 Light Accent 5"/>
  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 5"/>
  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 5"/>
  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 5"/>
  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 5"/>
  <w:LsdException Locked="false" Priority="51"
   Name="List Table 6 Colorful Accent 5"/>
  <w:LsdException Locked="false" Priority="52"
   Name="List Table 7 Colorful Accent 5"/>
  <w:LsdException Locked="false" Priority="46"
   Name="List Table 1 Light Accent 6"/>
  <w:LsdException Locked="false" Priority="47" Name="List Table 2 Accent 6"/>
  <w:LsdException Locked="false" Priority="48" Name="List Table 3 Accent 6"/>
  <w:LsdException Locked="false" Priority="49" Name="List Table 4 Accent 6"/>
  <w:LsdException Locked="false" Priority="50" Name="List Table 5 Dark Accent 6"/>
  <w:LsdException Locked="false" Priority="51"
   Name="List Table 6 Colorful Accent 6"/>
  <w:LsdException Locked="false" Priority="52"
   Name="List Table 7 Colorful Accent 6"/>
 </w:LatentStyles>
</xml><![endif]-->
<style>
<!--
 /* Font Definitions */
 @font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;
	mso-font-charset:0;
	mso-generic-font-family:roman;
	mso-font-pitch:variable;
	mso-font-signature:-536869121 1107305727 33554432 0 415 0;}
@font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;
	mso-font-charset:0;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:-536859905 -1073732485 9 0 511 0;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-parent:"";
	margin-top:0in;
	margin-right:0in;
	margin-bottom:8.0pt;
	margin-left:0in;
	line-height:107%;
	mso-pagination:widow-orphan;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;
	mso-ascii-font-family:Calibri;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:Calibri;
	mso-fareast-theme-font:minor-latin;
	mso-hansi-font-family:Calibri;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:minor-bidi;}
span.SpellE
	{mso-style-name:"";
	mso-spl-e:yes;}
.MsoChpDefault
	{mso-style-type:export-only;
	mso-default-props:yes;
	font-family:"Calibri",sans-serif;
	mso-ascii-font-family:Calibri;
	mso-ascii-theme-font:minor-latin;
	mso-fareast-font-family:Calibri;
	mso-fareast-theme-font:minor-latin;
	mso-hansi-font-family:Calibri;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:minor-bidi;}
.MsoPapDefault
	{mso-style-type:export-only;
	margin-bottom:8.0pt;
	line-height:107%;}
@page WordSection1
	{size:8.5in 11.0in;
	margin:1.0in 1.0in 1.0in 1.0in;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;
	mso-paper-source:0;}
div.WordSection1
	{page:WordSection1;}
-->
</style>
<!--[if gte mso 10]>
<style>
 /* Style Definitions */
 table.MsoNormalTable
	{mso-style-name:"Table Normal";
	mso-tstyle-rowband-size:0;
	mso-tstyle-colband-size:0;
	mso-style-noshow:yes;
	mso-style-priority:99;
	mso-style-parent:"";
	mso-padding-alt:0in 5.4pt 0in 5.4pt;
	mso-para-margin-top:0in;
	mso-para-margin-right:0in;
	mso-para-margin-bottom:8.0pt;
	mso-para-margin-left:0in;
	line-height:107%;
	mso-pagination:widow-orphan;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;
	mso-ascii-font-family:Calibri;
	mso-ascii-theme-font:minor-latin;
	mso-hansi-font-family:Calibri;
	mso-hansi-theme-font:minor-latin;
	mso-bidi-font-family:"Times New Roman";
	mso-bidi-theme-font:minor-bidi;}
</style>
<![endif]--><!--[if gte mso 9]><xml>
 <o:shapedefaults v:ext="edit" spidmax="1026"/>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout></xml><![endif]-->
</head>

<body lang=EN-US style='tab-interval:.5in'>
<form name="RateBuilderRates" method="post" action ="saveRates.asp">
<div class=WordSection1>

<table class=MsoNormalTable border=0 cellspacing=0 cellpadding=0 width=700
 style='width:525.1pt;border-collapse:collapse;mso-yfti-tbllook:1184;
 mso-padding-alt:0in 5.4pt 0in 5.4pt'>
 <tr style='mso-yfti-irow:0;mso-yfti-firstrow:yes;height:15.0pt'>
  <td width=700 nowrap colspan=4 valign=bottom style='width:525.1pt;border:
  solid windowtext 1.0pt;mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal align=center style='margin-bottom:0in;margin-bottom:.0001pt;
  text-align:center;line-height:normal'><b style='mso-bidi-font-weight:normal'><span
  style='font-size:18.0pt;mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Rate Builder<o:p></o:p></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:1;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;border:none;mso-border-top-alt:
  solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;border:none;
  mso-border-top-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;border:none;
  border-bottom:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-bottom-alt:solid windowtext .5pt;
  padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;border:none;
  mso-border-top-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:2;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;border:none;
  border-right:solid windowtext 1.0pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;border-top:none;
  border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal align=center style='margin-bottom:0in;margin-bottom:.0001pt;
  text-align:center;line-height:normal'><b style='mso-bidi-font-weight:normal'><span
  style='font-size:18.0pt;mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'><%= monthname(month, true) &"" & year %></span></b></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;border:none;
  mso-border-left-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:3;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;border:none;border-bottom:
  solid windowtext 1.0pt;mso-border-bottom-alt:solid windowtext .5pt;
  padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;border:none;
  border-bottom:solid windowtext 1.0pt;mso-border-bottom-alt:solid windowtext .5pt;
  padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;border:none;
  border-bottom:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-bottom-alt:solid windowtext .5pt;
  padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;border:none;
  border-bottom:solid windowtext 1.0pt;mso-border-bottom-alt:solid windowtext .5pt;
  padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:4;height:15.0pt'>
  <td width=700 nowrap colspan=4 valign=bottom style='width:525.1pt;border:
  solid windowtext 1.0pt;border-top:none;mso-border-top-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal align=center style='margin-bottom:0in;margin-bottom:.0001pt;
  text-align:center;line-height:normal'><b style='mso-bidi-font-weight:normal'><span
  style='font-size:18.0pt;mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Computed Rates<o:p></o:p></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:5;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;border:none;mso-border-top-alt:
  solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;border:none;
  mso-border-top-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;border:none;
  mso-border-top-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;border:none;
  mso-border-top-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:6;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:7;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='font-size:10.0pt;font-family:"Times New Roman",serif;
  mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='font-size:10.0pt;font-family:"Times New Roman",serif;
  mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='font-size:10.0pt;font-family:"Times New Roman",serif;
  mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='font-size:10.0pt;font-family:"Times New Roman",serif;
  mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:8;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='font-size:10.0pt;font-family:"Times New Roman",serif;
  mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='font-size:10.0pt;font-family:"Times New Roman",serif;
  mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='font-size:10.0pt;font-family:"Times New Roman",serif;
  mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='font-size:10.0pt;font-family:"Times New Roman",serif;
  mso-fareast-font-family:"Times New Roman"'><o:p>&nbsp;</o:p></span></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:9;height:15.0pt'>
  <td width=324 nowrap colspan=2 valign=bottom style='width:243.2pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><b><span style='font-size:14.0pt;mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></b></span><b><span
  style='font-size:14.0pt;mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'> SC9R1<o:p></o:p></span></b></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:10;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Static<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Billing and Payment Processing Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1_s_bppc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:11;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><span style='mso-ascii-font-family:Calibri;
  mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></span><span
  style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
  mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black'>
  Meter Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1_s_cmc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:12;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1_e_er %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:13;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Systems Benefit Charge <o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1_e_sbc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:14;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Merchant Function Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1_e_mfc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Clean Energy Standard Supply Surcharge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1_e_cesss %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1_e_tsc %></span></b></p>
  </td>
 </tr>
  
 <tr style='mso-yfti-irow:16;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>MSC-CAP<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1_d_msccap %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:17;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>0-5<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1_d_dr_l05 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:18;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>6-99999999999<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1_d_dr_l999 %></span></b></p>
  </td>
 </tr>
<tr style='mso-yfti-irow:17;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Dynamic Load Management Surcharge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'><o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1_d_dlms %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1_d_tsc %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:19;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:20;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:21;height:15.0pt'>
  <td width=324 nowrap colspan=2 valign=bottom style='width:243.2pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><b><span style='font-size:14.0pt;mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></b></span><b><span
  style='font-size:14.0pt;mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'> SC9R1-Rider M<o:p></o:p></span></b></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:22;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Static<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Billing and Payment Processing Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1m_s_bppc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:23;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><span style='mso-ascii-font-family:Calibri;
  mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></span><span
  style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
  mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black'>
  Meter Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1m_s_cmc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:24;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1m_e_er %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:25;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Systems Benefit Charge <o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1m_e_sbc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:26;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Merchant Function Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1m_e_mfc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:27;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Clean Energy Standard Supply Surcharge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1m_e_cesss %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1m_e_tsc %></span></b></p>
  </td>
 </tr>  
 <tr style='mso-yfti-irow:28;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>MSC-CAP<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1m_d_msccap %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:29;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>0-5<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1m_d_dr_l05 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:30;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>6-99999999999<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1m_d_dr_l999 %></span></b></p>
  </td>
 </tr>
<tr style='mso-yfti-irow:17;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Dynamic Load Management Surcharge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'><o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1m_d_dlms %></span></b></p>
  </td>
 </tr>  
 <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r1m_d_tsc %></span></b></p>
  </td>
 </tr>  
 <tr style='mso-yfti-irow:31;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:32;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:33;height:15.0pt'>
  <td width=324 nowrap colspan=2 valign=bottom style='width:243.2pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><b><span style='font-size:14.0pt;mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></b></span><b><span
  style='font-size:14.0pt;mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'> SC9RA1<o:p></o:p></span></b></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:34;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Static<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Billing and Payment Processing Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1_s_bppc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:35;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><span style='mso-ascii-font-family:Calibri;
  mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></span><span
  style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
  mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black'>
  Meter Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1_s_cmc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:36;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy <o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1_e_er %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:37;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Systems Benefit Charge <o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1_e_sbc %></span></b></p>
  </td>
 </tr>
  <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1_e_tsc %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:38;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>0-5<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1_d_dr_l05 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:39;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>6-99999999999<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1_d_dr_l999 %></span></b></p>
  </td>
 </tr>
<tr style='mso-yfti-irow:17;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Dynamic Load Management Surcharge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'><o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1_d_dlms %></span></b></p>
  </td>
 </tr>  
  <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1_d_tsc %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:40;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:41;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:42;height:15.0pt'>
  <td width=324 nowrap colspan=2 valign=bottom style='width:243.2pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><b><span style='font-size:14.0pt;mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></b></span><b><span
  style='font-size:14.0pt;mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'> SC9RA1-Rider M<o:p></o:p></span></b></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:43;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Static<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Billing and Payment Processing Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1m_s_bppc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:44;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><span style='mso-ascii-font-family:Calibri;
  mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></span><span
  style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
  mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black'>
  Meter Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1m_s_cmc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:45;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy <o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1m_e_er %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:46;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Systems Benefit Charge <o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1m_e_sbc %></span></b></p>
  </td>
 </tr>
  <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1m_e_tsc %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:47;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>0-5<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1m_d_dr_l05 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:48;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>6-99999999999<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1m_d_dr_l999 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:17;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Dynamic Load Management Surcharge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'><o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1m_d_dlms %></span></b></p>
  </td>
 </tr>  
  <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra1m_d_tsc %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:49;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:50;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:51;height:15.0pt'>
  <td width=324 nowrap colspan=2 valign=bottom style='width:243.2pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><b><span style='font-size:14.0pt;mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></b></span><b><span
  style='font-size:14.0pt;mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'> SC9R2<o:p></o:p></span></b></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:52;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Summer: June-September<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Winter: other 8 months<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:53;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=492 nowrap colspan=2 valign=bottom style='width:369.1pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:red'>All rate peaks of this rate's entry need to choose winter or
  summer<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:54;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Static<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Billing and Payment Processing Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2_s_bppc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:55;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><span style='mso-ascii-font-family:Calibri;
  mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></span><span
  style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
  mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black'>
  Meter Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2_s_cmc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:56;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy <o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy Rate <o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2_e_er %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:57;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Systems Benefit Charge <o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2_e_sbc_p1800 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:58;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Merchant Function Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2_e_mfc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:59;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Clean Energy Standard Supply Surcharge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2_e_cesss %></span></b></p>
  </td>
 </tr>
  <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2_e_tsc %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:60;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>MSC-CAP<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-10)(Mo-Fr 800-2200)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2_d_mscadj_p2200 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:61;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-6)(Mo-Fr 800-1800)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2_d_dr_p1800 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:62;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-10)(Mo-Fr 800-2200)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2_d_dr_p2200 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:63;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Off Peak (Weekend) (Sa-Su 0-2359)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2_d_dr_p2359 %></span></b></p>
  </td>
 </tr>
  <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2_d_tsc %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:64;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:65;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:66;height:15.0pt'>
  <td width=324 nowrap colspan=2 valign=bottom style='width:243.2pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><b><span style='font-size:14.0pt;mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></b></span><b><span
  style='font-size:14.0pt;mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'> SC9R2-Rider M<o:p></o:p></span></b></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:67;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Summer: June-September<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Winter: other 8 months<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:68;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=492 nowrap colspan=2 valign=bottom style='width:369.1pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:red'>All rate peaks of this rate's entry need to choose winter or
  summer<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:69;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Static<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Billing and Payment Processing Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2m_s_bppc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:70;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><span style='mso-ascii-font-family:Calibri;
  mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></span><span
  style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
  mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black'>
  Meter Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2m_s_cmc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:71;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy <o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy Rate <o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2m_e_er %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:72;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Systems Benefit Charge <o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2m_e_sbc_p1800 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:73;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Merchant Function Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2m_e_mfc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:74;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Clean Energy Standard Supply Surcharge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2m_e_cesss %></span></b></p>
  </td>
 </tr>
  <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2_e_tsc %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:75;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>MSC-CAP<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-10)(Mo-Fr 800-2200)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2m_d_mscadj_p2200 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:76;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-6)(Mo-Fr 800-1800)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2m_d_dr_p1800 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:77;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-10)(Mo-Fr 800-2200)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2m_d_dr_p2200 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:78;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Off Peak (Weekend) (Sa-Su 0-2359)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2m_d_dr_p2359 %></span></b></p>
  </td>
 </tr>
  <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9r2_d_tsc %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:79;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:80;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:81;height:15.0pt'>
  <td width=324 nowrap colspan=2 valign=bottom style='width:243.2pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><b><span style='font-size:14.0pt;mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></b></span><b><span
  style='font-size:14.0pt;mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'> SC9RA2<o:p></o:p></span></b></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:82;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Summer: June-September<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Winter: other 8 months<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:83;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=492 nowrap colspan=2 valign=bottom style='width:369.1pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:red'>All rate peaks of this rate's entry need to choose winter or
  summer<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:84;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Static<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Billing and Payment Processing Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra2_s_bppc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:85;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><span style='mso-ascii-font-family:Calibri;
  mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></span><span
  style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
  mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black'>
  Meter Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra2_s_cmc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:86;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy <o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy Rate <o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra2_e_er %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:87;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Systems Benefit Charge <o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra2_e_sbc_p1800 %></span></b></p>
  </td>
 </tr>
  <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra2_e_tsc %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:88;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-6)(Mo-Fr 800-1800)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra2_d_dr_p1800 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:89;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-10)(Mo-Fr 800-2200)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra2_d_dr_p2200 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:90;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Off Peak (Weekend) (Sa-Su 0-2359)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra2_d_dr_p2359 %></span></b></p>
  </td>
 </tr>
  <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra2_d_tsc %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:91;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:92;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:93;height:15.0pt'>
  <td width=324 nowrap colspan=2 valign=bottom style='width:243.2pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><b><span style='font-size:14.0pt;mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></b></span><b><span
  style='font-size:14.0pt;mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'> SC9RA3<o:p></o:p></span></b></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:94;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Summer: June-September<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Winter: other 8 months<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:95;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=492 nowrap colspan=2 valign=bottom style='width:369.1pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:red'>All rate peaks of this rate's entry need to choose winter or
  summer<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:96;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Static<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Billing and Payment Processing Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3_s_bppc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:97;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><span style='mso-ascii-font-family:Calibri;
  mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></span><span
  style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
  mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black'>
  Meter Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3_s_cmc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:98;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy <o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy Rate <o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3_e_er %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:99;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Systems Benefit Charge <o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3_e_sbc_p1800 %></span></b></p>
  </td>
 </tr>
  <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3_e_tsc %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:100;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-6)(Mo-Fr 800-1800)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3_d_dr_p1800 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:101;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-10)(Mo-Fr 800-2200)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3_d_dr_p2200 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:102;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Off Peak (Weekend) (Sa-Su 0-2359)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3_d_dr_p2359 %></span></b></p>
  </td>
 </tr>
  <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3_d_tsc %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:103;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:104;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:105;height:15.0pt'>
  <td width=324 nowrap colspan=2 valign=bottom style='width:243.2pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><b><span style='font-size:14.0pt;mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></b></span><b><span
  style='font-size:14.0pt;mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'> SC9RA3 - Rider M<o:p></o:p></span></b></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:106;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Summer: June-September<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Winter: other 8 months<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:107;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=492 nowrap colspan=2 valign=bottom style='width:369.1pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:red'>All rate peaks of this rate's entry need to choose winter or
  summer<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:108;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Static<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Billing and Payment Processing Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3m_s_bppc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:109;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><span style='mso-ascii-font-family:Calibri;
  mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></span><span
  style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
  mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black'>
  Meter Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3m_s_cmc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:110;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy <o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy Rate <o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3m_e_er %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:111;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Systems Benefit Charge <o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3m_e_sbc_p1800 %></span></b></p>
  </td>
 </tr>
  <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3_e_tsc %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:112;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-6)(Mo-Fr 800-1800)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3m_d_dr_p1800 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:113;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-10)(Mo-Fr 800-2200)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3m_d_dr_p2200 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:114;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Off Peak (Weekend) (Sa-Su 0-2359)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3m_d_dr_p2359 %></span></b></p>
  </td>
 </tr>
  <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc9ra3_d_tsc %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:115;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:116;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:117;height:15.0pt'>
  <td width=324 nowrap colspan=2 valign=bottom style='width:243.2pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><b><span style='font-size:14.0pt;mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></b></span><b><span
  style='font-size:14.0pt;mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'> SC12RA2<o:p></o:p></span></b></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:118;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Summer: June-September<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Winter: other 8 months<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:119;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=492 nowrap colspan=2 valign=bottom style='width:369.1pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:red'>All rate peaks of this rate's entry need to choose winter or
  summer<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:120;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Static<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Billing and Payment Processing Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc12ra2_s_bppc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:121;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span class=SpellE><span style='mso-ascii-font-family:Calibri;
  mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>ConEd</span></span><span
  style='mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
  mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black'>
  Meter Charge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc12ra2_s_cmc %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:122;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy <o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy Rate <o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc12ra2_e_er %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:123;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Systems Benefit Charge <o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc12ra2_e_sbc_p1800 %></span></b></p>
  </td>
 </tr>
  <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc12ra2_e_tsc %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:124;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand<o:p></o:p></span></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-6)(Mo-Fr 800-1800)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc12ra2_d_dr_p1800 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:125;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-10)(Mo-Fr 800-2200)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc12ra2_d_dr_p2200 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:126;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Off Peak (Weekend) (Sa-Su 0-2359)<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc12ra2_d_dr_p2359 %></span></b></p>
  </td>
 </tr>
  <tr style='mso-yfti-irow:15;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Tax Sur-credit<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= conedsc12ra2_d_tsc %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:127;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:128;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:129;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b><span style='font-size:14.0pt;mso-ascii-font-family:Calibri;
  mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'>SC9R1<o:p></o:p></span></b></p>
  </td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:130;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy <o:p></o:p></span></p>
  </td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy Rate<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9r1_e_er %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:131;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>MAC Adj. Factor</p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9r1_e_macadj %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:132;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand<o:p></o:p></span></p>
  </td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>MSC Adj. Factor</p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9r1_d_mscadj %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:133;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>0-5<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9r1_d_dr_l05 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:134;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>6-100<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9r1_d_dr_l100 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:135;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>101-99999999999<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9r1_d_dr_l999 %></span></b></p>
  </td>
 </tr>
<tr style='mso-yfti-irow:17;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Dynamic Load Management Surcharge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'><o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9r1_d_dlms %></span></b></p>
  </td>
 </tr>  
 <tr style='mso-yfti-irow:136;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:137;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:138;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='font-size:14.0pt;
  mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
  mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black'>SC9R2<o:p></o:p></span></b></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:139;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Summer: June-September<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Winter: other 8 months<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:140;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=492 nowrap colspan=2 valign=bottom style='width:369.1pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:red'>All rate peaks of this rate's entry need to choose winter or
  summer<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:141;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy <o:p></o:p></span></p>
  </td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy Rate <o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9r2_e_er %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:142;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>MAC Adj. Factor</p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9r2_e_macadj %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:143;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand<o:p></o:p></span></p>
  </td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>MSC Adj. Factor<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-10)(Mo-Fr 800-2200)<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9r2_d_mscadj_p2200 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:144;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-6)(Mo-Fr 800-1800)<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9r2_d_dr_p1800 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:145;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-10)(Mo-Fr 800-2200)<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9r2_d_dr_p2200 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:146;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>2359 Peak (Weekend) (Sa-Su 0-2359)<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9r2_d_dr_p2359 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:147;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:148;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:149;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='font-size:14.0pt;
  mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
  mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black'>SC9RA1<o:p></o:p></span></b></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:150;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy <o:p></o:p></span></p>
  </td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy Rate<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9ra1_e_er %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:151;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>MAC Adj. Factor<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9ra1_e_macadj %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:152;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand<o:p></o:p></span></p>
  </td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>0-5<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9ra1_d_dr_l05 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:153;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>6-100<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9ra1_d_dr_l100 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:154;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>101-99999999999<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9ra1_d_dr_l999 %></span></b></p>
  </td>
 </tr>
<tr style='mso-yfti-irow:17;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Dynamic Load Management Surcharge<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'><o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9ra1_d_dlms %></span></b></p>
  </td>
 </tr> 
 <tr style='mso-yfti-irow:155;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:156;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:157;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='font-size:14.0pt;
  mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
  mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black'>SC9RA2<o:p></o:p></span></b></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:158;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Summer: June-September<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Winter: other 8 months<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:159;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=492 nowrap colspan=2 valign=bottom style='width:369.1pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:red'>All rate peaks of this rate's entry need to choose winter or
  summer<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:160;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy <o:p></o:p></span></p>
  </td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy Rate <o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9ra2_e_er %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:161;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>MAC Adj. Factor</p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9ra2_e_macadj %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:162;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand<o:p></o:p></span></p>
  </td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-6)(Mo-Fr 800-1800)<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9ra2_d_dr_p1800 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:163;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-10)(Mo-Fr 800-2200)<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9ra2_d_dr_p2200 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:164;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>2359 Peak (Weekend) (Sa-Su 0-2359)<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9ra2_d_dr_p2359 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:165;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:166;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:167;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='font-size:14.0pt;
  mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
  mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black'>SC9RA3<o:p></o:p></span></b></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:168;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Summer: June-September<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Winter: other 8 months<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:169;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=492 nowrap colspan=2 valign=bottom style='width:369.1pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:red'>All rate peaks of this rate's entry need to choose winter or
  summer<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:170;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy <o:p></o:p></span></p>
  </td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy Rate <o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9ra3_e_er %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:171;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>MAC Adj. Factor</p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9ra3_e_macadj %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:172;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand<o:p></o:p></span></p>
  </td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-6)(Mo-Fr 800-1800)<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9ra3_d_dr_p1800 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:173;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-10)(Mo-Fr 800-2200)<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9ra3_d_dr_p2200 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:174;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>2359 Peak (Weekend) (Sa-Su 0-2359)<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc9ra3_d_dr_p2359 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:175;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:176;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:177;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='font-size:14.0pt;
  mso-ascii-font-family:Calibri;mso-fareast-font-family:"Times New Roman";
  mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;color:black'>SC12RA2<o:p></o:p></span></b></p>
  </td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:178;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=241 nowrap valign=bottom style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Summer: June-September<o:p></o:p></span></p>
  </td>
  <td width=251 nowrap valign=bottom style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Winter: other 8 months<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:179;height:15.0pt'>
  <td width=83 nowrap valign=bottom style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
  <td width=492 nowrap colspan=2 valign=bottom style='width:369.1pt;padding:
  0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:red'>All rate peaks of this rate's entry need to choose winter or
  summer<o:p></o:p></span></p>
  </td>
  <td width=125 nowrap valign=bottom style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'></td>
 </tr>
 <tr style='mso-yfti-irow:180;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy <o:p></o:p></span></p>
  </td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Energy Rate <o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc12ra2_e_er %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:181;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>MAC Adj. Factor</p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>all 4 rate peaks<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc12ra2_e_macadj %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:182;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand<o:p></o:p></span></p>
  </td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-6)(Mo-Fr 800-1800)<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc12ra2_d_dr_p1800 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:183;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Peak(8-10)(Mo-Fr 800-2200)<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc12ra2_d_dr_p2200 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:184;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'></td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Demand Rate<o:p></o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>2359 Peak (Weekend) (Sa-Su 0-2359)<o:p></o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><%= sc12ra2_d_dr_p2359 %></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:185;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal align=right style='margin-bottom:0in;margin-bottom:.0001pt;
  text-align:right;line-height:normal'><span style='mso-ascii-font-family:Calibri;
  mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><o:p>&nbsp;</o:p></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:186;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal align=right style='margin-bottom:0in;margin-bottom:.0001pt;
  text-align:right;line-height:normal'><span style='mso-ascii-font-family:Calibri;
  mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=241 style='width:180.7pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><o:p>&nbsp;</o:p></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:187;height:15.0pt'>
  <td width=83 style='width:62.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal align=right style='margin-bottom:0in;margin-bottom:.0001pt;
  text-align:right;line-height:normal'><span style='mso-ascii-font-family:Calibri;
  mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=241 style='width:180.7pt;border:none;border-bottom:solid windowtext 1.0pt;
  mso-border-bottom-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=251 style='width:188.4pt;border:none;border-bottom:solid windowtext 1.0pt;
  mso-border-bottom-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><span style='mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=125 style='width:93.5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><o:p>&nbsp;</o:p></span></b></p>
  </td>
 </tr>
 <tr style='mso-yfti-irow:188;mso-yfti-lastrow:yes;height:15.0pt'>
  <td width=83 style='width:62.5pt;border:none;border-right:solid windowtext 1.0pt;
  mso-border-right-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;
  height:15.0pt'>
  <p class=MsoNormal align=right style='margin-bottom:0in;margin-bottom:.0001pt;
  text-align:right;line-height:normal'><span style='mso-ascii-font-family:Calibri;
  mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><o:p>&nbsp;</o:p></span></p>
  </td>
  <td width=241 style='width:180.7pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal align=center style='margin-bottom:0in;margin-bottom:.0001pt;
  text-align:center;line-height:normal'><b style='mso-bidi-font-weight:normal'><span
  style='font-size:18.0pt;mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'>Edit Components<o:p></o:p></span></b></p>
  </td>
  <td width=251 style='width:188.4pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-alt:solid windowtext .5pt;padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal align=center style='margin-bottom:0in;margin-bottom:.0001pt;
  text-align:center;line-height:normal'><b style='mso-bidi-font-weight:normal'><span
  style='font-size:18.0pt;mso-ascii-font-family:Calibri;mso-fareast-font-family:
 "Times New Roman";mso-hansi-font-family:Calibri;mso-bidi-font-family:Calibri;
  color:black'><input type="submit" name ="action" value="Save Rate Values" class="standard" /><o:p></o:p></span></b></p>
  </td>
  <td width=125 style='width:93.5pt;border:none;mso-border-left-alt:solid windowtext .5pt;
  padding:0in 5.4pt 0in 5.4pt;height:15.0pt'>
  <p class=MsoNormal style='margin-bottom:0in;margin-bottom:.0001pt;line-height:
  normal'><b style='mso-bidi-font-weight:normal'><span style='mso-ascii-font-family:
  Calibri;mso-fareast-font-family:"Times New Roman";mso-hansi-font-family:Calibri;
  mso-bidi-font-family:Calibri;color:black'><o:p>&nbsp;</o:p></span></b></p>
  </td>
 </tr>
</table>

<p class=MsoNormal><o:p>&nbsp;</o:p></p>

<p class=MsoNormal><o:p>&nbsp;</o:p></p>

<p class=MsoNormal style='margin-left:.5in;text-indent:.5in'><b
style='mso-bidi-font-weight:normal'><span style='font-size:14.0pt;line-height:
107%'><o:p>&nbsp;</o:p></span></b></p>

</div>
		<input type="hidden" value="<%= rateperiod             %>" name="rateperiod"/>      
		<input type="hidden" value="<%=rbcid             %>" name="rbcid"/>        
		<input type="hidden" value="<%=rbid             %>" name="rbid"/>
		<input type="hidden" value="<%=sc9r1_e_er             %>" name="sc9r1_e_er"/>        
		<input type="hidden" value="<%=sc9r1_e_macadj         %>" name="sc9r1_e_macadj"/>
		<input type="hidden" value="<%=sc9r1_d_mscadj         %>" name="sc9r1_d_mscadj"/>
		<input type="hidden" value="<%=sc9r1_d_dr_l05           %>" name="sc9r1_d_dr_l05"/>
		<input type="hidden" value="<%=sc9r1_d_dr_l100         %>" name="sc9r1_d_dr_l100"/>
		<input type="hidden" value="<%=sc9r1_d_dr_l999         %>" name="sc9r1_d_dr_l999"/>
		<input type="hidden" value="<%=sc9r1_d_dlms         %>" name="sc9r1_d_dlms"/>
		<input type="hidden" value="<%=conedsc9r1_s_bppc      %>" name="conedsc9r1_s_bppc"/>
		<input type="hidden" value="<%=conedsc9r1_s_cmc       %>" name="conedsc9r1_s_cmc"/>
		<input type="hidden" value="<%=conedsc9r1_e_er        %>" name="conedsc9r1_e_er"/>
		<input type="hidden" value="<%=conedsc9r1_e_sbc    %>" name="conedsc9r1_e_sbc"/>
		<input type="hidden" value="<%=conedsc9r1_e_mfc       %>" name="conedsc9r1_e_mfc"/>
		<input type="hidden" value="<%=conedsc9r1_d_msccap    %>" name="conedsc9r1_d_msccap"/>
		<input type="hidden" value="<%=conedsc9r1_d_dr_l05      %>" name="conedsc9r1_d_dr_l05"/>
		<input type="hidden" value="<%=conedsc9r1_d_dr_l999    %>" name="conedsc9r1_d_dr_l999"/>
		<input type="hidden" value="<%=conedsc9r1_e_cesss    %>" name="conedsc9r1_e_cesss"/>
		<input type="hidden" value="<%=conedsc9r1_d_dlms    %>" name="conedsc9r1_d_dlms"/>
		<input type="hidden" value="<%=sc9r2_e_er_p759              %>" name="sc9r2_e_er_p759"/>
		<input type="hidden" value="<%=sc9r2_e_er_p1800             %>" name="sc9r2_e_er_p1800"/>
		<input type="hidden" value="<%=sc9r2_e_er_p2200             %>" name="sc9r2_e_er_p2200"/>
		<input type="hidden" value="<%=sc9r2_e_er_p2359             %>" name="sc9r2_e_er_p2359"/>
		<input type="hidden" value="<%=sc9r2_e_macadj_p759          %>" name="sc9r2_e_macadj_p759"/>
		<input type="hidden" value="<%=sc9r2_e_macadj_p1800         %>" name="sc9r2_e_macadj_p1800"/>
		<input type="hidden" value="<%=sc9r2_e_macadj_p2200         %>" name="sc9r2_e_macadj_p2200"/>
		<input type="hidden" value="<%=sc9r2_e_macadj_p2359         %>" name="sc9r2_e_macadj_p2359"/>
		<input type="hidden" value="<%=sc9r2_d_mscadj_p2200         %>" name="sc9r2_d_mscadj_p2200"/>
		<input type="hidden" value="<%=sc9r2_d_dr_p1800        %>" name="sc9r2_d_dr_p1800"/>
		<input type="hidden" value="<%=sc9r2_d_dr_p2200       %>" name="sc9r2_d_dr_p2200"/>
		<input type="hidden" value="<%=sc9r2_d_dr_p2359         %>" name="sc9r2_d_dr_p2359"/>
		<input type="hidden" value="<%=sc9ra1_e_er            %>" name="sc9ra1_e_er"/>
		<input type="hidden" value="<%=sc9ra1_e_macadj        %>" name="sc9ra1_e_macadj"/>
		<input type="hidden" value="<%=sc9ra1_d_dr_l05          %>" name="sc9ra1_d_dr_l05"/>
		<input type="hidden" value="<%=sc9ra1_d_dr_l100        %>" name="sc9ra1_d_dr_l100"/>
		<input type="hidden" value="<%=sc9ra1_d_dr_l999        %>" name="sc9ra1_d_dr_l999"/>
		<input type="hidden" value="<%=sc9ra1_d_dlms        %>" name="sc9ra1_d_dlms"/>
		<input type="hidden" value="<%=sc9ra2_e_er_p759             %>" name="sc9ra2_e_er_p759"/>
		<input type="hidden" value="<%=sc9ra2_e_er_p1800            %>" name="sc9ra2_e_er_p1800"/>
		<input type="hidden" value="<%=sc9ra2_e_er_p2200            %>" name="sc9ra2_e_er_p2200"/>
		<input type="hidden" value="<%=sc9ra2_e_er_p2359            %>" name="sc9ra2_e_er_p2359"/>
		<input type="hidden" value="<%=sc9ra2_e_macadj_p759         %>" name="sc9ra2_e_macadj_p759"/>
		<input type="hidden" value="<%=sc9ra2_e_macadj_p1800        %>" name="sc9ra2_e_macadj_p1800"/>
		<input type="hidden" value="<%=sc9ra2_e_macadj_p2200        %>" name="sc9ra2_e_macadj_p2200"/>
		<input type="hidden" value="<%=sc9ra2_e_macadj_p2359        %>" name="sc9ra2_e_macadj_p2359"/>
		<input type="hidden" value="<%=sc9ra2_d_dr_p1800       %>" name="sc9ra2_d_dr_p1800"/>
		<input type="hidden" value="<%=sc9ra2_d_dr_p2200      %>" name="sc9ra2_d_dr_p2200"/>
		<input type="hidden" value="<%=sc9ra2_d_dr_p2359        %>" name="sc9ra2_d_dr_p2359"/>
		<input type="hidden" value="<%=sc9ra3_e_er_p759             %>" name="sc9ra3_e_er_p759"/>
		<input type="hidden" value="<%=sc9ra3_e_er_p1800            %>" name="sc9ra3_e_er_p1800"/>
		<input type="hidden" value="<%=sc9ra3_e_er_p2200            %>" name="sc9ra3_e_er_p2200"/>
		<input type="hidden" value="<%=sc9ra3_e_er_p2359            %>" name="sc9ra3_e_er_p2359"/>
		<input type="hidden" value="<%=sc9ra3_e_macadj_p759         %>" name="sc9ra3_e_macadj_p759"/>
		<input type="hidden" value="<%=sc9ra3_e_macadj_p1800        %>" name="sc9ra3_e_macadj_p1800"/>
		<input type="hidden" value="<%=sc9ra3_e_macadj_p2200        %>" name="sc9ra3_e_macadj_p2200"/>
		<input type="hidden" value="<%=sc9ra3_e_macadj_p2359        %>" name="sc9ra3_e_macadj_p2359"/>
		<input type="hidden" value="<%=sc9ra3_d_dr_p1800       %>" name="sc9ra3_d_dr_p1800"/>
		<input type="hidden" value="<%=sc9ra3_d_dr_p2200      %>" name="sc9ra3_d_dr_p2200"/>
		<input type="hidden" value="<%=sc9ra3_d_dr_p2359        %>" name="sc9ra3_d_dr_p2359"/>
		<input type="hidden" value="<%=sc12ra2_e_er_p759            %>" name="sc12ra2_e_er_p759"/>
		<input type="hidden" value="<%=sc12ra2_e_er_p1800           %>" name="sc12ra2_e_er_p1800"/>
		<input type="hidden" value="<%=sc12ra2_e_er_p2200           %>" name="sc12ra2_e_er_p2200"/>
		<input type="hidden" value="<%=sc12ra2_e_er_p2359           %>" name="sc12ra2_e_er_p2359"/>
		<input type="hidden" value="<%=sc12ra2_e_macadj_p759        %>" name="sc12ra2_e_macadj_p759"/>
		<input type="hidden" value="<%=sc12ra2_e_macadj_p1800       %>" name="sc12ra2_e_macadj_p1800"/>
		<input type="hidden" value="<%=sc12ra2_e_macadj_p2200       %>" name="sc12ra2_e_macadj_p2200"/>
		<input type="hidden" value="<%=sc12ra2_e_macadj_p2359       %>" name="sc12ra2_e_macadj_p2359"/>
		<input type="hidden" value="<%=sc12ra2_d_dr_p1800      %>" name="sc12ra2_d_dr_p1800"/>
		<input type="hidden" value="<%=sc12ra2_d_dr_p2200     %>" name="sc12ra2_d_dr_p2200"/>
		<input type="hidden" value="<%=sc12ra2_d_dr_p2359       %>" name="sc12ra2_d_dr_p2359"/>

<input type="hidden" value="<%= conedsc12ra2_d_dr_p1800                  %>" name="conedsc12ra2_d_dr_p1800"/>
<input type="hidden" value="<%= conedsc12ra2_d_dr_p2200                  %>" name="conedsc12ra2_d_dr_p2200"/>
<input type="hidden" value="<%= conedsc12ra2_d_dr_p2359                  %>" name="conedsc12ra2_d_dr_p2359"/>

<input type="hidden" value="<%= conedsc12ra2_e_er                        %>" name="conedsc12ra2_e_er"/>
<input type="hidden" value="<%= conedsc12ra2_e_er_p1800                  %>" name="conedsc12ra2_e_er_p1800"/>
<input type="hidden" value="<%= conedsc12ra2_e_er_p2200                  %>" name="conedsc12ra2_e_er_p2200"/>
<input type="hidden" value="<%= conedsc12ra2_e_er_p2359                  %>" name="conedsc12ra2_e_er_p2359"/>
<input type="hidden" value="<%= conedsc12ra2_e_er_p759                   %>" name="conedsc12ra2_e_er_p759"/>
<input type="hidden" value="<%= conedsc12ra2_e_macadj                    %>" name="conedsc12ra2_e_macadj"/>
<input type="hidden" value="<%= conedsc12ra2_e_macadj_p1800              %>" name="conedsc12ra2_e_macadj_p1800"/>
<input type="hidden" value="<%= conedsc12ra2_e_macadj_p2200              %>" name="conedsc12ra2_e_macadj_p2200"/>
<input type="hidden" value="<%= conedsc12ra2_e_macadj_p2359              %>" name="conedsc12ra2_e_macadj_p2359"/>
<input type="hidden" value="<%= conedsc12ra2_e_macadj_p759               %>" name="conedsc12ra2_e_macadj_p759"/>
<input type="hidden" value="<%= conedsc12ra2_e_sbc_p1800              %>" name="conedsc12ra2_e_sbc_p1800"/>
<input type="hidden" value="<%= conedsc12ra2_e_sbc_p2200              %>" name="conedsc12ra2_e_sbc_p2200"/>
<input type="hidden" value="<%= conedsc12ra2_e_sbc_p2359              %>" name="conedsc12ra2_e_sbc_p2359"/>

<input type="hidden" value="<%= conedsc12ra2_s_bppc                      %>" name="conedsc12ra2_s_bppc"/>
<input type="hidden" value="<%= conedsc12ra2_s_cmc                       %>" name="conedsc12ra2_s_cmc"/>
<input type="hidden" value="<%= conedsc9r1m_d_dr_l05                      %>" name="conedsc9r1m_d_dr_l05"/>
<input type="hidden" value="<%= conedsc9r1m_d_dr_l999                    %>" name="conedsc9r1m_d_dr_l999"/>
<input type="hidden" value="<%= conedsc9r1m_d_mscadj                     %>" name="conedsc9r1m_d_mscadj"/>
<input type="hidden" value="<%= conedsc9r1m_d_msccap                     %>" name="conedsc9r1m_d_msccap"/>
<input type="hidden" value="<%= conedsc9r1m_e_cesss                      %>" name="conedsc9r1m_e_cesss"/>
<input type="hidden" value="<%= conedsc9r1m_e_er                         %>" name="conedsc9r1m_e_er"/>
<input type="hidden" value="<%= conedsc9r1m_e_macadj                     %>" name="conedsc9r1m_e_macadj"/>
<input type="hidden" value="<%= conedsc9r1m_e_sbc                     %>" name="conedsc9r1m_e_sbc"/>
<input type="hidden" value="<%= conedsc9r1m_e_mfc                        %>" name="conedsc9r1m_e_mfc"/>
<input type="hidden" value="<%= conedsc9r1m_s_bppc                       %>" name="conedsc9r1m_s_bppc"/>
<input type="hidden" value="<%= conedsc9r1m_s_cmc                        %>" name="conedsc9r1m_s_cmc"/>
<input type="hidden" value="<%= conedsc9r1m_d_dlms                        %>" name="conedsc9r1m_d_dlms"/>

<input type="hidden" value="<%= conedsc9r2_e_cesss                       %>" name="conedsc9r2_e_cesss"/>
<input type="hidden" value="<%= conedsc9r2_e_er                          %>" name="conedsc9r2_e_er"/>
<input type="hidden" value="<%= conedsc9r2_e_er_p1800                    %>" name="conedsc9r2_e_er_p1800"/>
<input type="hidden" value="<%= conedsc9r2_e_er_p2200                    %>" name="conedsc9r2_e_er_p2200"/>
<input type="hidden" value="<%= conedsc9r2_e_er_p2359                    %>" name="conedsc9r2_e_er_p2359"/>
<input type="hidden" value="<%= conedsc9r2_e_er_p759                     %>" name="conedsc9r2_e_er_p759"/>
<input type="hidden" value="<%= conedsc9r2_e_macadj                      %>" name="conedsc9r2_e_macadj"/>
<input type="hidden" value="<%= conedsc9r2_e_sbc_p1800                      %>" name="conedsc9r2_e_sbc_p1800"/>
<input type="hidden" value="<%= conedsc9r2_e_sbc_p2200                      %>" name="conedsc9r2_e_sbc_p2200"/>
<input type="hidden" value="<%= conedsc9r2_e_sbc_p2359                      %>" name="conedsc9r2_e_sbc_p2359"/>
<input type="hidden" value="<%= conedsc9r2_e_mfc                         %>" name="conedsc9r2_e_mfc"/>
<input type="hidden" value="<%= conedsc9r2_s_bppc                        %>" name="conedsc9r2_s_bppc"/>
<input type="hidden" value="<%= conedsc9r2_s_cmc                         %>" name="conedsc9r2_s_cmc"/>
<input type="hidden" value="<%= conedsc9r2_d_msccap                      %>" name="conedsc9r2_d_msccap"/>	

<input type="hidden" value="<%= conedsc9r2m_d_dr_p1800                   %>" name="conedsc9r2m_d_dr_p1800"/>
<input type="hidden" value="<%= conedsc9r2m_d_dr_p2200                   %>" name="conedsc9r2m_d_dr_p2200"/>
<input type="hidden" value="<%= conedsc9r2m_d_dr_p2359                   %>" name="conedsc9r2m_d_dr_p2359"/>
<input type="hidden" value="<%= conedsc9r2m_d_mscadj_p2200               %>" name="conedsc9r2m_d_mscadj_p2200"/>
<input type="hidden" value="<%= conedsc9r2m_e_cesss                      %>" name="conedsc9r2m_e_cesss"/>
<input type="hidden" value="<%= conedsc9r2m_e_er                         %>" name="conedsc9r2m_e_er"/>
<input type="hidden" value="<%= conedsc9r2m_e_er_p1800                   %>" name="conedsc9r2m_e_er_p1800"/>
<input type="hidden" value="<%= conedsc9r2m_e_er_p2200                   %>" name="conedsc9r2m_e_er_p2200"/>
<input type="hidden" value="<%= conedsc9r2m_e_er_p2359                   %>" name="conedsc9r2m_e_er_p2359"/>
<input type="hidden" value="<%= conedsc9r2m_e_er_p759                    %>" name="conedsc9r2m_e_er_p759"/>
<input type="hidden" value="<%= conedsc9r2m_e_macadj                     %>" name="conedsc9r2m_e_macadj"/>
<input type="hidden" value="<%= conedsc9r2m_e_macadj_p1800               %>" name="conedsc9r2m_e_macadj_p1800"/>
<input type="hidden" value="<%= conedsc9r2m_e_macadj_p2200               %>" name="conedsc9r2m_e_macadj_p2200"/>
<input type="hidden" value="<%= conedsc9r2m_d_msccap_p2200               %>" name="conedsc9r2m_d_msccap_p2200"/>
<input type="hidden" value="<%= conedsc9r2m_e_macadj_p2359               %>" name="conedsc9r2m_e_macadj_p2359"/>
<input type="hidden" value="<%= conedsc9r2m_e_macadj_p759                %>" name="conedsc9r2m_e_macadj_p759"/>
<input type="hidden" value="<%= conedsc9r2m_e_sbc_p1800               %>" name="conedsc9r2m_e_sbc_p1800"/>
<input type="hidden" value="<%= conedsc9r2m_e_sbc_p2200               %>" name="conedsc9r2m_e_sbc_p2200"/>
<input type="hidden" value="<%= conedsc9r2m_e_sbc_p2359               %>" name="conedsc9r2m_e_sbc_p2359"/>
<input type="hidden" value="<%= conedsc9r2m_e_mfc                        %>" name="conedsc9r2m_e_mfc"/>
<input type="hidden" value="<%= conedsc9r2m_s_bppc                       %>" name="conedsc9r2m_s_bppc"/>
<input type="hidden" value="<%= conedsc9r2m_s_cmc                        %>" name="conedsc9r2m_s_cmc"/>

<input type="hidden" value="<%= conedsc9ra1_d_dr_l05                      %>" name="conedsc9ra1_d_dr_l05"/>
<input type="hidden" value="<%= conedsc9ra1_d_dr_l999                    %>" name="conedsc9ra1_d_dr_l999"/>
<input type="hidden" value="<%= conedsc9ra1_d_dlms                     %>" name="conedsc9ra1_d_dlms"/>
<input type="hidden" value="<%= conedsc9ra1_e_er                         %>" name="conedsc9ra1_e_er"/>
<input type="hidden" value="<%= conedsc9ra1_e_sbc                     %>" name="conedsc9ra1_e_sbc"/>
<input type="hidden" value="<%= conedsc9ra1_s_bppc				         %>" name="conedsc9ra1_s_bppc"/>
<input type="hidden" value="<%= conedsc9ra1_s_cmc                        %>" name="conedsc9ra1_s_cmc"/>


<input type="hidden" value="<%= conedsc9ra1m_d_dr_l05                     %>" name="conedsc9ra1m_d_dr_l05"/>
<input type="hidden" value="<%= conedsc9ra1m_d_dr_l999                   %>" name="conedsc9ra1m_d_dr_l999"/>
<input type="hidden" value="<%= conedsc9ra1m_e_er                        %>" name="conedsc9ra1m_e_er"/>
<input type="hidden" value="<%= conedsc9ra1m_e_macadj                    %>" name="conedsc9ra1m_e_macadj"/>
<input type="hidden" value="<%= conedsc9ra1m_e_sbc                    %>" name="conedsc9ra1m_e_sbc"/>
<input type="hidden" value="<%= conedsc9ra1m_s_bppc                      %>" name="conedsc9ra1m_s_bppc"/>
<input type="hidden" value="<%= conedsc9ra1m_s_cmc                       %>" name="conedsc9ra1m_s_cmc"/>
<input type="hidden" value="<%= conedsc9ra1m_d_dlms                       %>" name="conedsc9ra1m_d_dlms"/>
<input type="hidden" value="<%= conedsc9ra1m_d_msccap                       %>" name="conedsc9ra1m_d_msccap"/>

<input type="hidden" value="<%= conedsc9ra2_d_dr_p1800                   %>" name="conedsc9ra2_d_dr_p1800"/>
<input type="hidden" value="<%= conedsc9ra2_d_dr_p2200                   %>" name="conedsc9ra2_d_dr_p2200"/>
<input type="hidden" value="<%= conedsc9ra2_d_dr_p2359                   %>" name="conedsc9ra2_d_dr_p2359"/>


<input type="hidden" value="<%= conedsc9ra2_e_er                         %>" name="conedsc9ra2_e_er"/>
<input type="hidden" value="<%= conedsc9ra2_e_er_p1800                   %>" name="conedsc9ra2_e_er_p1800"/>
<input type="hidden" value="<%= conedsc9ra2_e_er_p2200                   %>" name="conedsc9ra2_e_er_p2200"/>
<input type="hidden" value="<%= conedsc9ra2_e_er_p2359                   %>" name="conedsc9ra2_e_er_p2359"/>
<input type="hidden" value="<%= conedsc9ra2_e_er_p759                    %>" name="conedsc9ra2_e_er_p759"/>
<input type="hidden" value="<%= conedsc9ra2_e_macadj                     %>" name="conedsc9ra2_e_macadj"/>
<input type="hidden" value="<%= conedsc9ra2_e_macadj_p1800               %>" name="conedsc9ra2_e_macadj_p1800"/>
<input type="hidden" value="<%= conedsc9ra2_e_macadj_p2200               %>" name="conedsc9ra2_e_macadj_p2200"/>
<input type="hidden" value="<%= conedsc9ra2_e_macadj_p2359               %>" name="conedsc9ra2_e_macadj_p2359"/>
<input type="hidden" value="<%= conedsc9ra2_e_macadj_p759                %>" name="conedsc9ra2_e_macadj_p759"/>
<input type="hidden" value="<%= conedsc9ra2_e_sbc_p1800               %>" name="conedsc9ra2_e_sbc_p1800"/>
<input type="hidden" value="<%= conedsc9ra2_e_sbc_p2200               %>" name="conedsc9ra2_e_sbc_p2200"/>
<input type="hidden" value="<%= conedsc9ra2_e_sbc_p2359               %>" name="conedsc9ra2_e_sbc_p2359"/>
<input type="hidden" value="<%= conedsc9ra2_s_bppc                       %>" name="conedsc9ra2_s_bppc"/>
<input type="hidden" value="<%= conedsc9ra2_s_cmc                        %>" name="conedsc9ra2_s_cmc"/>

<input type="hidden" value="<%= conedsc9ra3_d_dr_p1800                   %>" name="conedsc9ra3_d_dr_p1800"/>
<input type="hidden" value="<%= conedsc9ra3_d_dr_p2200                   %>" name="conedsc9ra3_d_dr_p2200"/>
<input type="hidden" value="<%= conedsc9ra3_d_dr_p2359                   %>" name="conedsc9ra3_d_dr_p2359"/>

<input type="hidden" value="<%= conedsc9ra3_e_er                         %>" name="conedsc9ra3_e_er"/>
<input type="hidden" value="<%= conedsc9ra3_e_er_p1800                   %>" name="conedsc9ra3_e_er_p1800"/>
<input type="hidden" value="<%= conedsc9ra3_e_er_p2200                   %>" name="conedsc9ra3_e_er_p2200"/>
<input type="hidden" value="<%= conedsc9ra3_e_er_p2359                   %>" name="conedsc9ra3_e_er_p2359"/>
<input type="hidden" value="<%= conedsc9ra3_e_er_p759                    %>" name="conedsc9ra3_e_er_p759"/>
<input type="hidden" value="<%= conedsc9ra3_e_macadj                     %>" name="conedsc9ra3_e_macadj"/>
<input type="hidden" value="<%= conedsc9ra3_e_macadj_p1800               %>" name="conedsc9ra3_e_macadj_p1800"/>
<input type="hidden" value="<%= conedsc9ra3_e_macadj_p2200               %>" name="conedsc9ra3_e_macadj_p2200"/>
<input type="hidden" value="<%= conedsc9ra3_e_macadj_p2359               %>" name="conedsc9ra3_e_macadj_p2359"/>
<input type="hidden" value="<%= conedsc9ra3_e_macadj_p759                %>" name="conedsc9ra3_e_macadj_p759"/>
<input type="hidden" value="<%= conedsc9ra3_e_sbc_p1800               %>" name="conedsc9ra3_e_sbc_p1800"/>
<input type="hidden" value="<%= conedsc9ra3_e_sbc_p2200               %>" name="conedsc9ra3_e_sbc_p2200"/>
<input type="hidden" value="<%= conedsc9ra3_e_sbc_p2359               %>" name="conedsc9ra3_e_sbc_p2359"/>


<input type="hidden" value="<%= conedsc9ra3_s_bppc                       %>" name="conedsc9ra3_s_bppc"/>
<input type="hidden" value="<%= conedsc9ra3_s_cmc                        %>" name="conedsc9ra3_s_cmc"/>

<input type="hidden" value="<%= conedsc9ra3m_d_dr_p1800                  %>" name="conedsc9ra3m_d_dr_p1800"/>
<input type="hidden" value="<%= conedsc9ra3m_d_dr_p2200                  %>" name="conedsc9ra3m_d_dr_p2200"/>
<input type="hidden" value="<%= conedsc9ra3m_d_dr_p2359                  %>" name="conedsc9ra3m_d_dr_p2359"/>

<input type="hidden" value="<%= conedsc9ra3m_e_er                        %>" name="conedsc9ra3m_e_er"/>
<input type="hidden" value="<%= conedsc9ra3m_e_er_p1800                  %>" name="conedsc9ra3m_e_er_p1800"/>
<input type="hidden" value="<%= conedsc9ra3m_e_er_p2200                  %>" name="conedsc9ra3m_e_er_p2200"/>
<input type="hidden" value="<%= conedsc9ra3m_e_er_p2359                  %>" name="conedsc9ra3m_e_er_p2359"/>
<input type="hidden" value="<%= conedsc9ra3m_e_er_p759                   %>" name="conedsc9ra3m_e_er_p759"/>
<input type="hidden" value="<%= conedsc9ra3m_e_macadj                    %>" name="conedsc9ra3m_e_macadj"/>
<input type="hidden" value="<%= conedsc9ra3m_e_macadj_p1800              %>" name="conedsc9ra3m_e_macadj_p1800"/>
<input type="hidden" value="<%= conedsc9ra3m_e_macadj_p2200              %>" name="conedsc9ra3m_e_macadj_p2200"/>
<input type="hidden" value="<%= conedsc9ra3m_e_macadj_p2359              %>" name="conedsc9ra3m_e_macadj_p2359"/>
<input type="hidden" value="<%= conedsc9ra3m_e_macadj_p759               %>" name="conedsc9ra3m_e_macadj_p759"/>
<input type="hidden" value="<%= conedsc9ra3m_e_sbc_p1800              %>" name="conedsc9ra3m_e_sbc_p1800"/>
<input type="hidden" value="<%= conedsc9ra3m_e_sbc_p2200              %>" name="conedsc9ra3m_e_sbc_p2200"/>
<input type="hidden" value="<%= conedsc9ra3m_e_sbc_p2359              %>" name="conedsc9ra3m_e_sbc_p2359"/>

<input type="hidden" value="<%= conedsc9ra3m_s_bppc                      %>" name="conedsc9ra3m_s_bppc"/>
<input type="hidden" value="<%= conedsc9ra3m_s_cmc                       %>" name="conedsc9ra3m_s_cmc"/>	
<input type="hidden" value="<%=conedsc9r1_e_tsc    %>" name="conedsc9r1_e_tsc"/> 
<input type="hidden" value="<%=conedsc9r1_d_tsc    %>" name="conedsc9r1_d_tsc"/>
<input type="hidden" value="<%=conedsc9r1m_e_tsc   %>" name="conedsc9r1m_e_tsc"/>
<input type="hidden" value="<%=conedsc9r1m_d_tsc   %>" name="conedsc9r1m_d_tsc"/>
<input type="hidden" value="<%=conedsc9ra1_e_tsc   %>" name="conedsc9ra1_e_tsc"/>
<input type="hidden" value="<%=conedsc9ra1_d_tsc   %>" name="conedsc9ra1_d_tsc"/>
<input type="hidden" value="<%=conedsc9ra1m_e_tsc  %>" name="conedsc9ra1m_e_tsc"/>
<input type="hidden" value="<%=conedsc9ra1m_d_tsc  %>" name="conedsc9ra1m_d_tsc"/>
<input type="hidden" value="<%=conedsc9r2_e_tsc    %>" name="conedsc9r2_e_tsc"/>
<input type="hidden" value="<%=conedsc9r2_d_tsc    %>" name="conedsc9r2_d_tsc"/>
<input type="hidden" value="<%=conedsc9r2m_e_tsc   %>" name="conedsc9r2m_e_tsc"/>
<input type="hidden" value="<%=conedsc9r2m_d_tsc   %>" name="conedsc9r2m_d_tsc"/>
<input type="hidden" value="<%=conedsc9ra2_e_tsc   %>" name="conedsc9ra2_e_tsc"/>
<input type="hidden" value="<%=conedsc9ra2_d_tsc   %>" name="conedsc9ra2_d_tsc"/>
<input type="hidden" value="<%=conedsc9ra3_e_tsc   %>" name="conedsc9ra3_e_tsc"/>
<input type="hidden" value="<%=conedsc9ra3_d_tsc   %>" name="conedsc9ra3_d_tsc"/>
<input type="hidden" value="<%=conedsc9ra3m_e_tsc  %>" name="conedsc9ra3m_e_tsc"/>
<input type="hidden" value="<%=conedsc9ra3m_d_tsc  %>" name="conedsc9ra3m_d_tsc"/>
<input type="hidden" value="<%=conedsc12ra2_e_tsc   %>" name="conedsc12ra2_e_tsc"/>
<input type="hidden" value="<%=conedsc12ra2_d_tsc   %>" name="conedsc12ra2_d_tsc"/>	

	</form>
</body>

</html>
