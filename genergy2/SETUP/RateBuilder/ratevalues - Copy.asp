<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
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
	dim s_cms_r,s_cms_sc9_1,s_cms_sc9_m_1,s_cms_sc9_2,s_cms_sc1_1,s_cms_sc1_low_1,s_bppc_r,s_bppc_el_1,s_bppc_elgs_1, e_cesds_r, e_cesds_1, e_cesss_r, e_cesss_1, s_cms_sc9_3, s_cms_sc9_m_3, s_cms_sc12_2
			
	dim sc9r1_e_er, sc9r1_e_macadj, sc9r1_d_mscadj, sc9r1_d_dr_l5, sc9r1_d_dr_l100, sc9r1_d_dr_l999, conedsc9r1_s_bppc, conedsc9r1_s_cmc, conedsc9r1_e_er, conedsc9r1_e_macadj, conedsc9r1_e_mfc, conedsc9r1_d_mscadj, conedsc9r1_d_dr_l5, conedsc9r1_d_dr_l999, conedsc9r1_e_cesss, sc9r2_e_er_p759 ,sc9r2_e_er_p1800,sc9r2_e_er_p2200,sc9r2_e_er_p2359,sc9r2_e_er,sc9r2_e_macadj_p759 ,sc9r2_e_macadj_p1800,sc9r2_e_macadj_p2200,sc9r2_e_macadj_p2359,sc9r2_e_macadj,sc9r2_d_mscadj_p2200, sc9r2_d_dr_p1800, sc9r2_d_dr_p2200, sc9r2_d_dr_p2359, sc9ra1_e_er, sc9ra1_e_macadj, sc9ra1_d_dr_l5, sc9ra1_d_dr_l100, sc9ra1_d_dr_l999,sc9ra2_e_er_p759 ,sc9ra2_e_er_p1800,sc9ra2_e_er_p2200,sc9ra2_e_er_p2359,sc9ra2_e_er,sc9ra2_e_macadj_p759 ,sc9ra2_e_macadj_p1800,sc9ra2_e_macadj_p2200,sc9ra2_e_macadj_p2359,sc9ra2_e_macadj,sc9ra2_d_dr_p1800, sc9ra2_d_dr_p2200, sc9ra2_d_dr_p2359,sc9ra3_e_er_p759 ,sc9ra3_e_er_p1800,sc9ra3_e_er_p2200,sc9ra3_e_er_p2359,sc9ra3_e_er,sc9ra3_e_macadj_p759 ,sc9ra3_e_macadj_p1800,sc9ra3_e_macadj_p2200,sc9ra3_e_macadj_p2359,sc9ra3_e_macadj,sc9ra3_d_dr_p1800, sc9ra3_d_dr_p2200, sc9ra3_d_dr_p2359,sc12ra2_e_er_p759 ,sc12ra2_e_er_p1800,sc12ra2_e_er_p2200,sc12ra2_e_er_p2359,sc12ra2_e_er,sc12ra2_e_macadj_p759 ,sc12ra2_e_macadj_p1800,sc12ra2_e_macadj_p2200,sc12ra2_e_macadj_p2359,sc12ra2_e_macadj,sc12ra2_d_dr_p1800, sc12ra2_d_dr_p2200, sc12ra2_d_dr_p2359
	
	dim conedsc9r1m_s_bppc, conedsc9r1m_s_cmc, conedsc9r1m_e_er, conedsc9r1m_e_macadj, conedsc9r1m_e_mfc, conedsc9r1m_d_mscadj, conedsc9r1m_d_dr_l5, conedsc9r1m_d_dr_l999, conedsc9r1m_e_cesss
	dim conedsc9ra1_s_bppc, conedsc9ra1_s_cmc, conedsc9ra1_e_er, conedsc9ra1_e_macadj, conedsc9ra1_e_mfc, conedsc9ra1_d_mscadj, conedsc9ra1_d_dr_l5, conedsc9ra1_d_dr_l999, conedsc9ra1_e_cesss
	dim conedsc9ra1m_s_bppc, conedsc9ra1m_s_cmc, conedsc9ra1m_e_er, conedsc9ra1m_e_macadj, conedsc9ra1m_e_mfc, conedsc9ra1m_d_mscadj, conedsc9ra1m_d_dr_l5, conedsc9ra1m_d_dr_l999, conedsc9ra1m_e_cesss, conedsc9r2_s_bppc, conedsc9r2_s_cmc, conedsc9r2_e_er, conedsc9r2_e_macadj, conedsc9r2_e_mfc, conedsc9r2_d_mscadj,  conedsc9r2_e_cesss, conedsc9r2_d_dr_p1800, conedsc9r2_d_dr_p2200, conedsc9r2_d_dr_p2359, conedsc9r2_d_mscadj_p2200, conedsc9r2_e_macadj_p759, conedsc9r2_e_macadj_p1800, conedsc9r2_e_macadj_p2200, conedsc9r2_e_macadj_p2359, conedsc9r2_e_er_p759, conedsc9r2_e_er_p1800, conedsc9r2_e_er_p2200, conedsc9r2_e_er_p2359
	
	dim conedsc9r2m_s_bppc, conedsc9r2m_s_cmc, conedsc9r2m_e_er, conedsc9r2m_e_macadj, conedsc9r2m_e_mfc, conedsc9r2m_d_mscadj,  conedsc9r2m_e_cesss, conedsc9r2m_d_dr_p1800, conedsc9r2m_d_dr_p2200, conedsc9r2m_d_dr_p2359, conedsc9r2m_d_mscadj_p2200, conedsc9r2m_e_macadj_p759, conedsc9r2m_e_macadj_p1800, conedsc9r2m_e_macadj_p2200, conedsc9r2m_e_macadj_p2359, conedsc9r2m_e_er_p759, conedsc9r2m_e_er_p1800, conedsc9r2m_e_er_p2200, conedsc9r2m_e_er_p2359
	
	dim conedsc9ra2_s_bppc, conedsc9ra2_s_cmc, conedsc9ra2_e_er, conedsc9ra2_e_macadj, conedsc9ra2_e_mfc, conedsc9ra2_d_mscadj,  conedsc9ra2_e_cesss, conedsc9ra2_d_dr_p1800, conedsc9ra2_d_dr_p2200, conedsc9ra2_d_dr_p2359, conedsc9ra2_d_mscadj_p2200, conedsc9ra2_e_macadj_p759, conedsc9ra2_e_macadj_p1800, conedsc9ra2_e_macadj_p2200, conedsc9ra2_e_macadj_p2359, conedsc9ra2_e_er_p759, conedsc9ra2_e_er_p1800, conedsc9ra2_e_er_p2200, conedsc9ra2_e_er_p2359
	
	dim conedsc9ra3_s_bppc, conedsc9ra3_s_cmc, conedsc9ra3_e_er, conedsc9ra3_e_macadj, conedsc9ra3_e_mfc, conedsc9ra3_d_mscadj,  conedsc9ra3_e_cesss, conedsc9ra3_d_dr_p1800, conedsc9ra3_d_dr_p2200, conedsc9ra3_d_dr_p2359, conedsc9ra3_d_mscadj_p2200, conedsc9ra3_e_macadj_p759, conedsc9ra3_e_macadj_p1800, conedsc9ra3_e_macadj_p2200, conedsc9ra3_e_macadj_p2359, conedsc9ra3_e_er_p759, conedsc9ra3_e_er_p1800, conedsc9ra3_e_er_p2200, conedsc9ra3_e_er_p2359
	
	dim conedsc9ra3m_s_bppc, conedsc9ra3m_s_cmc, conedsc9ra3m_e_er, conedsc9ra3m_e_macadj, conedsc9ra3m_e_mfc, conedsc9ra3m_d_mscadj,  conedsc9ra3m_e_cesss, conedsc9ra3m_d_dr_p1800, conedsc9ra3m_d_dr_p2200, conedsc9ra3m_d_dr_p2359, conedsc9ra3m_d_mscadj_p2200, conedsc9ra3m_e_macadj_p759, conedsc9ra3m_e_macadj_p1800, conedsc9ra3m_e_macadj_p2200, conedsc9ra3m_e_macadj_p2359, conedsc9ra3m_e_er_p759, conedsc9ra3m_e_er_p1800, conedsc9ra3m_e_er_p2200, conedsc9ra3m_e_er_p2359
	
	dim conedsc12ra2_s_bppc, conedsc12ra2_s_cmc, conedsc12ra2_e_er, conedsc12ra2_e_macadj, conedsc12ra2_e_mfc, conedsc12ra2_d_mscadj,  conedsc12ra2_e_cesss, conedsc12ra2_d_dr_p1800, conedsc12ra2_d_dr_p2200, conedsc12ra2_d_dr_p2359, conedsc12ra2_d_mscadj_p2200, conedsc12ra2_e_macadj_p759, conedsc12ra2_e_macadj_p1800, conedsc12ra2_e_macadj_p2200, conedsc12ra2_e_macadj_p2359, conedsc12ra2_e_er_p759, conedsc12ra2_e_er_p1800, conedsc12ra2_e_er_p2200, conedsc12ra2_e_er_p2359
	
	
	rbid = trim(secureRequest("rbid"))
	strsql = "select * from ratebuilder where rbid='" & rbid & "'"
	rst1.Open strsql, cnn1
	rbcid = rst1("rbcid")
	rateperiod = rst1("rateperiod")
	year = datepart("yyyy", rateperiod)
	month = datepart("m", rateperiod)
	rst1.close
	
	strsql = "select * from ratebuildercomponents where rbcid='" & rbcid & "'"'strsql = "select * from ratebuildercomponents rbc join ratebuilder rb on rb.rbcid = rbc.rbcid where rb.rbid ="&rbid
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
		d_msccap_sc9r1_1 = (rst1("d_msccap_sc9r1_1"))
		d_msccap_sc9r2_1 = (rst1("d_msccap_sc9r2_1"))
		
		s_cms_r = (rst1("s_cms_r"))
		s_cms_sc9_1 = (rst1("s_cms_sc9_1"))
		s_cms_sc9_m_1	=(rst1("s_cms_sc9_m_1"))
		s_cms_sc9_2		=(rst1("s_cms_sc9_2"))
		s_cms_sc9_3		=(rst1("s_cms_sc9_3"))
		s_cms_sc9_m_3	=(rst1("s_cms_sc9_m_3"))
		s_cms_sc12_2	=(rst1("s_cms_sc12_2"))
		
		s_bppc_r = (rst1("s_bppc_r"))
		s_bppc_el_1 = (rst1("s_bppc_el_1"))
		
		sc9r1_e_er = ( e_edc_sc9_1 + e_tra_sc9_1 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_mfc_sc912_1 + e_cesds_1 + e_cesss_1) / 100
		sc9r1_e_macadj = e_mac_1 / 100
		sc9r1_d_mscadj = d_msccap_sc9r1_1
		sc9r1_d_dr_l5 = d_mc_1 + d_tra_mc_1
		sc9r1_d_dr_l100 = d_o5_1 + d_tra_o5_1
		sc9r1_d_dr_l999 = d_o5_1 + d_tra_o5_1
		
		conedsc9r1_s_bppc = s_bppc_el_1
		conedsc9r1_s_cmc = s_cms_sc9_1
		conedsc9r1_e_er = ( e_edc_sc9_1 + e_tra_sc9_1 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
		conedsc9r1_e_macadj = e_mac_1 / 100
		conedsc9r1_e_mfc = e_mfc_sc912_1 / 100
		conedsc9r1_e_cesss = e_cesss_1 / 100
		conedsc9r1_d_mscadj = d_msccap_sc9r1_1
		conedsc9r1_d_dr_l5 = d_mc_1 + d_tra_mc_1
		conedsc9r1_d_dr_l999 = d_o5_1 + d_tra_o5_1
		
		conedsc9r1m_s_bppc = s_bppc_el_1
		conedsc9r1m_s_cmc = s_cms_sc9_m_1
		conedsc9r1m_e_er = ( e_edc_sc9_1 + e_tra_sc9_1 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
		conedsc9r1m_e_macadj = e_mac_1 / 100
		conedsc9r1m_e_mfc = e_mfc_sc912_1 / 100
		conedsc9r1m_e_cesss = e_cesss_1 / 100
		conedsc9r1m_d_mscadj = d_msccap_m_1
		conedsc9r1m_d_dr_l5 = d_mc_1 + d_tra_mc_1
		conedsc9r1m_d_dr_l999 = d_o5_1 + d_tra_o5_1

		conedsc9ra1_s_bppc = s_bppc_el_1
		conedsc9ra1_s_cmc = s_cms_sc9_1
		conedsc9ra1_e_er = ( e_edc_sc9_1 + e_tra_sc9_1 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
		conedsc9ra1_e_macadj = e_mac_1 / 100
		'conedsc9ra1_e_mfc = e_mfc_sc912_1 / 100
		'conedsc9ra1_e_cesss = e_cesss_1 / 100
		'conedsc9ra1_d_mscadj = d_msccap_sc9r1_1
		conedsc9ra1_d_dr_l5 = d_mc_1 + d_tra_mc_1
		conedsc9ra1_d_dr_l999 = d_o5_1 + d_tra_o5_1

		conedsc9ra1m_s_bppc = s_bppc_el_1
		conedsc9ra1m_s_cmc = s_cms_sc9_m_1
		conedsc9ra1m_e_er = ( e_edc_sc9_1 + e_tra_sc9_1 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
		conedsc9ra1m_e_macadj = e_mac_1 / 100
		'conedsc9ra1m_e_mfc = e_mfc_sc912_1 / 100
		'conedsc9ra1m_e_cesss = e_cesss_1 / 100
		'conedsc9ra1m_d_mscadj = d_msccap_sc9r1_1
		conedsc9ra1m_d_dr_l5 = d_mc_1 + d_tra_mc_1
		conedsc9ra1m_d_dr_l999 = d_o5_1 + d_tra_o5_1
		
		conedsc9r2_s_bppc = s_bppc_el_1
		conedsc9r2_s_cmc = s_cms_sc9_2
		conedsc9r2_e_er = ( e_edc_sc9_2 + e_tra_sc9_2 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
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
		conedsc9r2_d_dr_p1800 = d_mf86_sc9_2 + d_tra_mf86_sc9_2
		conedsc9r2_d_dr_p2200 = d_mf810_sc9_2 + d_tra_mf810_sc9_2
		conedsc9r2_d_dr_p2359 = d_all_sc9_2 + d_tra_all_sc9_2
		
		conedsc9r2m_s_bppc = s_bppc_el_1
		conedsc9r2m_s_cmc = s_cms_sc9_2
		conedsc9r2m_e_er = ( e_edc_sc9_2 + e_tra_sc9_2 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
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
		conedsc9r2m_d_dr_p1800 = d_mf86_sc9_2 + d_tra_mf86_sc9_2
		conedsc9r2m_d_dr_p2200 = d_mf810_sc9_2 + d_tra_mf810_sc9_2
		conedsc9r2m_d_dr_p2359 = d_all_sc9_2 + d_tra_all_sc9_2
	
		conedsc9ra2_s_bppc = s_bppc_el_1
		conedsc9ra2_s_cmc = s_cms_sc9_2
		conedsc9ra2_e_er = ( e_edc_sc9_2 + e_tra_sc9_3 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
		conedsc9ra2_e_er_p759  = conedsc9ra2_e_er
		conedsc9ra2_e_er_p1800 = conedsc9ra2_e_er
		conedsc9ra2_e_er_p2200 = conedsc9ra2_e_er
		conedsc9ra2_e_er_p2359 = conedsc9ra2_e_er
		conedsc9ra2_e_macadj = e_mac_1 / 100
		conedsc9ra2_e_macadj_p759  = conedsc9ra2_e_macadj
		conedsc9ra2_e_macadj_p1800 = conedsc9ra2_e_macadj
		conedsc9ra2_e_macadj_p2200 = conedsc9ra2_e_macadj
		conedsc9ra2_e_macadj_p2359 = conedsc9ra2_e_macadj		
		'conedsc9ra2_e_mfc = e_mfc_sc912_1 / 100
		'conedsc9ra2_e_cesss = e_cesss_1 / 100
		'conedsc9ra2_d_mscadj_p2200 = d_msccap_m_1
		conedsc9ra2_d_dr_p1800 = d_mf86_sc9_2 + d_tra_mf86_sc9_2
		conedsc9ra2_d_dr_p2200 = d_mf810_sc9_2 + d_tra_mf810_sc9_2
		conedsc9ra2_d_dr_p2359 = d_all_sc9_2 + d_tra_all_sc9_2

		conedsc9ra3_s_bppc = s_bppc_el_1
		conedsc9ra3_s_cmc = s_cms_sc9_3
		conedsc9ra3_e_er = ( e_edc_sc9_3 + e_tra_sc9_3 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
		conedsc9ra3_e_er_p759  = conedsc9ra3_e_er
		conedsc9ra3_e_er_p1800 = conedsc9ra3_e_er
		conedsc9ra3_e_er_p2200 = conedsc9ra3_e_er
		conedsc9ra3_e_er_p2359 = conedsc9ra3_e_er
		conedsc9ra3_e_macadj = e_mac_1 / 100
		conedsc9ra3_e_macadj_p759  = conedsc9ra3_e_macadj
		conedsc9ra3_e_macadj_p1800 = conedsc9ra3_e_macadj
		conedsc9ra3_e_macadj_p2200 = conedsc9ra3_e_macadj
		conedsc9ra3_e_macadj_p2359 = conedsc9ra3_e_macadj		
		'conedsc9ra3_e_mfc = e_mfc_sc912_1 / 100
		'conedsc9ra3_e_cesss = e_cesss_1 / 100
		'conedsc9ra3_d_mscadj_p2200 = d_msccap_m_1
		conedsc9ra3_d_dr_p1800 = d_mf86_sc9_3 + d_tra_mf86_sc9_3
		conedsc9ra3_d_dr_p2200 = d_mf810_sc9_3 + d_tra_mf810_sc9_3
		conedsc9ra3_d_dr_p2359 = d_all_sc9_3 + d_tra_all_sc9_3

		conedsc9ra3m_s_bppc = s_bppc_el_1
		conedsc9ra3m_s_cmc = s_cms_sc9_m_3
		conedsc9ra3m_e_er = ( e_edc_sc9_3 + e_tra_sc9_3 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
		conedsc9ra3m_e_er_p759  = conedsc9ra3m_e_er
		conedsc9ra3m_e_er_p1800 = conedsc9ra3m_e_er
		conedsc9ra3m_e_er_p2200 = conedsc9ra3m_e_er
		conedsc9ra3m_e_er_p2359 = conedsc9ra3m_e_er
		conedsc9ra3m_e_macadj = e_mac_1 / 100
		conedsc9ra3m_e_macadj_p759  = conedsc9ra3m_e_macadj
		conedsc9ra3m_e_macadj_p1800 = conedsc9ra3m_e_macadj
		conedsc9ra3m_e_macadj_p2200 = conedsc9ra3m_e_macadj
		conedsc9ra3m_e_macadj_p2359 = conedsc9ra3m_e_macadj		
		'conedsc9ra3m_e_mfc = e_mfc_sc912_1 / 100
		'conedsc9ra3m_e_cesss = e_cesss_1 / 100
		'conedsc9ra3m_d_mscadj_p2200 = d_msccap_m_1
		conedsc9ra3m_d_dr_p1800 = d_mf86_sc9_3 + d_tra_mf86_sc9_3
		conedsc9ra3m_d_dr_p2200 = d_mf810_sc9_3 + d_tra_mf810_sc9_3
		conedsc9ra3m_d_dr_p2359 = d_all_sc9_3 + d_tra_all_sc9_3

		conedsc12ra2_s_bppc = s_bppc_el_1
		conedsc12ra2_s_cmc = s_cms_sc12_2
		conedsc12ra2_e_er = ( e_edc_sc12_2 + e_tra_sc12_2 + e_sbc_1 + e_rpsp_1 + e_psls_sc12_2 + e_rdm_sc12_2 + e_drs_sc12_2 + e_cesds_1) / 100
		conedsc12ra2_e_er_p759  = conedsc12ra2_e_er
		conedsc12ra2_e_er_p1800 = conedsc12ra2_e_er
		conedsc12ra2_e_er_p2200 = conedsc12ra2_e_er
		conedsc12ra2_e_er_p2359 = conedsc12ra2_e_er
		conedsc12ra2_e_macadj = e_mac_1 / 100
		conedsc12ra2_e_macadj_p759  = conedsc12ra2_e_macadj
		conedsc12ra2_e_macadj_p1800 = conedsc12ra2_e_macadj
		conedsc12ra2_e_macadj_p2200 = conedsc12ra2_e_macadj
		conedsc12ra2_e_macadj_p2359 = conedsc12ra2_e_macadj		
		'conedsc12ra2_e_mfc = e_mfc_sc912_1 / 100
		'conedsc12ra2_e_cesss = e_cesss_1 / 100
		'conedsc12ra2_d_mscadj_p2200 = d_msccap_m_1
		conedsc12ra2_d_dr_p1800 = d_mf86_sc12_2 + d_tra_mf86_sc12_2
		conedsc12ra2_d_dr_p2200 = d_mf810_sc12_2 + d_tra_mf810_sc12_2
		conedsc12ra2_d_dr_p2359 = d_all_sc12_2 + d_tra_all_sc12_2
		
		sc9r2_e_er = ( e_edc_sc9_2 + e_tra_sc9_2 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_mfc_sc912_1 + e_cesds_1 + e_cesss_1) / 100
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
		sc9r2_d_dr_p2200 = d_mf810_sc9_2 + d_tra_mf810_sc9_2
		sc9r2_d_dr_p2359 = d_all_sc9_2 + d_tra_all_sc9_2
		
		sc9ra1_e_er = ( e_edc_sc9_1 + e_tra_sc9_1 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
		sc9ra1_e_macadj = e_mac_1 / 100
		sc9ra1_d_dr_l5 = d_mc_1 + d_tra_mc_1
		sc9ra1_d_dr_l100 = d_o5_1 + d_tra_o5_1
		sc9ra1_d_dr_l999 = d_o5_1 + d_tra_o5_1
		
		sc9ra2_e_er = ( e_edc_sc9_2 + e_tra_sc9_2 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
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
		sc9ra2_d_dr_p2200 = d_mf810_sc9_2 + d_tra_mf810_sc9_2
		sc9ra2_d_dr_p2359 = d_all_sc9_2 + d_tra_all_sc9_2
		
		sc9ra3_e_er = ( e_edc_sc9_3 + e_tra_sc9_3 + e_sbc_1 + e_rpsp_1 + e_psls_sc9_1 + e_rdm_sc9_1 + e_drs_sc9_1 + e_cesds_1) / 100
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
		sc9ra3_d_dr_p2200 = d_mf810_sc9_3 + d_tra_mf810_sc9_3
		sc9ra3_d_dr_p2359 = d_all_sc9_3 + d_tra_all_sc9_3
		
		sc12ra2_e_er = ( e_edc_sc12_2 + e_tra_sc12_2 + e_sbc_1 + e_rpsp_1 + e_psls_sc12_2 + e_rdm_sc12_2 + e_drs_sc12_2 + e_cesds_1) / 100
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
		sc12ra2_d_dr_p2200 = d_mf810_sc12_2 + d_tra_mf810_sc12_2
		sc12ra2_d_dr_p2359 = d_all_sc12_2 + d_tra_all_sc12_2
	end if
%>
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:2359ice:2359ice"
xmlns:x="urn:schemas-microsoft-com:2359ice:excel"
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

<body link="#0563C1" vlink="#954F72" class=xl97>
	<form name="RateBuilderRates" method="post" action ="saveRates.asp">
		<table border=0 cellpadding=0 cellspacing=0 width=1689 style='border-collapse:
		 collapse;table-layout:fixed;width:1269pt'>
			 <col class=xl97 width=72 style='width:54pt'>
			 <col class=xl97 width=72 style='width:54pt'>
			 <col class=xl109 width=114 style='mso-width-source:userset;mso-width-alt:3648;
			 width:86pt'>
			 <col class=xl109 width=13 style='mso-width-source:userset;mso-width-alt:416;
			 width:10pt'>
			 <col class=xl109 width=78 style='mso-width-source:userset;mso-width-alt:2496;
			 width:59pt'>
			 <col class=xl109 width=241 span=2 style='mso-width-source:userset;mso-width-alt:
			 7712;width:181pt'>
			 <col class=xl109 width=40 style='mso-width-source:userset;mso-width-alt:1280;
			 width:30pt'>
			 <col class=xl97 width=241 style='mso-width-source:userset;mso-width-alt:7712;
			 width:181pt'>
			 <col class=xl97 width=73 style='mso-width-source:userset;mso-width-alt:2336;
			 width:55pt'>
			 <col class=xl97 width=72 span=7 style='width:54pt'>
			 <tr height=43 style='mso-height-source:userset;height:32.25pt'>
			  <td height=43 class=xl97 width=72 style='height:32.25pt;width:54pt'><a
			  name="Print_Area"></a></td>
			  <td colspan=9 class=xl148 width=1113 style='border-right:1.0pt solid black;
			  width:837pt'>Rate Builder</td>
			  <td class=xl97 width=72 style='width:54pt'></td>
			  <td class=xl97 width=72 style='width:54pt'></td>
			  <td class=xl97 width=72 style='width:54pt'></td>
			  <td class=xl97 width=72 style='width:54pt'></td>
			  <td class=xl97 width=72 style='width:54pt'></td>
			  <td class=xl97 width=72 style='width:54pt'></td>
			  <td class=xl97 width=72 style='width:54pt'></td>
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
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=31 style='mso-height-source:userset;height:23.25pt'>
			  <td height=31 class=xl97 style='height:23.25pt'></td>
			  <td class=xl97></td>
			  <td class=xl109></td>
			  <td class=xl109></td>
			  <td class=xl109></td>
			  <td class=xl109></td>
			  <td class=xl109></td>
			  <td colspan=2 class=xl189 style='border-right:1.0pt solid black'><%= monthname(month, true) & " " & year %></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
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
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=20 style='height:15.0pt'>
			  <td height=20 class=xl97 style='height:15.0pt'></td>
			  <td class=xl99>&nbsp;</td>
			  <td class=xl110>&nbsp;</td>
			  <td class=xl110>&nbsp;</td>
			  <td class=xl110>&nbsp;</td>
			  <td class=xl110>&nbsp;</td>
			  <td class=xl110>&nbsp;</td>
			  <td class=xl110>&nbsp;</td>
			  <td class=xl100>&nbsp;</td>
			  <td class=xl115>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td colspan=7 class=xl188>Computed Rates</td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=20 style='height:15.0pt'>
			  <td height=20 class=xl97 style='height:15.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl109></td>
			  <td class=xl109></td>
			  <td class=xl109></td>
			  <td class=xl109></td>
			  <td class=xl109></td>
			  <td class=xl109></td>
			  <td class=xl97></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>

			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl142>ConEd SC9R1</td>
			  <td class=xl118>&nbsp;</td>
			  <td class=xl120>Static</td>
			  <td class=xl121>Billing and Payment Processing Charge</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl126><%= conedsc9r1_s_bppc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>ConEd Meter Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_s_cmc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Energy</td>
			  <td class=xl123>Energy Rate</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_er %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>MAC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_macadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Merchant Function Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_mfc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Clean Energy Standard Supply Surcharge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_cesss %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Demand</td>
			  <td class=xl123>MSC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_mscadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>0-5</td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_dr_l5 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl144>&nbsp;</td>
			  <td class=xl119>&nbsp;</td>
			  <td class=xl124>&nbsp;</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl125>6-99999999999</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl128><%= conedsc9r1_d_dr_l999 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl145></td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl129></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>

			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl142>ConEd SC9R1</br> Rider M</td>
			  <td class=xl118>&nbsp;</td>
			  <td class=xl120>Static</td>
			  <td class=xl121>Billing and Payment Processing Charge</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl126><%= conedsc9r1_s_bppc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>ConEd Meter Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_s_cmc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Energy</td>
			  <td class=xl123>Energy Rate</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_er %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>MAC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_macadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Merchant Function Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_mfc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Clean Energy Standard Supply Surcharge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_cesss %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Demand</td>
			  <td class=xl123>MSC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_mscadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>0-5</td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_dr_l5 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl144>&nbsp;</td>
			  <td class=xl119>&nbsp;</td>
			  <td class=xl124>&nbsp;</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl125>6-99999999999</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl128><%= conedsc9r1_d_dr_l999 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl145></td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl129></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>

			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl142>ConEd SC9RA1</td>
			  <td class=xl118>&nbsp;</td>
			  <td class=xl120>Static</td>
			  <td class=xl121>Billing and Payment Processing Charge</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl126><%= conedsc9r1_s_bppc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>ConEd Meter Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_s_cmc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Energy</td>
			  <td class=xl123>Energy Rate</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_er %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>MAC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_macadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Merchant Function Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_mfc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Clean Energy Standard Supply Surcharge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_cesss %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Demand</td>
			  <td class=xl123>MSC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_mscadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>0-5</td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_dr_l5 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl144>&nbsp;</td>
			  <td class=xl119>&nbsp;</td>
			  <td class=xl124>&nbsp;</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl125>6-99999999999</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl128><%= conedsc9r1_d_dr_l999 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl145></td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl129></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>

			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl142>ConEd SC9RA1</br> Rider M</td>
			  <td class=xl118>&nbsp;</td>
			  <td class=xl120>Static</td>
			  <td class=xl121>Billing and Payment Processing Charge</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl126><%= conedsc9r1_s_bppc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>ConEd Meter Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_s_cmc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Energy</td>
			  <td class=xl123>Energy Rate</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_er %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>MAC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_macadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Merchant Function Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_mfc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Clean Energy Standard Supply Surcharge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_cesss %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Demand</td>
			  <td class=xl123>MSC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_mscadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>0-5</td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_dr_l5 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl144>&nbsp;</td>
			  <td class=xl119>&nbsp;</td>
			  <td class=xl124>&nbsp;</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl125>6-99999999999</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl128><%= conedsc9r1_d_dr_l999 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl145></td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl129></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>

			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl142>ConEd SC9R2</td>
			  <td class=xl118>&nbsp;</td>
			  <td class=xl120>Static</td>
			  <td class=xl121>Billing and Payment Processing Charge</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl126><%= conedsc9r1_s_bppc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>ConEd Meter Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_s_cmc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Energy</td>
			  <td class=xl123>Energy Rate</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_er %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>MAC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_macadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Merchant Function Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_mfc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Clean Energy Standard Supply Surcharge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_cesss %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Demand</td>
			  <td class=xl123>MSC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_mscadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>0-5</td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_dr_l5 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl144>&nbsp;</td>
			  <td class=xl119>&nbsp;</td>
			  <td class=xl124>&nbsp;</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl125>6-99999999999</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl128><%= conedsc9r1_d_dr_l999 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl145></td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl129></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>

			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl142>ConEd SC9R2</br> Rider M</td>
			  <td class=xl118>&nbsp;</td>
			  <td class=xl120>Static</td>
			  <td class=xl121>Billing and Payment Processing Charge</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl126><%= conedsc9r1_s_bppc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>ConEd Meter Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_s_cmc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Energy</td>
			  <td class=xl123>Energy Rate</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_er %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>MAC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_macadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Merchant Function Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_mfc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Clean Energy Standard Supply Surcharge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_cesss %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Demand</td>
			  <td class=xl123>MSC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_mscadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>0-5</td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_dr_l5 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl144>&nbsp;</td>
			  <td class=xl119>&nbsp;</td>
			  <td class=xl124>&nbsp;</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl125>6-99999999999</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl128><%= conedsc9r1_d_dr_l999 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl145></td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl129></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl142>ConEd SC9RA2</td>
			  <td class=xl118>&nbsp;</td>
			  <td class=xl120>Static</td>
			  <td class=xl121>Billing and Payment Processing Charge</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl126><%= conedsc9r1_s_bppc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>ConEd Meter Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_s_cmc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Energy</td>
			  <td class=xl123>Energy Rate</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_er %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>MAC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_macadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Merchant Function Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_mfc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Clean Energy Standard Supply Surcharge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_cesss %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Demand</td>
			  <td class=xl123>MSC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_mscadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>0-5</td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_dr_l5 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl144>&nbsp;</td>
			  <td class=xl119>&nbsp;</td>
			  <td class=xl124>&nbsp;</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl125>6-99999999999</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl128><%= conedsc9r1_d_dr_l999 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl145></td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl129></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl142>ConEd SC9RA3</td>
			  <td class=xl118>&nbsp;</td>
			  <td class=xl120>Static</td>
			  <td class=xl121>Billing and Payment Processing Charge</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl126><%= conedsc9r1_s_bppc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>ConEd Meter Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_s_cmc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Energy</td>
			  <td class=xl123>Energy Rate</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_er %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>MAC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_macadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Merchant Function Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_mfc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Clean Energy Standard Supply Surcharge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_cesss %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Demand</td>
			  <td class=xl123>MSC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_mscadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>0-5</td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_dr_l5 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl144>&nbsp;</td>
			  <td class=xl119>&nbsp;</td>
			  <td class=xl124>&nbsp;</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl125>6-99999999999</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl128><%= conedsc9r1_d_dr_l999 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl145></td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl129></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl142>ConEd SC9RA3</br> Rider M</td>
			  <td class=xl118>&nbsp;</td>
			  <td class=xl120>Static</td>
			  <td class=xl121>Billing and Payment Processing Charge</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl126><%= conedsc9r1_s_bppc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>ConEd Meter Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_s_cmc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Energy</td>
			  <td class=xl123>Energy Rate</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_er %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>MAC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_macadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Merchant Function Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_mfc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Clean Energy Standard Supply Surcharge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_cesss %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Demand</td>
			  <td class=xl123>MSC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_mscadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>0-5</td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_dr_l5 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl144>&nbsp;</td>
			  <td class=xl119>&nbsp;</td>
			  <td class=xl124>&nbsp;</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl125>6-99999999999</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl128><%= conedsc9r1_d_dr_l999 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl145></td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl129></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl142>ConEd SC12RA2</td>
			  <td class=xl118>&nbsp;</td>
			  <td class=xl120>Static</td>
			  <td class=xl121>Billing and Payment Processing Charge</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl126><%= conedsc9r1_s_bppc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>ConEd Meter Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_s_cmc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Energy</td>
			  <td class=xl123>Energy Rate</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_er %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>MAC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_macadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Merchant Function Charge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_mfc %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Clean Energy Standard Supply Surcharge</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_e_cesss %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Demand</td>
			  <td class=xl123>MSC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_mscadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>0-5</td>
			  <td class=xl123></td>
			  <td class=xl127><%= conedsc9r1_d_dr_l5 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl144>&nbsp;</td>
			  <td class=xl119>&nbsp;</td>
			  <td class=xl124>&nbsp;</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl125>6-99999999999</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl128><%= conedsc9r1_d_dr_l999 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl145></td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl129></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl142>SC9R1</td>
			  <td class=xl118>&nbsp;</td>
			  <td class=xl120>Energy<span style='mso-spacerun:yes'> </span></td>
			  <td class=xl121>Energy Rate</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl126><%= sc9r1_e_er %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>MAC Adj Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc9r1_e_macadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Demand</td>
			  <td class=xl123>MSC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc9r1_d_mscadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>0-5</td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc9r1_d_dr_l5 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123></td>
			  <td class=xl123>6-100</td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc9r1_d_dr_l100 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl144>&nbsp;</td>
			  <td class=xl119>&nbsp;</td>
			  <td class=xl124>&nbsp;</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl125>101-99999999999</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl128><%= sc9r1_d_dr_l999 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl145></td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl129></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>			 
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl142>SC9R2</td>
			  <td class=xl118>&nbsp;</td>
			  <td class=xl120>Energy<span style='mso-spacerun:yes'> </span></td>
			  <td class=xl121>Energy Rate<span style='mso-spacerun:yes'> </span></td>
			  <td class=xl121>all 4 rate peaks</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl126><%= sc9r2_e_er %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97 colspan=3 style='mso-ignore:colspan'>Summer: June-September</td>
			  <td colspan=3 style='mso-ignore:colspan'>Winter: other 8 months</td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>MAC Adj Factor</td>
			  <td class=xl123>all 4 rate peaks</td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc9r2_e_macadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl116 colspan=6 style='mso-ignore:colspan'>All rate peaks of this
			  rate's entry need to choose winter or summer</td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Demand</td>
			  <td class=xl123>MSC Adj. Factor</td>
			  <td class=xl123>Peak(8-10)(Mo-Fr 800-2200)</td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc9r2_d_mscadj_p2200 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>Peak(8-6)(Mo-Fr 800-1800)</td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc9r2_d_dr_p1800 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>Peak(8-10)(Mo-Fr 800-2200)</td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc9r2_d_dr_p2200 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl144>&nbsp;</td>
			  <td class=xl119>&nbsp;</td>
			  <td class=xl124>&nbsp;</td>
			  <td class=xl125>Demand Rate</td>
			  <td class=xl125>2359 Peak (Weekend) (Sa-Su 0-2359)</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl128><%= sc9r2_d_dr_p2359 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl145></td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl129></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl142>SC9RA1</td>
			  <td class=xl118>&nbsp;</td>
			  <td class=xl120>Energy<span style='mso-spacerun:yes'> </span></td>
			  <td class=xl121>Energy Rate</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl126><%= sc9ra1_e_er %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>MAC Adj. Factor</td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc9ra1_e_macadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Demand</td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>0-5</td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc9ra1_d_dr_l5 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123></td>
			  <td class=xl123>6-100</td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc9ra1_d_dr_l100 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl144>&nbsp;</td>
			  <td class=xl119>&nbsp;</td>
			  <td class=xl124>&nbsp;</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl125>101-99999999999</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl128><%= sc9ra1_d_dr_l999 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl145></td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl129></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl142>SC9RA2</td>
			  <td class=xl118>&nbsp;</td>
			  <td class=xl120>Energy<span style='mso-spacerun:yes'> </span></td>
			  <td class=xl121>Energy Rate<span style='mso-spacerun:yes'> </span></td>
			  <td class=xl121>all 4 rate peaks</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl126><%= sc9ra2_e_er %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97 colspan=3 style='mso-ignore:colspan'>Summer: June-September</td>
			  <td colspan=3 style='mso-ignore:colspan'>Winter: other 8 months</td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>MAC Adj Factor</td>
			  <td class=xl123>all 4 rate peaks</td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc9ra2_e_macadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl116 colspan=6 style='mso-ignore:colspan'>All rate peaks of this
			  rate's entry need to choose winter or summer</td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Demand</td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>Peak(8-6)(Mo-Fr 800-1800)</td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc9ra2_d_dr_p1800 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>Peak(8-10)(Mo-Fr 800-2200)</td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc9ra2_d_dr_p2200 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl144>&nbsp;</td>
			  <td class=xl119>&nbsp;</td>
			  <td class=xl124>&nbsp;</td>
			  <td class=xl125>Demand Rate</td>
			  <td class=xl125>2359 Peak (Weekend) (Sa-Su 0-2359)</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl128><%= sc9ra2_d_dr_p2359 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl145></td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl129></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl142>SC9RA3</td>
			  <td class=xl118>&nbsp;</td>
			  <td class=xl120>Energy<span style='mso-spacerun:yes'> </span></td>
			  <td class=xl121>Energy Rate<span style='mso-spacerun:yes'> </span></td>
			  <td class=xl121>all 4 rate peaks</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl126><%= sc9ra3_e_er %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97 colspan=3 style='mso-ignore:colspan'>Summer: June-September</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>MAC Adj Factor</td>
			  <td class=xl123>all 4 rate peaks</td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc9ra3_e_macadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl116 colspan=6 style='mso-ignore:colspan'>All rate peaks of this
			  rate's entry need to choose winter or summer</td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Demand</td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>Peak(8-6)(Mo-Fr 800-1800)</td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc9ra3_d_dr_p1800 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>Peak(8-10)(Mo-Fr 800-2200)</td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc9ra3_d_dr_p2200 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl144>&nbsp;</td>
			  <td class=xl119>&nbsp;</td>
			  <td class=xl124>&nbsp;</td>
			  <td class=xl125>Demand Rate</td>
			  <td class=xl125>2359 Peak (Weekend) (Sa-Su 0-2359)</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl128><%= sc9ra3_d_dr_p2359 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl145></td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl123></td>
			  <td class=xl129></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl142>SC12RA2</td>
			  <td class=xl118>&nbsp;</td>
			  <td class=xl120>Energy<span style='mso-spacerun:yes'> </span></td>
			  <td class=xl121>Energy Rate<span style='mso-spacerun:yes'> </span></td>
			  <td class=xl121>all 4 rate peaks</td>
			  <td class=xl121>&nbsp;</td>
			  <td class=xl126><%= sc12ra2_e_er %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97 colspan=3 style='mso-ignore:colspan'>Summer: June-September</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>MAC Adj Factor</td>
			  <td class=xl123>all 4 rate peaks</td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc12ra2_e_macadj %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl116 colspan=6 style='mso-ignore:colspan'>All rate peaks of this
			  rate's entry need to choose winter or summer</td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122>Demand</td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>Peak(8-6)(Mo-Fr 800-1800)</td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc12ra2_d_dr_p1800 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl143>&nbsp;</td>
			  <td></td>
			  <td class=xl122></td>
			  <td class=xl123>Demand Rate</td>
			  <td class=xl123>Peak(8-10)(Mo-Fr 800-2200)</td>
			  <td class=xl123></td>
			  <td class=xl127><%= sc12ra2_d_dr_p2200 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=24 style='height:18.0pt'>
			  <td height=24 class=xl97 style='height:18.0pt'></td>
			  <td class=xl95>&nbsp;</td>
			  <td class=xl144>&nbsp;</td>
			  <td class=xl119>&nbsp;</td>
			  <td class=xl124>&nbsp;</td>
			  <td class=xl125>Demand Rate</td>
			  <td class=xl125>2359 Peak (Weekend) (Sa-Su 0-2359)</td>
			  <td class=xl125>&nbsp;</td>
			  <td class=xl128><%= sc12ra2_d_dr_p2359 %></td>
			  <td class=xl96>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
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
			  <td class=xl105>&nbsp;</td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
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
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
			  <td height=20 class=xl97 style='height:15.0pt'></td>
			  <td class=xl97></td>
			  <td colspan=3 rowspan=3 class=xl179 style='border-right:1.0pt solid black;
			  border-bottom:1.0pt solid black'>Edit Components</td>
			  <td class=xl117></td>
			  <td colspan=3 rowspan=3 class=xl170 style='border-right:1.0pt solid black;
			  border-bottom:1.0pt solid black'><input type="submit" name ="action" value="Save Rate Values" class="standard" /></td>
			  <td></td>
			  <td></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=20 style='mso-height-source:userset;height:15.0pt'>
			  <td height=20 class=xl97 style='height:15.0pt'></td>
			  <td class=xl97></td>
			  <td class=xl117></td>
			  <td></td>
			  <td></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <tr height=21 style='mso-height-source:userset;height:15.75pt'>
			  <td height=21 class=xl97 style='height:15.75pt'></td>
			  <td class=xl97></td>
			  <td class=xl117></td>
			  <td></td>
			  <td></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			  <td class=xl97></td>
			 </tr>
			 <![if supportMisalignedColumns]>
			 <tr height=0 style='display:none'>
			  <td width=72 style='width:54pt'></td>
			  <td width=72 style='width:54pt'></td>
			  <td width=114 style='width:86pt'></td>
			  <td width=13 style='width:10pt'></td>
			  <td width=78 style='width:59pt'></td>
			  <td width=241 style='width:181pt'></td>
			  <td width=241 style='width:181pt'></td>
			  <td width=40 style='width:30pt'></td>
			  <td width=241 style='width:181pt'></td>
			  <td width=73 style='width:55pt'></td>
			  <td width=72 style='width:54pt'></td>
			  <td width=72 style='width:54pt'></td>
			  <td width=72 style='width:54pt'></td>
			  <td width=72 style='width:54pt'></td>
			  <td width=72 style='width:54pt'></td>
			  <td width=72 style='width:54pt'></td>
			  <td width=72 style='width:54pt'></td>
			 </tr>
			 <![endif]>
		</table>
		<input type="hidden" value="<%=rateperiod             %>" name="rateperiod"/>      
		<input type="hidden" value="<%=rbcid             %>" name="rbcid"/>        
		<input type="hidden" value="<%=rbid             %>" name="rbid"/>
		<input type="hidden" value="<%=sc9r1_e_er             %>" name="sc9r1_e_er"/>        
		<input type="hidden" value="<%=sc9r1_e_macadj         %>" name="sc9r1_e_macadj"/>
		<input type="hidden" value="<%=sc9r1_d_mscadj         %>" name="sc9r1_d_mscadj"/>
		<input type="hidden" value="<%=sc9r1_d_dr_l5           %>" name="sc9r1_d_dr_l5"/>
		<input type="hidden" value="<%=sc9r1_d_dr_l100         %>" name="sc9r1_d_dr_l100"/>
		<input type="hidden" value="<%=sc9r1_d_dr_l999         %>" name="sc9r1_d_dr_l999"/>
		<input type="hidden" value="<%=conedsc9r1_s_bppc      %>" name="conedsc9r1_s_bppc"/>
		<input type="hidden" value="<%=conedsc9r1_s_cmc       %>" name="conedsc9r1_s_cmc"/>
		<input type="hidden" value="<%=conedsc9r1_e_er        %>" name="conedsc9r1_e_er"/>
		<input type="hidden" value="<%=conedsc9r1_e_macadj    %>" name="conedsc9r1_e_macadj"/>
		<input type="hidden" value="<%=conedsc9r1_e_mfc       %>" name="conedsc9r1_e_mfc"/>
		<input type="hidden" value="<%=conedsc9r1_d_mscadj    %>" name="conedsc9r1_d_mscadj"/>
		<input type="hidden" value="<%=conedsc9r1_d_dr_l5      %>" name="conedsc9r1_d_dr_l5"/>
		<input type="hidden" value="<%=conedsc9r1_d_dr_l999    %>" name="conedsc9r1_d_dr_l999"/>
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
		<input type="hidden" value="<%=sc9ra1_d_dr_l5          %>" name="sc9ra1_d_dr_l5"/>
		<input type="hidden" value="<%=sc9ra1_d_dr_l100        %>" name="sc9ra1_d_dr_l100"/>
		<input type="hidden" value="<%=sc9ra1_d_dr_l999        %>" name="sc9ra1_d_dr_l999"/>
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

	</form>
</body>
</html>
