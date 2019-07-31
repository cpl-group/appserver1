<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/nonsecure.inc"-->
<%


			dim cnn1, rst1, rst2, rst3, insertSql, sql
			set cnn1 = server.createobject("ADODB.connection")
			set rst1 = server.createobject("ADODB.recordset")
			set rst2 = server.createobject("ADODB.recordset")
			set rst3 = server.createobject("ADODB.recordset")
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
				if val="" or isnull(val) then
					val = 0
				end if
				if IsNumeric(CStr(val)) then
					toNumb = cdbl(val)
				end if
			end function
			function linecharge(val)
				dim lc
				select case val
				case"er"
					lc ="Energy Rate"	
				case"dr"
					lc ="Demand Rate"
				case"mscadj"
					lc ="MSC Adj. Factor"
				case"macadj"
					lc ="MAC Adj. Factor"
				case"bppc"
					lc ="Billing and Payment processing charge"
				case"cmc"
					lc ="ConEd Meter Charge"
				case"mfc"
					lc ="Merchant Function Charge"
				case"cesss"
					lc = "Clean Energy Standard Supply Charge"
				case"dlms"
					lc = "Dynamic Load Management Surcharge"
				case"sbc"
					lc = "System Benefit Charge"
				case"rpd"
					lc = "Reactive-Power Demand"
				case"msccap"
					lc = "MSC-CAP"
				case"tsc"
					lc = "Tax Sur-credit"
				case"edc"
					lc = "Energy Delivery Charge"
				end select
				linecharge = lc
			end function
			function itemtype(val)
				dim t
				select case val
				case"e"
					t ="Energy"
				case"d"
					t ="Demand"
				case"s"
					t ="Static"
				end select
				itemtype = t
			end function
		%>
		<% 	
			dim rbid,rbcid,rbrid,rateperiod,createdBy,createdOn,modifiedBy,modifiedOn, utility

			dim sc9r1_e_er,  sc9r1_e_macadj,  sc9r1_d_mscadj,  sc9r1_d_dr_l05,  sc9r1_d_dr_l100,  sc9r1_d_dr_l999,  sc9r1_d_dlms,  conedsc9r1_s_bppc,  conedsc9r1_s_cmc,  conedsc9r1_e_er,  conedsc9r1_e_mfc,  conedsc9r1_d_dr_l05,  conedsc9r1_d_dr_l999,  conedsc9r1_e_cesss,  conedsc9r1_d_dlms,  sc9r2_e_er_p759,  sc9r2_e_er_p1800,  sc9r2_e_er_p2200,  sc9r2_e_er_p2359,  sc9r2_e_macadj_p759,  sc9r2_e_macadj_p1800,  sc9r2_e_macadj_p2200,  sc9r2_e_macadj_p2359,  sc9r2_d_mscadj_p2200,  sc9r2_d_dr_p1800,  sc9r2_d_dr_p2200,  sc9r2_d_dr_p2359,  sc9ra1_e_er,  sc9ra1_e_macadj,  sc9ra1_d_dr_l05,  sc9ra1_d_dr_l100,  sc9ra1_d_dr_l999,  sc9ra2_e_er_p759,  sc9ra2_e_er_p1800,  sc9ra2_e_er_p2200,  sc9ra2_e_er_p2359,  sc9ra2_e_macadj_p759,  sc9ra2_e_macadj_p1800,  sc9ra2_e_macadj_p2200,  sc9ra2_e_macadj_p2359,  sc9ra2_d_dr_p1800,  sc9ra2_d_dr_p2200,  sc9ra2_d_dr_p2359,  sc9ra3_e_er_p759,  sc9ra3_e_er_p1800,  sc9ra3_e_er_p2200,  sc9ra3_e_er_p2359,  sc9ra3_e_macadj_p759,  sc9ra3_e_macadj_p1800,  sc9ra3_e_macadj_p2200,  sc9ra3_e_macadj_p2359,  sc9ra3_d_dr_p1800,  sc9ra3_d_dr_p2200,  sc9ra3_d_dr_p2359,  sc12ra2_e_er_p759,  sc12ra2_e_er_p1800,  sc12ra2_e_er_p2200,  sc12ra2_e_er_p2359,  sc12ra2_e_macadj_p759,  sc12ra2_e_macadj_p1800,  sc12ra2_e_macadj_p2200,  sc12ra2_e_macadj_p2359,  sc12ra2_d_dr_p1800,  sc12ra2_d_dr_p2200,  sc12ra2_d_dr_p2359,  conedsc9r1m_s_bppc,  conedsc9r1m_s_cmc,  conedsc9r1m_e_er,  conedsc9r1m_e_mfc,  conedsc9r1m_d_mscadj,  conedsc9r1m_d_dr_l05,  conedsc9r1m_d_dr_l999,  conedsc9r1m_e_cesss,  conedsc9ra1_s_bppc,  conedsc9ra1_s_cmc,  conedsc9ra1_e_er,  conedsc9ra1_d_dr_l05,  conedsc9ra1_d_dr_l999,  conedsc9ra1_d_dlms,  conedsc9ra1m_s_bppc,  conedsc9ra1m_s_cmc,  conedsc9ra1m_e_er,  conedsc9ra1m_d_dr_l05,  conedsc9ra1m_d_dr_l999,  conedsc9r2_s_bppc,  conedsc9r2_s_cmc,  conedsc9r2_e_mfc,  conedsc9r2_e_cesss,  conedsc9r2m_s_bppc,  conedsc9r2m_s_cmc,  conedsc9r2m_e_mfc,  conedsc9ra2_s_bppc,  conedsc9ra2_s_cmc,  conedsc9ra3_s_bppc,  conedsc9ra3_s_cmc,  conedsc9ra3m_s_bppc,  conedsc9ra3m_s_cmc,  conedsc12ra2_s_bppc,  conedsc12ra2_s_cmc,  conedsc9r2_e_er_p1800,  conedsc9r2_e_er_p2200,  conedsc9r2_e_er_p2359,  conedsc9r2m_e_er_p1800,  conedsc9r2m_e_er_p2200,  conedsc9r2m_e_er_p2359,  conedsc9ra2_e_er_p1800,  conedsc9ra2_e_er_p2200,  conedsc9ra2_e_er_p2359,  conedsc9ra3_e_er_p1800,  conedsc9ra3_e_er_p2200,  conedsc9ra3_e_er_p2359,  conedsc9ra3m_e_er_p1800,  conedsc9ra3m_e_er_p2200,  conedsc9ra3m_e_er_p2359,  conedsc12ra2_e_er_p1800,  conedsc12ra2_e_er_p2200,  conedsc12ra2_e_er_p2359,  conedsc12ra2_d_dr_p1800,  conedsc12ra2_d_dr_p2200,  conedsc12ra2_d_dr_p2359,  conedsc9ra3m_d_dr_p1800,  conedsc9ra3m_d_dr_p2200,  conedsc9ra3m_d_dr_p2359,  conedsc9ra3_d_dr_p1800,  conedsc9ra3_d_dr_p2200,  conedsc9ra3_d_dr_p2359,  conedsc9ra2_d_dr_p1800,  conedsc9ra2_d_dr_p2200,  conedsc9ra2_d_dr_p2359,  conedsc9r2m_d_dr_p1800,  conedsc9r2m_d_dr_p2200,  conedsc9r2m_d_dr_p2359,  conedsc9r2m_e_cesss,  conedsc9r1m_d_dlms,  conedsc9ra1m_d_dlms,  sc9ra1_d_dlms,  conedsc12ra2_e_sbc_p1800,  conedsc12ra2_e_sbc_p2200,  conedsc12ra2_e_sbc_p2359,  conedsc9r2_e_sbc_p1800,  conedsc9r2_e_sbc_p2200,  conedsc9r2_e_sbc_p2359,  conedsc9r2_d_msccap,  conedsc9r1_d_msccap,  conedsc9r1_e_sbc,  conedsc9r1m_e_sbc,  conedsc9r1m_d_msccap,  conedsc9r2m_e_sbc_p1800,  conedsc9r2m_e_sbc_p2200,  conedsc9r2m_e_sbc_p2359,  conedsc9r2m_d_msccap_p2200,  conedsc9ra1_e_sbc,  conedsc9ra1m_e_sbc,  conedsc9ra2_e_sbc_p1800,  conedsc9ra2_e_sbc_p2200,  conedsc9ra2_e_sbc_p2359,  conedsc9ra3_e_sbc_p1800,  conedsc9ra3_e_sbc_p2200,  conedsc9ra3_e_sbc_p2359,  conedsc9ra3m_e_sbc_p1800,  conedsc9ra3m_e_sbc_p2200,  conedsc9ra3m_e_sbc_p2359, conedsc9r1_e_tsc, conedsc9r1_d_tsc, conedsc9r1m_e_tsc, conedsc9r1m_d_tsc, conedsc9ra1_e_tsc, conedsc9ra1_d_tsc, conedsc9ra1m_e_tsc, conedsc9ra1m_d_tsc, conedsc9r2_e_tsc, conedsc9r2_d_tsc, conedsc9r2m_e_tsc, conedsc9r2m_d_tsc, conedsc9ra2_e_tsc, conedsc9ra2_d_tsc, conedsc9ra3_e_tsc, conedsc9ra3_d_tsc, conedsc9ra3m_e_tsc, conedsc9ra3m_d_tsc, conedsc12ra2_e_tsc, conedsc12ra2_d_tsc,conedsc9r2m_d_msccap
			
			dim sc9r1_e_er_id,  sc9r1_e_macadj_id,  sc9r1_d_mscadj_id,  sc9r1_d_dr_l05_id,  sc9r1_d_dr_l100_id,  sc9r1_d_dr_l999_id,  sc9r1_d_dlms_id,  conedsc9r1_s_bppc_id,  conedsc9r1_s_cmc_id,  conedsc9r1_e_er_id,  conedsc9r1_e_mfc_id,  conedsc9r1_d_dr_l05_id,  conedsc9r1_d_dr_l999_id,  conedsc9r1_e_cesss_id,  conedsc9r1_d_dlms_id,  sc9r2_e_er_p759_id,  sc9r2_e_er_p1800_id,  sc9r2_e_er_p2200_id,  sc9r2_e_er_p2359_id,  sc9r2_e_macadj_p759_id,  sc9r2_e_macadj_p1800_id,  sc9r2_e_macadj_p2200_id,  sc9r2_e_macadj_p2359_id,  sc9r2_d_mscadj_p2200_id,  sc9r2_d_dr_p1800_id,  sc9r2_d_dr_p2200_id,  sc9r2_d_dr_p2359_id,  sc9ra1_e_er_id,  sc9ra1_e_macadj_id,  sc9ra1_d_dr_l05_id,  sc9ra1_d_dr_l100_id,  sc9ra1_d_dr_l999_id,  sc9ra2_e_er_p759_id,  sc9ra2_e_er_p1800_id,  sc9ra2_e_er_p2200_id,  sc9ra2_e_er_p2359_id,  sc9ra2_e_macadj_p759_id,  sc9ra2_e_macadj_p1800_id,  sc9ra2_e_macadj_p2200_id,  sc9ra2_e_macadj_p2359_id,  sc9ra2_d_dr_p1800_id,  sc9ra2_d_dr_p2200_id,  sc9ra2_d_dr_p2359_id,  sc9ra3_e_er_p759_id,  sc9ra3_e_er_p1800_id,  sc9ra3_e_er_p2200_id,  sc9ra3_e_er_p2359_id,  sc9ra3_e_macadj_p759_id,  sc9ra3_e_macadj_p1800_id,  sc9ra3_e_macadj_p2200_id,  sc9ra3_e_macadj_p2359_id,  sc9ra3_d_dr_p1800_id,  sc9ra3_d_dr_p2200_id,  sc9ra3_d_dr_p2359_id,  sc12ra2_e_er_p759_id,  sc12ra2_e_er_p1800_id,  sc12ra2_e_er_p2200_id,  sc12ra2_e_er_p2359_id,  sc12ra2_e_macadj_p759_id,  sc12ra2_e_macadj_p1800_id,  sc12ra2_e_macadj_p2200_id,  sc12ra2_e_macadj_p2359_id,  sc12ra2_d_dr_p1800_id,  sc12ra2_d_dr_p2200_id,  sc12ra2_d_dr_p2359_id,  conedsc9r1m_s_bppc_id,  conedsc9r1m_s_cmc_id,  conedsc9r1m_e_er_id,  conedsc9r1m_e_mfc_id,  conedsc9r1m_d_mscadj_id,  conedsc9r1m_d_dr_l05_id,  conedsc9r1m_d_dr_l999_id,  conedsc9r1m_e_cesss_id,  conedsc9ra1_s_bppc_id,  conedsc9ra1_s_cmc_id,  conedsc9ra1_e_er_id,  conedsc9ra1_d_dr_l05_id,  conedsc9ra1_d_dr_l999_id,  conedsc9ra1_d_dlms_id,  conedsc9ra1m_s_bppc_id,  conedsc9ra1m_s_cmc_id,  conedsc9ra1m_e_er_id,  conedsc9ra1m_d_dr_l05_id,  conedsc9ra1m_d_dr_l999_id,  conedsc9r2_s_bppc_id,  conedsc9r2_s_cmc_id,  conedsc9r2_e_mfc_id,  conedsc9r2_e_cesss_id,  conedsc9r2m_s_bppc_id,  conedsc9r2m_s_cmc_id,  conedsc9r2m_e_mfc_id,  conedsc9ra2_s_bppc_id,  conedsc9ra2_s_cmc_id,  conedsc9ra3_s_bppc_id,  conedsc9ra3_s_cmc_id,  conedsc9ra3m_s_bppc_id,  conedsc9ra3m_s_cmc_id,  conedsc12ra2_s_bppc_id,  conedsc12ra2_s_cmc_id,  conedsc9r2_e_er_p1800_id,  conedsc9r2_e_er_p2200_id,  conedsc9r2_e_er_p2359_id,  conedsc9r2m_e_er_p1800_id,  conedsc9r2m_e_er_p2200_id,  conedsc9r2m_e_er_p2359_id,  conedsc9ra2_e_er_p1800_id,  conedsc9ra2_e_er_p2200_id,  conedsc9ra2_e_er_p2359_id,  conedsc9ra3_e_er_p1800_id,  conedsc9ra3_e_er_p2200_id,  conedsc9ra3_e_er_p2359_id,  conedsc9ra3m_e_er_p1800_id,  conedsc9ra3m_e_er_p2200_id,  conedsc9ra3m_e_er_p2359_id,  conedsc12ra2_e_er_p1800_id,  conedsc12ra2_e_er_p2200_id,  conedsc12ra2_e_er_p2359_id,  conedsc12ra2_d_dr_p1800_id,  conedsc12ra2_d_dr_p2200_id,  conedsc12ra2_d_dr_p2359_id,  conedsc9ra3m_d_dr_p1800_id,  conedsc9ra3m_d_dr_p2200_id,  conedsc9ra3m_d_dr_p2359_id,  conedsc9ra3_d_dr_p1800_id,  conedsc9ra3_d_dr_p2200_id,  conedsc9ra3_d_dr_p2359_id,  conedsc9ra2_d_dr_p1800_id,  conedsc9ra2_d_dr_p2200_id,  conedsc9ra2_d_dr_p2359_id,  conedsc9r2m_d_dr_p1800_id,  conedsc9r2m_d_dr_p2200_id,  conedsc9r2m_d_dr_p2359_id,  conedsc9r2m_e_cesss_id,  conedsc9r1m_d_dlms_id,  conedsc9ra1m_d_dlms_id,  sc9ra1_d_dlms_id,  conedsc12ra2_e_sbc_p1800_id,  conedsc12ra2_e_sbc_p2200_id,  conedsc12ra2_e_sbc_p2359_id,  conedsc9r2_e_sbc_p1800_id,  conedsc9r2_e_sbc_p2200_id,  conedsc9r2_e_sbc_p2359_id,  conedsc9r2_d_msccap_id,  conedsc9r1_d_msccap_id,  conedsc9r1_e_sbc_id,  conedsc9r1m_e_sbc_id,  conedsc9r1m_d_msccap_id,  conedsc9r2m_e_sbc_p1800_id,  conedsc9r2m_e_sbc_p2200_id,  conedsc9r2m_e_sbc_p2359_id,  conedsc9r2m_d_msccap_p2200_id,  conedsc9ra1_e_sbc_id,  conedsc9ra1m_e_sbc_id,  conedsc9ra2_e_sbc_p1800_id,  conedsc9ra2_e_sbc_p2200_id,  conedsc9ra2_e_sbc_p2359_id,  conedsc9ra3_e_sbc_p1800_id,  conedsc9ra3_e_sbc_p2200_id,  conedsc9ra3_e_sbc_p2359_id,  conedsc9ra3m_e_sbc_p1800_id,  conedsc9ra3m_e_sbc_p2200_id,  conedsc9ra3m_e_sbc_p2359_id, conedsc9r1_e_tsc_id, conedsc9r1_d_tsc_id, conedsc9r1m_e_tsc_id, conedsc9r1m_d_tsc_id, conedsc9ra1_e_tsc_id, conedsc9ra1_d_tsc_id, conedsc9ra1m_e_tsc_id, conedsc9ra1m_d_tsc_id, conedsc9r2_e_tsc_id, conedsc9r2_d_tsc_id, conedsc9r2m_e_tsc_id, conedsc9r2m_d_tsc_id, conedsc9ra2_e_tsc_id, conedsc9ra2_d_tsc_id, conedsc9ra3_e_tsc_id, conedsc9ra3_d_tsc_id, conedsc9ra3m_e_tsc_id, conedsc9ra3m_d_tsc_id, conedsc12ra2_e_tsc_id, conedsc12ra2_d_tsc_id,conedsc9r2m_d_msccap_id

			
			
			
			utility = 2
			
			rbcid             =toNumb(request.form("rbcid"))         
			rbid             =toNumb(request.form("rbid"))

	sc9r1_e_er = tonumb(request.form("sc9r1_e_er"))
	sc9r1_e_macadj = tonumb(request.form("sc9r1_e_macadj"))
	sc9r1_d_mscadj = tonumb(request.form("sc9r1_d_mscadj"))
	sc9r1_d_dr_l05 = tonumb(request.form("sc9r1_d_dr_l05"))
	sc9r1_d_dr_l100 = tonumb(request.form("sc9r1_d_dr_l100"))
	sc9r1_d_dr_l999 = tonumb(request.form("sc9r1_d_dr_l999"))
	sc9r1_d_dlms = tonumb(request.form("sc9r1_d_dlms"))
	conedsc9r1_s_bppc = tonumb(request.form("conedsc9r1_s_bppc"))
	conedsc9r1_s_cmc = tonumb(request.form("conedsc9r1_s_cmc"))
	conedsc9r1_e_er = tonumb(request.form("conedsc9r1_e_er"))
	conedsc9r1_e_mfc = tonumb(request.form("conedsc9r1_e_mfc"))
	conedsc9r1_d_dr_l05 = tonumb(request.form("conedsc9r1_d_dr_l05"))
	conedsc9r1_d_dr_l999 = tonumb(request.form("conedsc9r1_d_dr_l999"))
	conedsc9r1_e_cesss = tonumb(request.form("conedsc9r1_e_cesss"))
	conedsc9r1_d_dlms = tonumb(request.form("conedsc9r1_d_dlms"))
	sc9r2_e_er_p759 = tonumb(request.form("sc9r2_e_er_p759"))
	sc9r2_e_er_p1800 = tonumb(request.form("sc9r2_e_er_p1800"))
	sc9r2_e_er_p2200 = tonumb(request.form("sc9r2_e_er_p2200"))
	sc9r2_e_er_p2359 = tonumb(request.form("sc9r2_e_er_p2359"))
	sc9r2_e_macadj_p759 = tonumb(request.form("sc9r2_e_macadj_p759"))
	sc9r2_e_macadj_p1800 = tonumb(request.form("sc9r2_e_macadj_p1800"))
	sc9r2_e_macadj_p2200 = tonumb(request.form("sc9r2_e_macadj_p2200"))
	sc9r2_e_macadj_p2359 = tonumb(request.form("sc9r2_e_macadj_p2359"))
	sc9r2_d_mscadj_p2200 = tonumb(request.form("sc9r2_d_mscadj_p2200"))
	sc9r2_d_dr_p1800 = tonumb(request.form("sc9r2_d_dr_p1800"))
	sc9r2_d_dr_p2200 = tonumb(request.form("sc9r2_d_dr_p2200"))
	sc9r2_d_dr_p2359 = tonumb(request.form("sc9r2_d_dr_p2359"))
	sc9ra1_e_er = tonumb(request.form("sc9ra1_e_er"))
	sc9ra1_e_macadj = tonumb(request.form("sc9ra1_e_macadj"))
	sc9ra1_d_dr_l05 = tonumb(request.form("sc9ra1_d_dr_l05"))
	sc9ra1_d_dr_l100 = tonumb(request.form("sc9ra1_d_dr_l100"))
	sc9ra1_d_dr_l999 = tonumb(request.form("sc9ra1_d_dr_l999"))
	sc9ra2_e_er_p759 = tonumb(request.form("sc9ra2_e_er_p759"))
	sc9ra2_e_er_p1800 = tonumb(request.form("sc9ra2_e_er_p1800"))
	sc9ra2_e_er_p2200 = tonumb(request.form("sc9ra2_e_er_p2200"))
	sc9ra2_e_er_p2359 = tonumb(request.form("sc9ra2_e_er_p2359"))
	sc9ra2_e_macadj_p759 = tonumb(request.form("sc9ra2_e_macadj_p759"))
	sc9ra2_e_macadj_p1800 = tonumb(request.form("sc9ra2_e_macadj_p1800"))
	sc9ra2_e_macadj_p2200 = tonumb(request.form("sc9ra2_e_macadj_p2200"))
	sc9ra2_e_macadj_p2359 = tonumb(request.form("sc9ra2_e_macadj_p2359"))
	sc9ra2_d_dr_p1800 = tonumb(request.form("sc9ra2_d_dr_p1800"))
	sc9ra2_d_dr_p2200 = tonumb(request.form("sc9ra2_d_dr_p2200"))
	sc9ra2_d_dr_p2359 = tonumb(request.form("sc9ra2_d_dr_p2359"))
	sc9ra3_e_er_p759 = tonumb(request.form("sc9ra3_e_er_p759"))
	sc9ra3_e_er_p1800 = tonumb(request.form("sc9ra3_e_er_p1800"))
	sc9ra3_e_er_p2200 = tonumb(request.form("sc9ra3_e_er_p2200"))
	sc9ra3_e_er_p2359 = tonumb(request.form("sc9ra3_e_er_p2359"))
	sc9ra3_e_macadj_p759 = tonumb(request.form("sc9ra3_e_macadj_p759"))
	sc9ra3_e_macadj_p1800 = tonumb(request.form("sc9ra3_e_macadj_p1800"))
	sc9ra3_e_macadj_p2200 = tonumb(request.form("sc9ra3_e_macadj_p2200"))
	sc9ra3_e_macadj_p2359 = tonumb(request.form("sc9ra3_e_macadj_p2359"))
	sc9ra3_d_dr_p1800 = tonumb(request.form("sc9ra3_d_dr_p1800"))
	sc9ra3_d_dr_p2200 = tonumb(request.form("sc9ra3_d_dr_p2200"))
	sc9ra3_d_dr_p2359 = tonumb(request.form("sc9ra3_d_dr_p2359"))
	sc12ra2_e_er_p759 = tonumb(request.form("sc12ra2_e_er_p759"))
	sc12ra2_e_er_p1800 = tonumb(request.form("sc12ra2_e_er_p1800"))
	sc12ra2_e_er_p2200 = tonumb(request.form("sc12ra2_e_er_p2200"))
	sc12ra2_e_er_p2359 = tonumb(request.form("sc12ra2_e_er_p2359"))
	sc12ra2_e_macadj_p759 = tonumb(request.form("sc12ra2_e_macadj_p759"))
	sc12ra2_e_macadj_p1800 = tonumb(request.form("sc12ra2_e_macadj_p1800"))
	sc12ra2_e_macadj_p2200 = tonumb(request.form("sc12ra2_e_macadj_p2200"))
	sc12ra2_e_macadj_p2359 = tonumb(request.form("sc12ra2_e_macadj_p2359"))
	sc12ra2_d_dr_p1800 = tonumb(request.form("sc12ra2_d_dr_p1800"))
	sc12ra2_d_dr_p2200 = tonumb(request.form("sc12ra2_d_dr_p2200"))
	sc12ra2_d_dr_p2359 = tonumb(request.form("sc12ra2_d_dr_p2359"))
	conedsc9r1m_s_bppc = tonumb(request.form("conedsc9r1m_s_bppc"))
	conedsc9r1m_s_cmc = tonumb(request.form("conedsc9r1m_s_cmc"))
	conedsc9r1m_e_er = tonumb(request.form("conedsc9r1m_e_er"))
	conedsc9r1m_e_mfc = tonumb(request.form("conedsc9r1m_e_mfc"))
	conedsc9r1m_d_mscadj = tonumb(request.form("conedsc9r1m_d_mscadj"))
	conedsc9r1m_d_msccap = tonumb(request.form("conedsc9r1m_d_msccap"))
	conedsc9r1m_d_dr_l05 = tonumb(request.form("conedsc9r1m_d_dr_l05"))
	conedsc9r1m_d_dr_l999 = tonumb(request.form("conedsc9r1m_d_dr_l999"))
	conedsc9r1m_e_cesss = tonumb(request.form("conedsc9r1m_e_cesss"))
	conedsc9ra1_s_bppc = tonumb(request.form("conedsc9ra1_s_bppc"))
	conedsc9ra1_s_cmc = tonumb(request.form("conedsc9ra1_s_cmc"))
	conedsc9ra1_e_er = tonumb(request.form("conedsc9ra1_e_er"))
	conedsc9ra1_d_dr_l05 = tonumb(request.form("conedsc9ra1_d_dr_l05"))
	conedsc9ra1_d_dr_l999 = tonumb(request.form("conedsc9ra1_d_dr_l999"))
	conedsc9ra1_d_dlms = tonumb(request.form("conedsc9ra1_d_dlms"))
	conedsc9ra1m_s_bppc = tonumb(request.form("conedsc9ra1m_s_bppc"))
	conedsc9ra1m_s_cmc = tonumb(request.form("conedsc9ra1m_s_cmc"))
	conedsc9ra1m_e_er = tonumb(request.form("conedsc9ra1m_e_er"))
	conedsc9ra1m_d_dr_l05 = tonumb(request.form("conedsc9ra1m_d_dr_l05"))
	conedsc9ra1m_d_dr_l999 = tonumb(request.form("conedsc9ra1m_d_dr_l999"))
	conedsc9r2_s_bppc = tonumb(request.form("conedsc9r2_s_bppc"))
	conedsc9r2_s_cmc = tonumb(request.form("conedsc9r2_s_cmc"))
	conedsc9r2_e_mfc = tonumb(request.form("conedsc9r2_e_mfc"))
	conedsc9r2_e_cesss = tonumb(request.form("conedsc9r2_e_cesss"))
	conedsc9r2m_s_bppc = tonumb(request.form("conedsc9r2m_s_bppc"))
	conedsc9r2m_d_msccap = tonumb(request.form("conedsc9r2m_d_msccap"))
	conedsc9r2m_s_cmc = tonumb(request.form("conedsc9r2m_s_cmc"))
	conedsc9r2m_e_mfc = tonumb(request.form("conedsc9r2m_e_mfc"))
	conedsc9ra2_s_bppc = tonumb(request.form("conedsc9ra2_s_bppc"))
	conedsc9ra2_s_cmc = tonumb(request.form("conedsc9ra2_s_cmc"))
	conedsc9ra3_s_bppc = tonumb(request.form("conedsc9ra3_s_bppc"))
	conedsc9ra3_s_cmc = tonumb(request.form("conedsc9ra3_s_cmc"))
	conedsc9ra3m_s_bppc = tonumb(request.form("conedsc9ra3m_s_bppc"))
	conedsc9ra3m_s_cmc = tonumb(request.form("conedsc9ra3m_s_cmc"))
	conedsc12ra2_s_bppc = tonumb(request.form("conedsc12ra2_s_bppc"))
	conedsc12ra2_s_cmc = tonumb(request.form("conedsc12ra2_s_cmc"))
	conedsc9r2_e_er_p1800 = tonumb(request.form("conedsc9r2_e_er_p1800"))
	conedsc9r2_e_er_p2200 = tonumb(request.form("conedsc9r2_e_er_p2200"))
	conedsc9r2_e_er_p2359 = tonumb(request.form("conedsc9r2_e_er_p2359"))
	conedsc9r2m_e_er_p1800 = tonumb(request.form("conedsc9r2m_e_er_p1800"))
	conedsc9r2m_e_er_p2200 = tonumb(request.form("conedsc9r2m_e_er_p2200"))
	conedsc9r2m_e_er_p2359 = tonumb(request.form("conedsc9r2m_e_er_p2359"))
	conedsc9ra2_e_er_p1800 = tonumb(request.form("conedsc9ra2_e_er_p1800"))
	conedsc9ra2_e_er_p2200 = tonumb(request.form("conedsc9ra2_e_er_p2200"))
	conedsc9ra2_e_er_p2359 = tonumb(request.form("conedsc9ra2_e_er_p2359"))
	conedsc9ra3_e_er_p1800 = tonumb(request.form("conedsc9ra3_e_er_p1800"))
	conedsc9ra3_e_er_p2200 = tonumb(request.form("conedsc9ra3_e_er_p2200"))
	conedsc9ra3_e_er_p2359 = tonumb(request.form("conedsc9ra3_e_er_p2359"))
	conedsc9ra3m_e_er_p1800 = tonumb(request.form("conedsc9ra3m_e_er_p1800"))
	conedsc9ra3m_e_er_p2200 = tonumb(request.form("conedsc9ra3m_e_er_p2200"))
	conedsc9ra3m_e_er_p2359 = tonumb(request.form("conedsc9ra3m_e_er_p2359"))
	conedsc12ra2_e_er_p1800 = tonumb(request.form("conedsc12ra2_e_er_p1800"))
	conedsc12ra2_e_er_p2200 = tonumb(request.form("conedsc12ra2_e_er_p2200"))
	conedsc12ra2_e_er_p2359 = tonumb(request.form("conedsc12ra2_e_er_p2359"))
	conedsc12ra2_d_dr_p1800 = tonumb(request.form("conedsc12ra2_d_dr_p1800"))
	conedsc12ra2_d_dr_p2200 = tonumb(request.form("conedsc12ra2_d_dr_p2200"))
	conedsc12ra2_d_dr_p2359 = tonumb(request.form("conedsc12ra2_d_dr_p2359"))
	conedsc9ra3m_d_dr_p1800 = tonumb(request.form("conedsc9ra3m_d_dr_p1800"))
	conedsc9ra3m_d_dr_p2200 = tonumb(request.form("conedsc9ra3m_d_dr_p2200"))
	conedsc9ra3m_d_dr_p2359 = tonumb(request.form("conedsc9ra3m_d_dr_p2359"))
	conedsc9ra3_d_dr_p1800 = tonumb(request.form("conedsc9ra3_d_dr_p1800"))
	conedsc9ra3_d_dr_p2200 = tonumb(request.form("conedsc9ra3_d_dr_p2200"))
	conedsc9ra3_d_dr_p2359 = tonumb(request.form("conedsc9ra3_d_dr_p2359"))
	conedsc9ra2_d_dr_p1800 = tonumb(request.form("conedsc9ra2_d_dr_p1800"))
	conedsc9ra2_d_dr_p2200 = tonumb(request.form("conedsc9ra2_d_dr_p2200"))
	conedsc9ra2_d_dr_p2359 = tonumb(request.form("conedsc9ra2_d_dr_p2359"))
	conedsc9r2m_d_dr_p1800 = tonumb(request.form("conedsc9r2m_d_dr_p1800"))
	conedsc9r2m_d_dr_p2200 = tonumb(request.form("conedsc9r2m_d_dr_p2200"))
	conedsc9r2m_d_dr_p2359 = tonumb(request.form("conedsc9r2m_d_dr_p2359"))
	conedsc9r2m_e_cesss = tonumb(request.form("conedsc9r2m_e_cesss"))
	conedsc9r1m_d_dlms = tonumb(request.form("conedsc9r1m_d_dlms"))
	conedsc9ra1m_d_dlms = tonumb(request.form("conedsc9ra1m_d_dlms"))
	sc9ra1_d_dlms = tonumb(request.form("sc9ra1_d_dlms"))
	conedsc12ra2_e_sbc_p1800 = tonumb(request.form("conedsc12ra2_e_sbc_p1800"))
	conedsc12ra2_e_sbc_p2200 = tonumb(request.form("conedsc12ra2_e_sbc_p2200"))
	conedsc12ra2_e_sbc_p2359 = tonumb(request.form("conedsc12ra2_e_sbc_p2359"))
	conedsc9r2_e_sbc_p1800 = tonumb(request.form("conedsc9r2_e_sbc_p1800"))
	conedsc9r2_e_sbc_p2200 = tonumb(request.form("conedsc9r2_e_sbc_p2200"))
	conedsc9r2_e_sbc_p2359 = tonumb(request.form("conedsc9r2_e_sbc_p2359"))
	conedsc9r2_d_msccap = tonumb(request.form("conedsc9r2_d_msccap"))
	conedsc9r1_d_msccap = tonumb(request.form("conedsc9r1_d_msccap"))
	conedsc9r1_e_sbc = tonumb(request.form("conedsc9r1_e_sbc"))
	conedsc9r1m_e_sbc = tonumb(request.form("conedsc9r1m_e_sbc"))
	conedsc9r2m_e_sbc_p1800 = tonumb(request.form("conedsc9r2m_e_sbc_p1800"))
	conedsc9r2m_e_sbc_p2200 = tonumb(request.form("conedsc9r2m_e_sbc_p2200"))
	conedsc9r2m_e_sbc_p2359 = tonumb(request.form("conedsc9r2m_e_sbc_p2359"))
	conedsc9r2m_d_msccap_p2200 = tonumb(request.form("conedsc9r2m_d_msccap_p2200"))
	conedsc9ra1_e_sbc = tonumb(request.form("conedsc9ra1_e_sbc"))
	conedsc9ra1m_e_sbc = tonumb(request.form("conedsc9ra1m_e_sbc"))
	conedsc9ra2_e_sbc_p1800 = tonumb(request.form("conedsc9ra2_e_sbc_p1800"))
	conedsc9ra2_e_sbc_p2200 = tonumb(request.form("conedsc9ra2_e_sbc_p2200"))
	conedsc9ra2_e_sbc_p2359 = tonumb(request.form("conedsc9ra2_e_sbc_p2359"))
	conedsc9ra3_e_sbc_p1800 = tonumb(request.form("conedsc9ra3_e_sbc_p1800"))
	conedsc9ra3_e_sbc_p2200 = tonumb(request.form("conedsc9ra3_e_sbc_p2200"))
	conedsc9ra3_e_sbc_p2359 = tonumb(request.form("conedsc9ra3_e_sbc_p2359"))
	conedsc9ra3m_e_sbc_p1800 = tonumb(request.form("conedsc9ra3m_e_sbc_p1800"))
	conedsc9ra3m_e_sbc_p2200 = tonumb(request.form("conedsc9ra3m_e_sbc_p2200"))
	conedsc9ra3m_e_sbc_p2359 = tonumb(request.form("conedsc9ra3m_e_sbc_p2359")) 
	conedsc9r1_e_tsc    =  tonumb(request.form("conedsc9r1_e_tsc"))  
	conedsc9r1_d_tsc    =  tonumb(request.form("conedsc9r1_d_tsc")) 
	conedsc9r1m_e_tsc   =  tonumb(request.form("conedsc9r1m_e_tsc")) 
	conedsc9r1m_d_tsc   =  tonumb(request.form("conedsc9r1m_d_tsc")) 
	conedsc9ra1_e_tsc   =  tonumb(request.form("conedsc9ra1_e_tsc")) 
	conedsc9ra1_d_tsc   =  tonumb(request.form("conedsc9ra1_d_tsc")) 
	conedsc9ra1m_e_tsc  =  tonumb(request.form("conedsc9ra1m_e_tsc")) 
	conedsc9ra1m_d_tsc  =  tonumb(request.form("conedsc9ra1m_d_tsc")) 
	conedsc9r2_e_tsc    =  tonumb(request.form("conedsc9r2_e_tsc")) 
	conedsc9r2_d_tsc    =  tonumb(request.form("conedsc9r2_d_tsc")) 
	conedsc9r2m_e_tsc   =  tonumb(request.form("conedsc9r2m_e_tsc")) 
	conedsc9r2m_d_tsc   =  tonumb(request.form("conedsc9r2m_d_tsc")) 
	conedsc9ra2_e_tsc   =  tonumb(request.form("conedsc9ra2_e_tsc")) 
	conedsc9ra2_d_tsc   =  tonumb(request.form("conedsc9ra2_d_tsc")) 
	conedsc9ra3_e_tsc   =  tonumb(request.form("conedsc9ra3_e_tsc")) 
	conedsc9ra3_d_tsc   =  tonumb(request.form("conedsc9ra3_d_tsc")) 
	conedsc9ra3m_e_tsc  =  tonumb(request.form("conedsc9ra3m_e_tsc")) 
	conedsc9ra3m_d_tsc  =  tonumb(request.form("conedsc9ra3m_d_tsc")) 
	conedsc12ra2_e_tsc   =  tonumb(request.form("conedsc12ra2_e_tsc")) 
	conedsc12ra2_d_tsc   =  tonumb(request.form("conedsc12ra2_d_tsc"))


response.write conedsc9r1m_d_msccap  & "</br>"
response.write request.form("conedsc9r1m_d_msccap") & "</br>"




			
		%>

		<% 
			sql = "select rbrid from ratebuilder where rbid=" & rbid
			rst1.open sql, cnn1
			if not rst1.eof then rbrid = rst1("rbrid") end if
			rst1.close
			
			if rbrid <> "" then
				insertsql = "UPDATE [dbo].[RateBuilderRates]  SET [sc9r1_e_er] = '"&  sc9r1_e_er &"' ,  [sc9r1_e_macadj] = '"&  sc9r1_e_macadj &"' ,  [sc9r1_d_mscadj] = '"&  sc9r1_d_mscadj &"' ,  [sc9r1_d_dr_l05] = '"&  sc9r1_d_dr_l05 &"' ,  [sc9r1_d_dr_l100] = '"&  sc9r1_d_dr_l100 &"' ,  [sc9r1_d_dr_l999] = '"&  sc9r1_d_dr_l999 &"' ,  [sc9r1_d_dlms] = '"&  sc9r1_d_dlms &"' ,  [conedsc9r1_s_bppc] = '"&  conedsc9r1_s_bppc &"' ,  [conedsc9r1_s_cmc] = '"&  conedsc9r1_s_cmc &"' ,  [conedsc9r1_e_er] = '"&  conedsc9r1_e_er &"' ,  [conedsc9r1_e_mfc] = '"&  conedsc9r1_e_mfc &"' ,  [conedsc9r1_d_dr_l05] = '"&  conedsc9r1_d_dr_l05 &"' ,  [conedsc9r1_d_dr_l999] = '"&  conedsc9r1_d_dr_l999 &"' ,  [conedsc9r1_e_cesss] = '"&  conedsc9r1_e_cesss &"' ,  [conedsc9r1_d_dlms] = '"&  conedsc9r1_d_dlms &"' ,  [sc9r2_e_er_p759] = '"&  sc9r2_e_er_p759 &"' ,  [sc9r2_e_er_p1800] = '"&  sc9r2_e_er_p1800 &"' ,  [sc9r2_e_er_p2200] = '"&  sc9r2_e_er_p2200 &"' ,  [sc9r2_e_er_p2359] = '"&  sc9r2_e_er_p2359 &"' ,  [sc9r2_e_macadj_p759] = '"&  sc9r2_e_macadj_p759 &"' ,  [sc9r2_e_macadj_p1800] = '"&  sc9r2_e_macadj_p1800 &"' ,  [sc9r2_e_macadj_p2200] = '"&  sc9r2_e_macadj_p2200 &"' ,  [sc9r2_e_macadj_p2359] = '"&  sc9r2_e_macadj_p2359 &"' ,  [sc9r2_d_mscadj_p2200] = '"&  sc9r2_d_mscadj_p2200 &"' ,  [sc9r2_d_dr_p1800] = '"&  sc9r2_d_dr_p1800 &"' ,  [sc9r2_d_dr_p2200] = '"&  sc9r2_d_dr_p2200 &"' ,  [sc9r2_d_dr_p2359] = '"&  sc9r2_d_dr_p2359 &"' ,  [sc9ra1_e_er] = '"&  sc9ra1_e_er &"' ,  [sc9ra1_e_macadj] = '"&  sc9ra1_e_macadj &"' ,  [sc9ra1_d_dr_l05] = '"&  sc9ra1_d_dr_l05 &"' ,  [sc9ra1_d_dr_l100] = '"&  sc9ra1_d_dr_l100 &"' ,  [sc9ra1_d_dr_l999] = '"&  sc9ra1_d_dr_l999 &"' ,  [sc9ra2_e_er_p759] = '"&  sc9ra2_e_er_p759 &"' ,  [sc9ra2_e_er_p1800] = '"&  sc9ra2_e_er_p1800 &"' ,  [sc9ra2_e_er_p2200] = '"&  sc9ra2_e_er_p2200 &"' ,  [sc9ra2_e_er_p2359] = '"&  sc9ra2_e_er_p2359 &"' ,  [sc9ra2_e_macadj_p759] = '"&  sc9ra2_e_macadj_p759 &"' ,  [sc9ra2_e_macadj_p1800] = '"&  sc9ra2_e_macadj_p1800 &"' ,  [sc9ra2_e_macadj_p2200] = '"&  sc9ra2_e_macadj_p2200 &"' ,  [sc9ra2_e_macadj_p2359] = '"&  sc9ra2_e_macadj_p2359 &"' ,  [sc9ra2_d_dr_p1800] = '"&  sc9ra2_d_dr_p1800 &"' ,  [sc9ra2_d_dr_p2200] = '"&  sc9ra2_d_dr_p2200 &"' ,  [sc9ra2_d_dr_p2359] = '"&  sc9ra2_d_dr_p2359 &"' ,  [sc9ra3_e_er_p759] = '"&  sc9ra3_e_er_p759 &"' ,  [sc9ra3_e_er_p1800] = '"&  sc9ra3_e_er_p1800 &"' ,  [sc9ra3_e_er_p2200] = '"&  sc9ra3_e_er_p2200 &"' ,  [sc9ra3_e_er_p2359] = '"&  sc9ra3_e_er_p2359 &"' ,  [sc9ra3_e_macadj_p759] = '"&  sc9ra3_e_macadj_p759 &"' ,  [sc9ra3_e_macadj_p1800] = '"&  sc9ra3_e_macadj_p1800 &"' ,  [sc9ra3_e_macadj_p2200] = '"&  sc9ra3_e_macadj_p2200 &"' ,  [sc9ra3_e_macadj_p2359] = '"&  sc9ra3_e_macadj_p2359 &"' ,  [sc9ra3_d_dr_p1800] = '"&  sc9ra3_d_dr_p1800 &"' ,  [sc9ra3_d_dr_p2200] = '"&  sc9ra3_d_dr_p2200 &"' ,  [sc9ra3_d_dr_p2359] = '"&  sc9ra3_d_dr_p2359 &"' ,  [sc12ra2_e_er_p759] = '"&  sc12ra2_e_er_p759 &"' ,  [sc12ra2_e_er_p1800] = '"&  sc12ra2_e_er_p1800 &"' ,  [sc12ra2_e_er_p2200] = '"&  sc12ra2_e_er_p2200 &"' ,  [sc12ra2_e_er_p2359] = '"&  sc12ra2_e_er_p2359 &"' ,  [sc12ra2_e_macadj_p759] = '"&  sc12ra2_e_macadj_p759 &"' ,  [sc12ra2_e_macadj_p1800] = '"&  sc12ra2_e_macadj_p1800 &"' ,  [sc12ra2_e_macadj_p2200] = '"&  sc12ra2_e_macadj_p2200 &"' ,  [sc12ra2_e_macadj_p2359] = '"&  sc12ra2_e_macadj_p2359 &"' ,  [sc12ra2_d_dr_p1800] = '"&  sc12ra2_d_dr_p1800 &"' ,  [sc12ra2_d_dr_p2200] = '"&  sc12ra2_d_dr_p2200 &"' ,  [sc12ra2_d_dr_p2359] = '"&  sc12ra2_d_dr_p2359 &"' ,  [conedsc9r1m_s_bppc] = '"&  conedsc9r1m_s_bppc &"' ,  [conedsc9r1m_s_cmc] = '"&  conedsc9r1m_s_cmc &"' ,  [conedsc9r1m_e_er] = '"&  conedsc9r1m_e_er &"' ,  [conedsc9r1m_e_mfc] = '"&  conedsc9r1m_e_mfc &"' ,  [conedsc9r1m_d_mscadj] = '"&  conedsc9r1m_d_mscadj &"' ,  [conedsc9r1m_d_dr_l05] = '"&  conedsc9r1m_d_dr_l05 &"' ,  [conedsc9r1m_d_dr_l999] = '"&  conedsc9r1m_d_dr_l999 &"' ,  [conedsc9r1m_e_cesss] = '"&  conedsc9r1m_e_cesss &"' ,  [conedsc9ra1_s_bppc] = '"&  conedsc9ra1_s_bppc &"' ,  [conedsc9ra1_s_cmc] = '"&  conedsc9ra1_s_cmc &"' ,  [conedsc9ra1_e_er] = '"&  conedsc9ra1_e_er &"' ,  [conedsc9ra1_d_dr_l05] = '"&  conedsc9ra1_d_dr_l05 &"' ,  [conedsc9ra1_d_dr_l999] = '"&  conedsc9ra1_d_dr_l999 &"' ,  [conedsc9ra1_d_dlms] = '"&  conedsc9ra1_d_dlms &"' ,  [conedsc9ra1m_s_bppc] = '"&  conedsc9ra1m_s_bppc &"' ,  [conedsc9ra1m_s_cmc] = '"&  conedsc9ra1m_s_cmc &"' ,  [conedsc9ra1m_e_er] = '"&  conedsc9ra1m_e_er &"' ,  [conedsc9ra1m_d_dr_l05] = '"&  conedsc9ra1m_d_dr_l05 &"' ,  [conedsc9ra1m_d_dr_l999] = '"&  conedsc9ra1m_d_dr_l999 &"' ,  [conedsc9r2_s_bppc] = '"&  conedsc9r2_s_bppc &"' ,  [conedsc9r2_s_cmc] = '"&  conedsc9r2_s_cmc &"' ,  [conedsc9r2_e_mfc] = '"&  conedsc9r2_e_mfc &"' ,  [conedsc9r2_e_cesss] = '"&  conedsc9r2_e_cesss &"' ,  [conedsc9r2m_s_bppc] = '"&  conedsc9r2m_s_bppc &"' ,  [conedsc9r2m_s_cmc] = '"&  conedsc9r2m_s_cmc &"' ,  [conedsc9r2m_e_mfc] = '"&  conedsc9r2m_e_mfc &"' ,  [conedsc9ra2_s_bppc] = '"&  conedsc9ra2_s_bppc &"' ,  [conedsc9ra2_s_cmc] = '"&  conedsc9ra2_s_cmc &"' ,  [conedsc9ra3_s_bppc] = '"&  conedsc9ra3_s_bppc &"' ,  [conedsc9ra3_s_cmc] = '"&  conedsc9ra3_s_cmc &"' ,  [conedsc9ra3m_s_bppc] = '"&  conedsc9ra3m_s_bppc &"' ,  [conedsc9ra3m_s_cmc] = '"&  conedsc9ra3m_s_cmc &"' ,  [conedsc12ra2_s_bppc] = '"&  conedsc12ra2_s_bppc &"' ,  [conedsc12ra2_s_cmc] = '"&  conedsc12ra2_s_cmc &"' ,  [conedsc9r2_e_er_p1800] = '"&  conedsc9r2_e_er_p1800 &"' ,  [conedsc9r2_e_er_p2200] = '"&  conedsc9r2_e_er_p2200 &"' ,  [conedsc9r2_e_er_p2359] = '"&  conedsc9r2_e_er_p2359 &"' ,  [conedsc9r2m_e_er_p1800] = '"&  conedsc9r2m_e_er_p1800 &"' ,  [conedsc9r2m_e_er_p2200] = '"&  conedsc9r2m_e_er_p2200 &"' ,  [conedsc9r2m_e_er_p2359] = '"&  conedsc9r2m_e_er_p2359 &"' ,  [conedsc9ra2_e_er_p1800] = '"&  conedsc9ra2_e_er_p1800 &"' ,  [conedsc9ra2_e_er_p2200] = '"&  conedsc9ra2_e_er_p2200 &"' ,  [conedsc9ra2_e_er_p2359] = '"&  conedsc9ra2_e_er_p2359 &"' ,  [conedsc9ra3_e_er_p1800] = '"&  conedsc9ra3_e_er_p1800 &"' ,  [conedsc9ra3_e_er_p2200] = '"&  conedsc9ra3_e_er_p2200 &"' ,  [conedsc9ra3_e_er_p2359] = '"&  conedsc9ra3_e_er_p2359 &"' ,  [conedsc9ra3m_e_er_p1800] = '"&  conedsc9ra3m_e_er_p1800 &"' ,  [conedsc9ra3m_e_er_p2200] = '"&  conedsc9ra3m_e_er_p2200 &"' ,  [conedsc9ra3m_e_er_p2359] = '"&  conedsc9ra3m_e_er_p2359 &"' ,  [conedsc12ra2_e_er_p1800] = '"&  conedsc12ra2_e_er_p1800 &"' ,  [conedsc12ra2_e_er_p2200] = '"&  conedsc12ra2_e_er_p2200 &"' ,  [conedsc12ra2_e_er_p2359] = '"&  conedsc12ra2_e_er_p2359 &"' ,  [conedsc12ra2_d_dr_p1800] = '"&  conedsc12ra2_d_dr_p1800 &"' ,  [conedsc12ra2_d_dr_p2200] = '"&  conedsc12ra2_d_dr_p2200 &"' ,  [conedsc12ra2_d_dr_p2359] = '"&  conedsc12ra2_d_dr_p2359 &"' ,  [conedsc9ra3m_d_dr_p1800] = '"&  conedsc9ra3m_d_dr_p1800 &"' ,  [conedsc9ra3m_d_dr_p2200] = '"&  conedsc9ra3m_d_dr_p2200 &"' ,  [conedsc9ra3m_d_dr_p2359] = '"&  conedsc9ra3m_d_dr_p2359 &"' ,  [conedsc9ra3_d_dr_p1800] = '"&  conedsc9ra3_d_dr_p1800 &"' ,  [conedsc9ra3_d_dr_p2200] = '"&  conedsc9ra3_d_dr_p2200 &"' ,  [conedsc9ra3_d_dr_p2359] = '"&  conedsc9ra3_d_dr_p2359 &"' ,  [conedsc9ra2_d_dr_p1800] = '"&  conedsc9ra2_d_dr_p1800 &"' ,  [conedsc9ra2_d_dr_p2200] = '"&  conedsc9ra2_d_dr_p2200 &"' ,  [conedsc9ra2_d_dr_p2359] = '"&  conedsc9ra2_d_dr_p2359 &"' ,  [conedsc9r2m_d_dr_p1800] = '"&  conedsc9r2m_d_dr_p1800 &"' ,  [conedsc9r2m_d_dr_p2200] = '"&  conedsc9r2m_d_dr_p2200 &"' ,  [conedsc9r2m_d_dr_p2359] = '"&  conedsc9r2m_d_dr_p2359 &"' ,  [conedsc9r2m_e_cesss] = '"&  conedsc9r2m_e_cesss &"' ,  [conedsc9r1m_d_dlms] = '"&  conedsc9r1m_d_dlms &"' ,  [conedsc9ra1m_d_dlms] = '"&  conedsc9ra1m_d_dlms &"' ,  [sc9ra1_d_dlms] = '"&  sc9ra1_d_dlms &"' ,  [conedsc12ra2_e_sbc_p1800] = '"&  conedsc12ra2_e_sbc_p1800 &"' ,  [conedsc12ra2_e_sbc_p2200] = '"&  conedsc12ra2_e_sbc_p2200 &"' ,  [conedsc12ra2_e_sbc_p2359] = '"&  conedsc12ra2_e_sbc_p2359 &"' ,  [conedsc9r2_e_sbc_p1800] = '"&  conedsc9r2_e_sbc_p1800 &"' ,  [conedsc9r2_e_sbc_p2200] = '"&  conedsc9r2_e_sbc_p2200 &"' ,  [conedsc9r2_e_sbc_p2359] = '"&  conedsc9r2_e_sbc_p2359 &"' ,  [conedsc9r2_d_msccap] = '"&  conedsc9r2_d_msccap &"' ,  [conedsc9r1_d_msccap] = '"&  conedsc9r1_d_msccap &"' ,  [conedsc9r1_e_sbc] = '"&  conedsc9r1_e_sbc &"' ,  [conedsc9r1m_e_sbc] = '"&  conedsc9r1m_e_sbc &"' ,  [conedsc9r1m_d_msccap] = '"&  conedsc9r1m_d_msccap &"' ,  [conedsc9r2m_e_sbc_p1800] = '"&  conedsc9r2m_e_sbc_p1800 &"' ,  [conedsc9r2m_e_sbc_p2200] = '"&  conedsc9r2m_e_sbc_p2200 &"' ,  [conedsc9r2m_e_sbc_p2359] = '"&  conedsc9r2m_e_sbc_p2359 &"' ,  [conedsc9r2m_d_msccap_p2200] = '"&  conedsc9r2m_d_msccap_p2200 &"' ,  [conedsc9ra1_e_sbc] = '"&  conedsc9ra1_e_sbc &"' ,  [conedsc9ra1m_e_sbc] = '"&  conedsc9ra1m_e_sbc &"' ,  [conedsc9ra2_e_sbc_p1800] = '"&  conedsc9ra2_e_sbc_p1800 &"' ,  [conedsc9ra2_e_sbc_p2200] = '"&  conedsc9ra2_e_sbc_p2200 &"' ,  [conedsc9ra2_e_sbc_p2359] = '"&  conedsc9ra2_e_sbc_p2359 &"' ,  [conedsc9ra3_e_sbc_p1800] = '"&  conedsc9ra3_e_sbc_p1800 &"' ,  [conedsc9ra3_e_sbc_p2200] = '"&  conedsc9ra3_e_sbc_p2200 &"' ,  [conedsc9ra3_e_sbc_p2359] = '"&  conedsc9ra3_e_sbc_p2359 &"' ,  [conedsc9ra3m_e_sbc_p1800] = '"&  conedsc9ra3m_e_sbc_p1800 &"' ,  [conedsc9ra3m_e_sbc_p2200] = '"&  conedsc9ra3m_e_sbc_p2200 &"' ,  [conedsc9ra3m_e_sbc_p2359] = '"&  conedsc9ra3m_e_sbc_p2359 &"' ,  [conedsc9r1_e_tsc]  = '"&  conedsc9r1_e_tsc &"' ,  [conedsc9r1_d_tsc]  = '"&  conedsc9r1_d_tsc  &"' ,  [conedsc9r1m_e_tsc]  = '"&  conedsc9r1m_e_tsc  &"' ,  [conedsc9r1m_d_tsc]  = '"&  conedsc9r1m_d_tsc  &"' ,  [conedsc9ra1_e_tsc]  = '"&  conedsc9ra1_e_tsc  &"' ,  [conedsc9ra1_d_tsc]  = '"&  conedsc9ra1_d_tsc  &"' ,  [conedsc9ra1m_e_tsc]  = '"&  conedsc9ra1m_e_tsc  &"' ,  [conedsc9ra1m_d_tsc]  = '"&  conedsc9ra1m_d_tsc  &"' ,  [conedsc9r2_e_tsc]  = '"&  conedsc9r2_e_tsc  &"' ,  [conedsc9r2_d_tsc]  = '"&  conedsc9r2_d_tsc  &"' ,  [conedsc9r2m_e_tsc]  = '"&  conedsc9r2m_e_tsc  &"' ,  [conedsc9r2m_d_tsc]  = '"&  conedsc9r2m_d_tsc  &"' ,  [conedsc9ra2_e_tsc]  = '"&  conedsc9ra2_e_tsc  &"' ,  [conedsc9ra2_d_tsc]  = '"&  conedsc9ra2_d_tsc  &"' ,  [conedsc9ra3_e_tsc]  = '"&  conedsc9ra3_e_tsc  &"' ,  [conedsc9ra3_d_tsc]  = '"&  conedsc9ra3_d_tsc  &"' ,  [conedsc9ra3m_e_tsc]  = '"&  conedsc9ra3m_e_tsc  &"' ,  [conedsc9ra3m_d_tsc]  = '"&  conedsc9ra3m_d_tsc  &"' ,  [conedsc12ra2_e_tsc]  = '"&  conedsc12ra2_e_tsc &"' ,  [conedsc12ra2_d_tsc]  = '"&  conedsc12ra2_d_tsc &"' where rbrid= "& rbrid 
					'response.write insertsql
					'response.end

				cnn1.Execute insertSql
				
			else
				insertSql ="insert into ratebuilderrates (	sc9r1_e_er,  sc9r1_e_macadj,  sc9r1_d_mscadj,  sc9r1_d_dr_l05,  sc9r1_d_dr_l100,  sc9r1_d_dr_l999,  sc9r1_d_dlms,  conedsc9r1_s_bppc,  conedsc9r1_s_cmc,  conedsc9r1_e_er,  conedsc9r1_e_mfc,  conedsc9r1_d_dr_l05,  conedsc9r1_d_dr_l999,  conedsc9r1_e_cesss,  conedsc9r1_d_dlms,  sc9r2_e_er_p759,  sc9r2_e_er_p1800,  sc9r2_e_er_p2200,  sc9r2_e_er_p2359,  sc9r2_e_macadj_p759,  sc9r2_e_macadj_p1800,  sc9r2_e_macadj_p2200,  sc9r2_e_macadj_p2359,  sc9r2_d_mscadj_p2200,  sc9r2_d_dr_p1800,  sc9r2_d_dr_p2200,  sc9r2_d_dr_p2359,  sc9ra1_e_er,  sc9ra1_e_macadj,  sc9ra1_d_dr_l05,  sc9ra1_d_dr_l100,  sc9ra1_d_dr_l999,  sc9ra2_e_er_p759,  sc9ra2_e_er_p1800,  sc9ra2_e_er_p2200,  sc9ra2_e_er_p2359,  sc9ra2_e_macadj_p759,  sc9ra2_e_macadj_p1800,  sc9ra2_e_macadj_p2200,  sc9ra2_e_macadj_p2359,  sc9ra2_d_dr_p1800,  sc9ra2_d_dr_p2200,  sc9ra2_d_dr_p2359,  sc9ra3_e_er_p759,  sc9ra3_e_er_p1800,  sc9ra3_e_er_p2200,  sc9ra3_e_er_p2359,  sc9ra3_e_macadj_p759,  sc9ra3_e_macadj_p1800,  sc9ra3_e_macadj_p2200,  sc9ra3_e_macadj_p2359,  sc9ra3_d_dr_p1800,  sc9ra3_d_dr_p2200,  sc9ra3_d_dr_p2359,  sc12ra2_e_er_p759,  sc12ra2_e_er_p1800,  sc12ra2_e_er_p2200,  sc12ra2_e_er_p2359,  sc12ra2_e_macadj_p759,  sc12ra2_e_macadj_p1800,  sc12ra2_e_macadj_p2200,  sc12ra2_e_macadj_p2359,  sc12ra2_d_dr_p1800,  sc12ra2_d_dr_p2200,  sc12ra2_d_dr_p2359,  conedsc9r1m_s_bppc,  conedsc9r1m_s_cmc,  conedsc9r1m_e_er,  conedsc9r1m_e_mfc,  conedsc9r1m_d_mscadj,  conedsc9r1m_d_dr_l05,  conedsc9r1m_d_dr_l999,  conedsc9r1m_e_cesss,  conedsc9ra1_s_bppc,  conedsc9ra1_s_cmc,  conedsc9ra1_e_er,  conedsc9ra1_d_dr_l05,  conedsc9ra1_d_dr_l999,  conedsc9ra1_d_dlms,  conedsc9ra1m_s_bppc,  conedsc9ra1m_s_cmc,  conedsc9ra1m_e_er,  conedsc9ra1m_d_dr_l05,  conedsc9ra1m_d_dr_l999,  conedsc9r2_s_bppc,  conedsc9r2_s_cmc,  conedsc9r2_e_mfc,  conedsc9r2_e_cesss,  conedsc9r2m_s_bppc,  conedsc9r2m_s_cmc,  conedsc9r2m_e_mfc,  conedsc9ra2_s_bppc,  conedsc9ra2_s_cmc,  conedsc9ra3_s_bppc,  conedsc9ra3_s_cmc,  conedsc9ra3m_s_bppc,  conedsc9ra3m_s_cmc,  conedsc12ra2_s_bppc,  conedsc12ra2_s_cmc,  conedsc9r2_e_er_p1800,  conedsc9r2_e_er_p2200,  conedsc9r2_e_er_p2359,  conedsc9r2m_e_er_p1800,  conedsc9r2m_e_er_p2200,  conedsc9r2m_e_er_p2359,  conedsc9ra2_e_er_p1800,  conedsc9ra2_e_er_p2200,  conedsc9ra2_e_er_p2359,  conedsc9ra3_e_er_p1800,  conedsc9ra3_e_er_p2200,  conedsc9ra3_e_er_p2359,  conedsc9ra3m_e_er_p1800,  conedsc9ra3m_e_er_p2200,  conedsc9ra3m_e_er_p2359,  conedsc12ra2_e_er_p1800,  conedsc12ra2_e_er_p2200,  conedsc12ra2_e_er_p2359,  conedsc12ra2_d_dr_p1800,  conedsc12ra2_d_dr_p2200,  conedsc12ra2_d_dr_p2359,  conedsc9ra3m_d_dr_p1800,  conedsc9ra3m_d_dr_p2200,  conedsc9ra3m_d_dr_p2359,  conedsc9ra3_d_dr_p1800,  conedsc9ra3_d_dr_p2200,  conedsc9ra3_d_dr_p2359,  conedsc9ra2_d_dr_p1800,  conedsc9ra2_d_dr_p2200,  conedsc9ra2_d_dr_p2359,  conedsc9r2m_d_dr_p1800,  conedsc9r2m_d_dr_p2200,  conedsc9r2m_d_dr_p2359,  conedsc9r2m_e_cesss,  conedsc9r1m_d_dlms,  conedsc9ra1m_d_dlms,  sc9ra1_d_dlms,  conedsc12ra2_e_sbc_p1800,  conedsc12ra2_e_sbc_p2200,  conedsc12ra2_e_sbc_p2359,  conedsc9r2_e_sbc_p1800,  conedsc9r2_e_sbc_p2200,  conedsc9r2_e_sbc_p2359,  conedsc9r2_d_msccap,  conedsc9r1_d_msccap,  conedsc9r1_e_sbc,  conedsc9r1m_e_sbc,  conedsc9r1m_d_msccap,  conedsc9r2m_e_sbc_p1800,  conedsc9r2m_e_sbc_p2200,  conedsc9r2m_e_sbc_p2359,  conedsc9r2m_d_msccap_p2200,  conedsc9ra1_e_sbc,  conedsc9ra1m_e_sbc,  conedsc9ra2_e_sbc_p1800,  conedsc9ra2_e_sbc_p2200,  conedsc9ra2_e_sbc_p2359,  conedsc9ra3_e_sbc_p1800,  conedsc9ra3_e_sbc_p2200,  conedsc9ra3_e_sbc_p2359,  conedsc9ra3m_e_sbc_p1800,  conedsc9ra3m_e_sbc_p2200,  conedsc9ra3m_e_sbc_p2359, conedsc9r1_e_tsc, conedsc9r1_d_tsc, conedsc9r1m_e_tsc, conedsc9r1m_d_tsc, conedsc9ra1_e_tsc, conedsc9ra1_d_tsc, conedsc9ra1m_e_tsc, conedsc9ra1m_d_tsc, conedsc9r2_e_tsc, conedsc9r2_d_tsc, conedsc9r2m_e_tsc, conedsc9r2m_d_tsc, conedsc9ra2_e_tsc, conedsc9ra2_d_tsc, conedsc9ra3_e_tsc, conedsc9ra3_d_tsc, conedsc9ra3m_e_tsc, conedsc9ra3m_d_tsc, conedsc12ra2_e_tsc, conedsc12ra2_d_tsc ) output Inserted.rbrid values ("&_
				"'"&  sc9r1_e_er  &"','"&  sc9r1_e_macadj  &"','"&  sc9r1_d_mscadj  &"','"&  sc9r1_d_dr_l05  &"','"&  sc9r1_d_dr_l100  &"','"&  sc9r1_d_dr_l999  &"','"&  sc9r1_d_dlms  &"','"&  conedsc9r1_s_bppc  &"','"&  conedsc9r1_s_cmc  &"','"&  conedsc9r1_e_er  &"','"&  conedsc9r1_e_mfc  &"','"&  conedsc9r1_d_dr_l05  &"','"&  conedsc9r1_d_dr_l999  &"','"&  conedsc9r1_e_cesss  &"','"&  conedsc9r1_d_dlms  &"','"&  sc9r2_e_er_p759  &"','"&  sc9r2_e_er_p1800  &"','"&  sc9r2_e_er_p2200  &"','"&  sc9r2_e_er_p2359  &"','"&  sc9r2_e_macadj_p759  &"','"&  sc9r2_e_macadj_p1800  &"','"&  sc9r2_e_macadj_p2200  &"','"&  sc9r2_e_macadj_p2359  &"','"&  sc9r2_d_mscadj_p2200  &"','"&  sc9r2_d_dr_p1800  &"','"&  sc9r2_d_dr_p2200  &"','"&  sc9r2_d_dr_p2359  &"','"&  sc9ra1_e_er  &"','"&  sc9ra1_e_macadj  &"','"&  sc9ra1_d_dr_l05  &"','"&  sc9ra1_d_dr_l100  &"','"&  sc9ra1_d_dr_l999  &"','"&  sc9ra2_e_er_p759  &"','"&  sc9ra2_e_er_p1800  &"','"&  sc9ra2_e_er_p2200  &"','"&  sc9ra2_e_er_p2359  &"','"&  sc9ra2_e_macadj_p759  &"','"&  sc9ra2_e_macadj_p1800  &"','"&  sc9ra2_e_macadj_p2200  &"','"&  sc9ra2_e_macadj_p2359  &"','"&  sc9ra2_d_dr_p1800  &"','"&  sc9ra2_d_dr_p2200  &"','"&  sc9ra2_d_dr_p2359  &"','"&  sc9ra3_e_er_p759  &"','"&  sc9ra3_e_er_p1800  &"','"&  sc9ra3_e_er_p2200  &"','"&  sc9ra3_e_er_p2359  &"','"&  sc9ra3_e_macadj_p759  &"','"&  sc9ra3_e_macadj_p1800  &"','"&  sc9ra3_e_macadj_p2200  &"','"&  sc9ra3_e_macadj_p2359  &"','"&  sc9ra3_d_dr_p1800  &"','"&  sc9ra3_d_dr_p2200  &"','"&  sc9ra3_d_dr_p2359  &"','"&  sc12ra2_e_er_p759  &"','"&  sc12ra2_e_er_p1800  &"','"&  sc12ra2_e_er_p2200  &"','"&  sc12ra2_e_er_p2359  &"','"&  sc12ra2_e_macadj_p759  &"','"&  sc12ra2_e_macadj_p1800  &"','"&  sc12ra2_e_macadj_p2200  &"','"&  sc12ra2_e_macadj_p2359  &"','"&  sc12ra2_d_dr_p1800  &"','"&  sc12ra2_d_dr_p2200  &"','"&  sc12ra2_d_dr_p2359  &"','"&  conedsc9r1m_s_bppc  &"','"&  conedsc9r1m_s_cmc  &"','"&  conedsc9r1m_e_er  &"','"&  conedsc9r1m_e_mfc  &"','"&  conedsc9r1m_d_mscadj  &"','"&  conedsc9r1m_d_dr_l05  &"','"&  conedsc9r1m_d_dr_l999  &"','"&  conedsc9r1m_e_cesss  &"','"&  conedsc9ra1_s_bppc  &"','"&  conedsc9ra1_s_cmc  &"','"&  conedsc9ra1_e_er  &"','"&  conedsc9ra1_d_dr_l05  &"','"&  conedsc9ra1_d_dr_l999  &"','"&  conedsc9ra1_d_dlms  &"','"&  conedsc9ra1m_s_bppc  &"','"&  conedsc9ra1m_s_cmc  &"','"&  conedsc9ra1m_e_er  &"','"&  conedsc9ra1m_d_dr_l05  &"','"&  conedsc9ra1m_d_dr_l999  &"','"&  conedsc9r2_s_bppc  &"','"&  conedsc9r2_s_cmc  &"','"&  conedsc9r2_e_mfc  &"','"&  conedsc9r2_e_cesss  &"','"&  conedsc9r2m_s_bppc  &"','"&  conedsc9r2m_s_cmc  &"','"&  conedsc9r2m_e_mfc  &"','"&  conedsc9ra2_s_bppc  &"','"&  conedsc9ra2_s_cmc  &"','"&  conedsc9ra3_s_bppc  &"','"&  conedsc9ra3_s_cmc  &"','"&  conedsc9ra3m_s_bppc  &"','"&  conedsc9ra3m_s_cmc  &"','"&  conedsc12ra2_s_bppc  &"','"&  conedsc12ra2_s_cmc  &"','"&  conedsc9r2_e_er_p1800  &"','"&  conedsc9r2_e_er_p2200  &"','"&  conedsc9r2_e_er_p2359  &"','"&  conedsc9r2m_e_er_p1800  &"','"&  conedsc9r2m_e_er_p2200  &"','"&  conedsc9r2m_e_er_p2359  &"','"&  conedsc9ra2_e_er_p1800  &"','"&  conedsc9ra2_e_er_p2200  &"','"&  conedsc9ra2_e_er_p2359  &"','"&  conedsc9ra3_e_er_p1800  &"','"&  conedsc9ra3_e_er_p2200  &"','"&  conedsc9ra3_e_er_p2359  &"','"&  conedsc9ra3m_e_er_p1800  &"','"&  conedsc9ra3m_e_er_p2200  &"','"&  conedsc9ra3m_e_er_p2359  &"','"&  conedsc12ra2_e_er_p1800  &"','"&  conedsc12ra2_e_er_p2200  &"','"&  conedsc12ra2_e_er_p2359  &"','"&  conedsc12ra2_d_dr_p1800  &"','"&  conedsc12ra2_d_dr_p2200  &"','"&  conedsc12ra2_d_dr_p2359  &"','"&  conedsc9ra3m_d_dr_p1800  &"','"&  conedsc9ra3m_d_dr_p2200  &"','"&  conedsc9ra3m_d_dr_p2359  &"','"&  conedsc9ra3_d_dr_p1800  &"','"&  conedsc9ra3_d_dr_p2200  &"','"&  conedsc9ra3_d_dr_p2359  &"','"&  conedsc9ra2_d_dr_p1800  &"','"&  conedsc9ra2_d_dr_p2200  &"','"&  conedsc9ra2_d_dr_p2359  &"','"&  conedsc9r2m_d_dr_p1800  &"','"&  conedsc9r2m_d_dr_p2200  &"','"&  conedsc9r2m_d_dr_p2359  &"','"&  conedsc9r2m_e_cesss  &"','"&  conedsc9r1m_d_dlms  &"','"&  conedsc9ra1m_d_dlms  &"','"&  sc9ra1_d_dlms  &"','"&  conedsc12ra2_e_sbc_p1800  &"','"&  conedsc12ra2_e_sbc_p2200  &"','"&  conedsc12ra2_e_sbc_p2359  &"','"&  conedsc9r2_e_sbc_p1800  &"','"&  conedsc9r2_e_sbc_p2200  &"','"&  conedsc9r2_e_sbc_p2359  &"','"&  conedsc9r2_d_msccap  &"','"&  conedsc9r1_d_msccap  &"','"&  conedsc9r1_e_sbc  &"','"&  conedsc9r1m_e_sbc  &"','"&  conedsc9r1m_d_msccap  &"','"&  conedsc9r2m_e_sbc_p1800  &"','"&  conedsc9r2m_e_sbc_p2200  &"','"&  conedsc9r2m_e_sbc_p2359  &"','"&  conedsc9r2m_d_msccap_p2200  &"','"&  conedsc9ra1_e_sbc  &"','"&  conedsc9ra1m_e_sbc  &"','"&  conedsc9ra2_e_sbc_p1800  &"','"&  conedsc9ra2_e_sbc_p2200  &"','"&  conedsc9ra2_e_sbc_p2359  &"','"&  conedsc9ra3_e_sbc_p1800  &"','"&  conedsc9ra3_e_sbc_p2200  &"','"&  conedsc9ra3_e_sbc_p2359  &"','"&  conedsc9ra3m_e_sbc_p1800  &"','"&  conedsc9ra3m_e_sbc_p2200  &"','"&  conedsc9ra3m_e_sbc_p2359 &"','"&  conedsc9r1_e_tsc &"','"&  conedsc9r1_d_tsc &"','"&  conedsc9r1m_e_tsc &"','"&  conedsc9r1m_d_tsc &"','"&  conedsc9ra1_e_tsc &"','"&  conedsc9ra1_d_tsc &"','"&  conedsc9ra1m_e_tsc &"','"&  conedsc9ra1m_d_tsc &"','"&  conedsc9r2_e_tsc &"','"&  conedsc9r2_d_tsc &"','"&  conedsc9r2m_e_tsc &"','"&  conedsc9r2m_d_tsc &"','"&  conedsc9ra2_e_tsc &"','"&  conedsc9ra2_d_tsc &"','"&  conedsc9ra3_e_tsc &"','"&  conedsc9ra3_d_tsc &"','"&  conedsc9ra3m_e_tsc &"','"&  conedsc9ra3m_d_tsc &"','"&  conedsc12ra2_e_tsc &"','"&  conedsc12ra2_d_tsc &"'" &" )"
				
				'response.write(insertSql) &"</br>"
				'response.end
				cnn1.Execute insertSql
				rst1.Open"Select rbrid from ratebuilderrates where rbrid = SCOPE_IDENTITY()", cnn1
				if not rst1.eof then rbrid = rst1("rbrid")
				rst1.Close		
				
				insertSql ="update ratebuilder set rbrid ="&rbrid&" where rbid ="&rbid
				
				'response.write(insertSql) &"</br>"
				cnn1.Execute insertSql
				'response.write"rbcid:"&rbcid &" | rbrid:"&rbrid &" | rbid:"&rbid&"</br>"
				'rst1.Close
				
				'-------------------
			end if
			
			dim table, cname, cid, sname, rname, rtype, rline, rrange, r, rid, nSql, clen, sSql, rperiod, m, vari, vlen
			dim ir_rid, ir_lc, ir_lcid, ir_type, ir_start, ir_end, ir_mstart, ir_mend, ir_rfrom, ir_rto, ir_season, ir_peak, ir_peakid, ir_level, slvl
			dim curr_id 
			dim last_id 
			dim delrateid()
			table ="ratebuilderrates"
			
			ssql ="select rateperiod from ratebuilder where rbid =" & rbid
			rst1.open ssql, cnn1
			rperiod = rst1("rateperiod")
			rst1.close
			ir_start = dateadd("d",0,rperiod)
			ir_end = dateadd("d", -1, dateadd("m",1,rperiod))
			
			m = toNumb(split(rperiod,"-")(1))
			ir_mstart = m
			ir_mend = m
			
			nSql ="select c.name from sys.columns c inner join sys.tables t on c.object_id = t.object_id where t.name = '"& table &"' order by c.name asc"
			rst1.open nSql, cnn1
			sSql ="select * from ratebuilderrates where rbrid =" & rbrid
			rst2.open ssql, cnn1
			slvl = 0
			do until rst1.eof
				cname = rst1("name")
				if right(cname,2) <>"id" then
					rid =-1
					ir_season ="Winter"
					sname = Split(cname,"_")
					clen = ubound(sname)
					
					if clen > 0 then
						rname=""
						rtype=""
						vari=""
						rline=""
						
						rname = sname(0)
						if clen > 0 then rtype = sname(1)
						if clen > 1 then rline = sname(2)
						if clen > 2 then 
							vari = sname(3)
							vlen = len(vari)
							rrange = right(vari, vlen-1)
							vari = left(vari,1)
						end if
						
						r = rst2(cname)
						
						response.write cname &":" & rname &" |" & rtype &" |" & rline &" |" & ucase(vari)&"."&rrange &" :" & r &"<" & rperiod &" >" &"</br>"
						
						sSql ="select id from ratetypes where replace(replace(replace(type,' ',''),'rider',''),'-','') like '%" & rname &"%'"
						rst3.open ssql, cnn1
						ir_rid = rst3("id")
						redim preserve delrateid(ir_rid)
						rst3.close
						
						sSql ="select id,description from ratedescription where description like '%" & linecharge(rline) &"%'"
						rst3.open ssql, cnn1
						ir_lcid = rst3("id")
						ir_lc = rst3("description")
						rst3.close
						
						ir_type = itemtype(rtype)
						if (ir_mstart >= 6 and ir_mstart <= 9) then ir_season ="Summer"
						response.write"+   "& ir_rid &"    "&ir_type &"   "&ir_lc &"</br>"
						response.write"+   "& ir_season&"."&ir_mstart &"/" & ir_mend &"   :  " & ir_start &"-" & ir_end 
						response.write"</br>"
						
						if vari ="p" then 
							sSql ="SELECT rp.id as rpid, * FROM ratepeak rp INNER JOIN rateSeasons rs ON rs.id=rp.seasonid WHERE rs.regionid=5 and rs.season like '%"&ir_season&"%' and rp.etime like '%"& rrange &"%'"
							response.write ssql &"</br>"
							rst3.open ssql, cnn1
							
						
							ir_peak = rst3("label") &":" & weekdayname(cint(rst3("sweekday"))) &"-" & weekdayname(cint(rst3("eweekday"))) & rst3("stime") &"-" & rst3("etime") &" (" & rst3("season") &")"
							ir_peakid = rst3("rpid")
							rst3.close
						elseif vari="pri" or vari="gt" or vari="sec" then
							sSql ="SELECT rp.id as rpid, * FROM ratepeak rp INNER JOIN rateSeasons rs ON rs.id=rp.seasonid WHERE rs.regionid=5 and rs.season like '%"&ir_season&"%' and rp.etime like '%"& rrange &"%' and peakname=99" 
							response.write ssql &"</br>"
							rst3.open ssql, cnn1
							
						
							ir_peak = rst3("label") &":" & weekdayname(cint(rst3("sweekday"))) &"-" & weekdayname(cint(rst3("eweekday"))) & rst3("stime") &"-" & rst3("etime") &" (" & rst3("season") &")"
							ir_peakid = rst3("rpid")
							rst3.close						
						else
							ir_peak ="No Rate Peak"
							ir_peakid = 0
							ir_season = monthname(month(ir_start))
						end if
						if vari ="l" then
							if rrange = 999 then rrange = 999999999
							ir_level = slvl &"-" & rrange
						else
							slvl = 0
							rrange = 999999999
							ir_level = slvl &"-" & rrange
						end if
						
						if ir_rid <> last_id and last_id <> 0 then
							%></table></br></br><%
						end if
						if ir_rid <> last_id then
							%>
							<table width="100%"><tr width="100%"><td colspan="5"><h3><%= ucase(rname) %></h3></td></tr>
							<%
						end if
						
						last_id = ir_rid
						
						cid = cname&"_id"
						
						ssql ="select isnull("& cid &",-1) as "& cid &" from ratebuilderrates where rbrid ="& rbrid 
						'&" and " & cid &" is not null"
						response.write ssql &"</br>"
						'response.end
						rst3.open ssql, cnn1
						if not rst3.eof then rid = rst3(cid) end if
						rst3.close
						
						'if rid <> -1 then
							'sql="select id from rate where id ="&rid
							'rst3.open sql,cnn1
							'if rst3.eof then rid =-1 end if
							'rst3.close
							response.write "rbrid="& rbrid &" | "& cid &" | rid="& rid &" | ()"& delrateid(ir_rid) &"</br>"
							if delrateid(ir_rid) <> true then
								response.write "<b>DETELING ROWS for " & ir_rid & "</b> </br>"
								insertsql="delete from rate where [type]="&  ir_rid  &" and monthstart="&  ir_mstart  &" and monthend="&  ir_mend  &" and startdate='"&  ir_start  &"' and enddate='"&  ir_end  &"'"
								cnn1.execute insertsql
								delrateid(ir_rid) = true
							end if
							'rid = -1
						'end if
						
						'if rid =-1 then
							insertsql ="SET NOCOUNT ON;"&_
							"insert into rate (rate, type, peak, utility, ratefrom, rateto, itemtype, linecharge, monthstart, monthend, startdate, enddate)"&_
							" output Inserted.id values ("&_ 
							r &"," & ir_rid &"," & ir_peakid &"," & utility &","& slvl &"," & rrange &",'"& ir_type &"',"& ir_lcid &"," & ir_mstart &"," & ir_mend &",'"& ir_start &"','"& ir_end &"');"
							response.write insertsql &"</br>"
							'response.end
							
							set rid = cnn1.Execute(insertSql)
							rid = rid(0).value
							'rst3.Open"Select id from rate where id = SCOPE_IDENTITY()", cnn1
							'if not rst3.eof then rid = rst3("id")
							'rst3.Close		
				
							insertSql ="update ratebuilderrates set "& cid &" ="&rid&" where rbrid ="&rbrid
							response.write(insertSql) &"</br></br>"
							cnn1.Execute insertSql
						'else
						'	insertSql ="update rate set rate ="& r &"  where id ="& rid
							response.write insertsql &"</br></br>"
							'response.end
						'	cnn1.Execute insertSql
						'end if
						
						if vari ="l" then
							slvl = rrange+1
							if slvl > 999 then slvl = 0
						end if
						
						%>
						<tr width="100%">
							<td width="5%">Electricity</td>
							<td width="5%"><%= r %> </td>
							<td width="25%"><%= ir_peak %></td>
							<td width="5%"><%= ir_level  %></td>
							<td width="5%"><%= ir_type %></td>
							<td width="15%"><%= ir_lc %></td>
							<td width="10%"><%= ir_season %></td>
							<td width="10%"><%= ir_start %> - <%= ir_end %></td>
							<td width="5%"><%= rid %></td>
						</tr>
						<%
					
					end if
				end if
			rst1.movenext
			loop
			rst1.close
			rst2.close
			
			
			'-------------------
			'response.end
			Response.redirect"../rateTypeView.asp?rid=5"
		%>

		