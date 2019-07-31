<%@Language="VBScript"%>
<%

		if isempty(Session("loginemail")) then
			Response.Redirect "http://www.genergyonline.com"	
		else
			nocache=rnd*1000000
			Response.Redirect "http://"&request.servervariables("server_name")&"/g1_clients/g1nav.asp?nfc="&clng(nocache)
		end if		
		
%>
