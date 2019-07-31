<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
		else
			if  Session("um") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				Response.Redirect "../main.asp"
			end if	
		end if	
			
' Write a browser-side script 
Response.Write "<script>" & vbCrLf
  Response.Write "window.open(" & Chr(34) & "http://10.0.7.23/um_gate.jsp?username=" & Trim(Session("login")) & "&userlevel=" & Session("um") & Chr(34) & "," & Chr(34) & "GenergyOne" & Chr(34) & "," & Chr(34) & "resizeable=yes,statusbar=yes,toolbar=yes,height=768,width=1024" & Chr(34) &")"  & vbCrLf
   
Response.Write "</script>" & vbCrLf

%>