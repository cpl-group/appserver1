<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, tid, lid, mid, action
action = secureRequest("action")
pid = secureRequest("pid")
bldg = secureRequest("bldg")
tid = secureRequest("tid")
lid = secureRequest("lid")
mid = secureRequest("meterid")
'response.write "[" & action & "]"
'response.end
%>
<html>
<head>
<title>Utility Manager</title>
</head>
<frameset rows="130,*" frameborder="0" framespacing="0">
<%
if trim(action)<>"" then
	%>
	<frame name="meterfrm" src="meterview.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&meterid=<%=mid%>" scrolling="auto" marginwidth="0" marginheight="0" border=1 style="border-right:1px solid #cccccc;">
	<%
	if trim(mid)<>"" or trim(action)="meteradd" then 
		%>
		<frame name="editfrm" src="meteredit.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&meterid=<%=mid%>" scrolling="auto" marginwidth="0" marginheight="0" border=1 style="border-right:1px solid #cccccc;">
      <%
	else
		%>
		<!--[[%if trim(action)="showmeters" then %]]-->
		<frame name="editfrm" src="meternull.htm" scrolling="auto" marginwidth="0" marginheight="0">
		<!--
        [[% else %]]
        [[frame name="editfrm" src="meteredit.asp?pid=[[%=pid%]]&bldg=[[%=bldg%]]&tid=[[%=tid%]]&lid=[[%=lid%]]" scrolling="auto" marginwidth="0" marginheight="0"]]
        [[% end if %]]
		-->
      <%
	end if
	%>
   <%
else 
	%>
	<frame name="meterfrm" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" border=1 style="border-right:1px solid #eeeeee;">
	<frame name="editfrm" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0">
	<%
end if
%>
</frameset><noframes></noframes>
</html>