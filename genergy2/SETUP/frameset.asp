<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if not(allowGroups("Genergy Users")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim pid, bldg, tid, lid, mid, action
action = secureRequest("action")
pid = secureRequest("pid")
bldg = secureRequest("bldg")
tid = secureRequest("tid")
lid = secureRequest("lid")
mid = secureRequest("meterid")

%>
<html>
<head>
<title>Utility Manager</title>
</head>
<frameset rows="100,*" frameborder="0" framespacing="0">
  <% if trim(action)<>"" then %>
  <frame name="toolbarfrm" src="toolbar.asp?pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>" scrolling="auto" marginwidth="0" marginheight="0">
  <frame name="contentfrm" src="contentfrm.asp?action=<%=action%>&pid=<%=pid%>&bldg=<%=bldg%>&tid=<%=tid%>&lid=<%=lid%>&meterid=<%=mid%>" scrolling="auto" marginwidth="0" marginheight="0" border=0>
  <% else %>
  <frame name="toolbarfrm" src="toolbar.asp" scrolling="auto" marginwidth="0" marginheight="0">
  <frame name="contentfrm" src="contentfrm.asp" scrolling="auto" marginwidth="0" marginheight="0" border=0>
  <% end if %>
</frameset>
</html>



