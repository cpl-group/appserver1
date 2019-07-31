<%
if trim(session("clientid"))="" then
	response.write "You are not logged in. To login please <a href=""javascript:window.open('index.asp', 'login','height=300,width=400,top=282,left=376,status=yes,scrollbars=no,resizable=no');top.close();"">click here</a>"
	response.end
end if

%>