<%
if trim(session("login"))="" then
'	response.write "Login expired. Please <a href=""/um/index.asp"" target=""login"">login</a> again"
	response.write "Login expired. Please login again."
	response.end
end if
%>
