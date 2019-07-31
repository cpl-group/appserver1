<%
If Session("loginemail") = "" Then
  Session("goback") = Request.ServerVariables("URL")
  response.write "login.asp"

End If
%>
