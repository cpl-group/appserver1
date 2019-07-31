<%
If Session("loginemail") = "" Then
  Session("goback") = Request.ServerVariables("URL")
  Response.redirect "login.asp"

End If
%>
