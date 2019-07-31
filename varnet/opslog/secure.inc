<%
If Session("loginemail") = "" Then
  Session("goback") = Request.ServerVariables("URL")
  Response.Redirect "login.asp"
End If
%>
