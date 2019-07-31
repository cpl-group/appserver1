<%
If Session("loginemail") = "" Then
  Session("goback") = Request.ServerVariables("URL")
  response.write "<script> window.close() </script>"

End If
%>
