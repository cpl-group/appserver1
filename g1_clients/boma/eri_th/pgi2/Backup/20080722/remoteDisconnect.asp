
<%
  Dim meterId 
  'Dim bId
  meterId = 11158
  'bId = 113LAW

  If Request.Form("SubmitButton") = "Submit" Then

	  If (Request.Form("Password") = "shutdown") Then
	  
		'Dim conn, rs, sql, pw
		'pw = Request.Form("Password")
		'conn = Server.CreateObject(ADODB.Connection)
		'rs = Server.CreateObject(ADODB.Recordset)
		'sql = "select [password] FROM a_table WHERE meterId = " & meterId & " buildingId = " & buildingId

		'Check to see if password is correct

			Response.Write("<font color='blue'> Load has been disconnected </font>")
			Response.Redirect "remoteDisconnectPopup.asp"
	   Else 
			'Response.Write("<font color='red'> Incorrect Password!  Please enter the correct password </font>")
			MsgBox("Hello")
	   End If

	End if
%>
<html>
<head>

<SCRIPT LANGUAGE="JavaScript">
<!--
	function confirmation() 
	{
		var answer = confirm(" *Warning: Once disconnected, the load served by this meter can not be re-enabled or reconnected.\n\n  Are you sure want to disconnect this load?");
		if(answer) {	
			return true;
		}
		else {
			return false;
		}
	}
-->
</SCRIPT>

</head>

<body>
  <h3> *Warning: Once disconnected, the load served by this meter can not be re-enabled or reconnected </h3>
  <br />
  <br />
  Enter your password below to disconnect the meter:
 <form action="remoteDisconnect.asp" name="remoteDisconnect" method="post" onSubmit="return confirmation()">
	<input type="password" name="Password" value="<%=Request.Form("Password")%>">
	<input type="submit" name="SubmitButton" value="Submit">
 </form>
</body>
</html>
