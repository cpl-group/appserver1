
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
			Response.Write("<font color='red'> Incorrect Password!  Please enter the correct password </font>")

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

<style>
h1 { font-family: Arial, Helvetica, sans-serif; font-size: 16px; margin: 0px; padding: 0px; }
p { font: Arial, Helvetica, sans-serif; line-height: 1.5px; }
</style>
</head>

<body>
    <table style="font-family:Arial, Helvetica, sans-serif; font-size: 12px" cellspacing="0" cellpadding="4" border="0" align="center">
        <tr>
        	<td height="20"></td>
        </tr>
        <tr>
            <td style="text-align: center"><h1 style="color: #ff0000;">!&nbsp;WARNING&nbsp;!</h1><br />
            	<p>Once you are disconnected, the load server by this meter can not be re-enabled or reconnected.</p></td>
        </tr>
        <tr>	
        	<td height="10"></td>
        </tr>
        <tr>
        	<td style="text-align: center">Enter your password below to disconnect the meter:</td>
        </tr>
        <tr>
        	<td height="10"></td>
        <tr>
        	<td style="text-align: center"><form action="remoteDisconnect.asp" name="remoteDisconnect" method="post" onSubmit="return confirmation()">
	<input type="password" name="Password" value="<%=Request.Form("Password")%>">
	<input type="submit" name="SubmitButton" value="Submit">
 </form></td>
        </tr>
    </table>

 
</body>
</html>
