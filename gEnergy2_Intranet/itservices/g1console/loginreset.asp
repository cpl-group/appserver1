<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<title>gEnergyOne Account Reset Module</title>
<div align="center">
  <p><font size="2" face="Arial, Helvetica, sans-serif"><strong>USER ACCOUNT RESET 
    MODULE</strong></font></p>
  <p><font size="2" face="Arial, Helvetica, sans-serif">Status: RESET IN PROGRESS...</font></p>
</div>
<%
dim appmode ,username

appmode = trim(request("mode"))
username = trim(request("username"))

if appmode = "" then 
	appmode = "confirm"
end if 

select case appmode


	case "reset"
	
			Set cnn1 	= Server.CreateObject("ADODB.Connection")
			Set rs 		= Server.CreateObject("ADODB.recordset")
			
			cnn1.Open getConnect(0,0,"dbCore")
			
			sql = "delete from LoginTracking where username = '"&username&"'"
			
			cnn1.execute sql
			set cnn1=nothing
			%>
			<script>
			function ResetAccount(username){
				alert("Account Has Been Reset, user can now log-in at www.genergyonline.com")
				window.close()			
			}
			ResetAccount('<%=username%>')
			</script>
			<%
	case "confirm"
			%>
			<script>
			function confirmReset(username){
				
			var yesno=confirm("Are you sure you want to reset user account ["+username+"]? \nNOTE: USER MUST EXIT ALL INTERNET EXPLORER WINDOWS BEFORE RESETTING")
        	if(yesno) {
         			document.location = "./loginreset.asp?mode=reset&username="+username
				}
			else
				window.close()
			}
			confirmReset('<%=username%>')
			</script>
			<%
end select
%>
