<%Option Explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE FILE="AccountManage.asp"-->
<%

	usrmg.AddUser "testname", "testword", "testtest", "GenergyOne Clients","clientFinancials|clientOperations"
	response.write "I made it here!!!!"
%>