<title>User ID Auto Re-order</title><%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<% 
response.write "AUTO-ORDER NOW IN PROGRESS..."
userid = session("editemail")

if trim(userid) <> "" then 
		Set cnn = Server.CreateObject("ADODB.Connection")
		Set rs 		= Server.CreateObject("ADODB.recordset")
		cnn.open getConnect(0,0,"dbCore")
		
		strsql = "select bldgid, bldgname from clientsetup where userid = '"&userid&"' group by bldgid, bldgname order by bldgid"
		rs.open strsql, cnn
		
		if not rs.eof then 
			x = 0
			while not rs.eof 
				bldgid = rs("bldgid")
				strsql = "update clientsetup set bldgorder = " & x & " where bldgid = '" &bldgid& "' and userid= '" & userid &"'"
				cnn.execute strsql
				rs.movenext
				x = x + 1
			wend			
		end if
		rs.close
		
end if 
%>
		<script>
		opener.window.location = opener.window.location
		window.close()
		</script>	
