<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<HTML>
<BODY BGCOLOR="white">

<H1><font size="3" face="Arial, Helvetica, sans-serif">File Upload Complete.</font></H1>
<HR>

<%
'  Variables
'  *********
   Dim mySmartUpload
   Dim intCount, downloadpath
   
'  Object creation
'  ***************
   Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")

'  Upload
'  ******
   mySmartUpload.Upload

'  Save the files with their original names in a virtual path of the web server
'  ****************************************************************************
	
   intCount = mySmartUpload.Save("/jobfiles/" & session("jobdir") & "/" & session("jobid"))
   ' sample with a physical path 
   ' intCount = mySmartUpload.Save("c:\temp\")

'  Display the number of files uploaded
'  ************************************
   Response.Write(intCount & " file(s) uploaded.")
   session("jobdir") = null
   session("jobid") = null
%>
<script>
window.close()
</script>
</BODY>
</HTML>