<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<HTML><title>Upload to Job Folder <%=Request("job")%></title>
<head><link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css"></head>
<BODY BGCOLOR="white">
<%
Session("jobdir") = request("dir")
Session("jobid") = request("job")
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#6699cc"><span class="standardheader">Upload to Job Folder 
      <%=Request("job")%>:</span></td>
  </tr>
  <tr> 
    <td bgcolor="#dddddd"><FORM METHOD="POST" ACTION="upload.asp" ENCTYPE="multipart/form-data">
   <INPUT TYPE="FILE" NAME="FILE1" SIZE="50"><BR>
   <INPUT TYPE="FILE" NAME="FILE2" SIZE="50"><BR>
   <INPUT TYPE="FILE" NAME="FILE3" SIZE="50"><BR>
   <INPUT TYPE="FILE" NAME="FILE4" SIZE="50"><BR>
   <INPUT TYPE="SUBMIT" VALUE="Upload">
</FORM>
</td>
  </tr>
</table>
<H1>&nbsp;</H1>

</BODY>
</HTML>
