<%@Language="VBScript"%>
<%Option Explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->

<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>
<body bgcolor="#FFFFFF" text="#000000">
<script> 
function loadentry(mkid){
	tempurl="mktview.asp?mkid=" + mkid
	document.location = tempurl

}
</script>
<%
Dim cnn1,rst1,sqlstr
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"

sqlstr = "SELECT m.status, m.recordingdate,m.situation,m.id,salesmanagers.manager as salesmanager ,[first_name]+ ' '+ [last_name] as contact,title,phone, email  FROM MKTLog m join contacts on contacts.id=m.contact full join salesmanagers on salesmanagers.id = m.salesmanager WHERE m.id NOT IN (SELECT mkid FROM rfplog WHERE mkid IS NOT NULL AND mkid <> 0) order by m.recordingdate desc"

rst1.Open sqlstr, cnn1, adOpenStatic

if rst1.EOF then

%>
<table width="100%" border="0">
  <tr>
    <td>
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="3"><b>NO 
        OPEN INTERACTIONS FOUND</b></font></div>
    </td>
  </tr>
</table>
<%
else
Dim numRecords
numRecords = rst1.RecordCount

%>
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="30"> <font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF" size="4">Genergy 
      Active Interactions - Currently there are <%=numRecords%> Active Interactions</font></b></font></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF"> 
      <div style="overflow:visible;height:20"> 
        <table width="100%" border="1" cellspacing="0" cellpadding="0" bordercolor="#000000">
          <tr>
            <td>
              <div align="center">
                <div style="overflow:visible;height:20">
                  <table width="900" border="0" cellspacing="0" cellpadding="0" align="center">
                    <tr> 
                      <td  bgcolor="#CCCCCC" height="28" width="10%"><font face="Arial, Helvetica, sans-serif"> 
                        Manager </font></td>
                      <td bgcolor="#CCCCCC" height="28" width="11%"><font face="Arial, Helvetica, sans-serif">Start 
                        Date </font></td>
                      <td  bgcolor="#CCCCCC" height="28" width="16%"><font face="Arial, Helvetica, sans-serif">Situation 
                        </font></td>
                      <td  bgcolor="#CCCCCC" height="28" width="9%"><font face="Arial, Helvetica, sans-serif">Status</font></td>
                      <td bgcolor="#CCCCCC" height="28" width="17%"> 
                        <div align="center"><font face="Arial, Helvetica, sans-serif"> 
                          Contact</font></div>
                      </td>
                      <td  bgcolor="#CCCCCC" height="28" width="19%"><font face="Arial, Helvetica, sans-serif">Phone 
                        # </font></td>
                      <td  bgcolor="#CCCCCC" height="28" width="18%"><font face="Arial, Helvetica, sans-serif">Email</font></td>
                    </tr>
                  </table>
                </div>
                <div style="overflow:auto;height:220"> 
                  <table width="900" border="0" cellpadding="0" cellspacing="3" align="center">
                    <% While not rst1.EOF %>
                    <tr align="left" valign="top" bgcolor="#FFFFCC" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = '#FFFFCC'" onclick="javascript:loadentry('<%=trim(rst1("id"))%>')"> 
                      <td  height="37" width="10%"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("salesmanager")%></font></td>
                      <td " height="37" width="8%"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("Recordingdate")%></font></td>
                      <td  height="37" width="10%"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("situation")%></font></td>
                      <td  height="37" width="16%"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("Status")%></font></td>
                      <td  height="37" width="10%"> 
                        <div align="left"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("contact")%></font></div>
                      </td>
                      <td  height="37" width="11%"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("phone")%></font></td>
                      <td width="12%" height="37"><font face="Arial, Helvetica, sans-serif" size="2"><%=rst1("email")%></font></td>
                    </tr>
                    <% 
		rst1.movenext
		Wend
		%>
                  </table>
                </div>
    </div>
            </td>
          </tr>
        </table>
      </div>
      </td>
  </tr>
  <tr>
    <td bgcolor="#3399CC"><font face="Arial, Helvetica, sans-serif" size="2"><i><font color="#FFFFFF">Note: 
      THIS LIST SHOWS ACTIVE INTERACTIONS;INTERACTIONS THAT HAVE NO RFP PRESENT 
      IN THE RFP LOG</font></i></font></td>
  </tr>
</table>
<p>&nbsp;</p>
<%
end if
%>
</body>
</html>