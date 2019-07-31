<%@Language=VBScript%>
<!--#include file="adovbs.inc"-->
<%

Dim cnn1
Dim rs
Dim strsql
Dim bldg
Dim meterid, lmpstart, lmpend,lmpdate

m=Request.QueryString("m")
d=Request.QueryString("d")
b=Request.QueryString("b")
s=Request.QueryString("s")
e=Request.QueryString("e")



Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"


strsql = "select date from pulse_" & b & "  where meterid='" & m & "' group by date order by date desc"


rs.Open strsql, cnn1, adOpenStatic

if rs.EOF then %>
<title>Available Profile Dates</title><body bgcolor="#FFFFFF">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF">
  <tr>
    <td>
      <div align="center"> <font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif">NO DATES CURRENTLY AVAILABLE</font></b></font></div>
    </td>
  </tr>
</table>
<%
else
%>
<script>
function loadentry(b,m,s,e,d,pd,nd){
	var lmp = opener.document.forms[0].lmp.value
	var luid = opener.document.forms[0].luid.value
	var temp="https://appserver1.genergy.com/eri_th/lmp/lmpload.asp?m="+m+"&d="+d+"&b="+b+"&s="+s+"&e="+e+"&luid="+luid+"&lmp="+lmp
	opener.document.forms[0].d.value=d
	opener.document.forms[0].nd.value = nd
	opener.document.forms[0].pd.value = pd
	opener.document.frames.lmp.location=temp;
	window.close()
}
</script>
<title>Available Profile Dates</title><body bgcolor="#FFFFFF">
<body bgcolor="#FFFFFF"> 
<table width="100%" border="1" cellspacing="0" cellpadding="0">
  <tr>
    <td>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td> 
            <div align="center"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="11%" bgcolor="#0099FF"> 
                    <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">DATES 
                      AVAILABLE FOR PROFILE</font></b></font></div>
                  </td>
                </tr>
                <%while not rs.EOF 
		%>
                <form name="form1" method="post" action="">
              <input type="hidden" name="b" value="<%=b%>">
              <input type="hidden" name="m" value="<%=m%>">
              <input type="hidden" name="s" value="<%=s%>">
              <input type="hidden" name="e" value="<%=e%>">
              <input type="hidden" name="d" value="<%=rs("date")%>">
			  <input type="hidden" name="pd" value=<%=DateAdd("d",-1,rs("date"))%>>
			  <input type="hidden" name="nd" value=<%=DateAdd("d",1,rs("date"))%>>
			  <input type="hidden" name="td" value=<%=Date()%>>
			  

                  <tr valign="top" onMouseOver="this.style.backgroundColor = 'lightgreen'" onMouseOut="this.style.backgroundColor = 'white'" onClick="loadentry(b.value, m.value, s.value, e.value, d.value,pd.value,nd.value)"> 
                    <td width="11%" height="17"> 
                      <div align="center"><font size="2"> <font face="Arial, Helvetica, sans-serif"><%=rs("date")%></font></font> 
                      </div>
                    </td>
                  </tr>
                </form>
                <%
		rs.movenext
		Wend
		%>
              </table>
            </div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<%end if
set cnn1=nothing
rs.close
%>
