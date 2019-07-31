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
luid= Request.QueryString("luid")
%>

<script>
function loadentry(b,m,d,s,e,inter, luid){
	var lmp = opener.document.forms[0].lmp.value

	var temp="https://appserver1.genergy.com/eri_th/lmp/lmpload.asp?m="+m+"&d="+d+"&b="+b+"&s="+s+"&e="+e+"&i="+inter+"&luid="+luid+"&lmp="+lmp
	opener.document.frames.lmp.location=temp;
	window.close()
}
</script>
<title>Zoom to Time/Interval</title><body bgcolor="#FFFFFF">
<table width="100%" border="1" cellspacing="0" cellpadding="0">
  <tr>
    <td>

      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td> 
            <div align="center"> 
			<form name="form1" method="post" action="">
              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr> 
                  <td width="11%" bgcolor="#0099FF"> 
                    <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">ZOOM</font></b></font></div>
                  </td>
                </tr>
                <tr> 
                  <td width="11%" height="17"> 
                    <div align="center"><font size="2" face="Arial, Helvetica, sans-serif"> 
                      <input type="hidden" name="b" value="<%=b%>">
                      <input type="hidden" name="m" value="<%=m%>">
                      <input type="hidden" name="d" value="<%=d%>">
                      <input type="hidden" name="luid" value="<%=luid%>">
                      FROM: 
                      <select name="s">
                        <%for i=1 to 24%>
                        <option value="<%=i*100%>"><%=i*100%></option>
                        <%next%>
                      </select>
                      TO:</font> <font size="2" face="Arial, Helvetica, sans-serif"> 
                      <select name="e">
                        <%for i=1 to 24%>
                        <option value="<%=2500-(i*100)%>"><%=2500-(i*100)%></option>
                        <%next%>
                      </select>
                      </font></div>
                  </td>
                </tr>
                <tr> 
                  <td width="11%" height="17"> 
                      <div align="center"><font face="Arial, Helvetica, sans-serif">
                        <input type="checkbox" name="inter" value="1">
                        15 Minute Profile</font></div>
                  </td>
                </tr>
                  <tr> 
                    <td width="11%" height="17"> 
                      <div align="center"> 
                        <input type="button" name="Button" value="View Profile" onClick="loadentry(b.value, m.value, d.value, s.value, e.value, inter.value, luid.value)">
                      </div>
                    </td>
                  </tr>
               
              </table> 
			  </form>
            </div>
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
