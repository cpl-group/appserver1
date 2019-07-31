<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%option explicit
dim pid, b, date1, rs, cnn1, sql
pid = Request.QueryString("pid")
b=Request.QueryString("b")
date1=Request.QueryString("date1")
	
Set rs = Server.CreateObject("ADODB.recordset")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_genergy1")

sql = "select *, convert(varchar,entrydate,101) as date from tblRPentries where pid='"&pid&"' and bldgnum='"&b&"' and year ='"&date1&"'"
'response.write sql
'response.end
	rs.Open sql, cnn1, 0, 1, 1
if rs.EOF then %>
<script>
function loadentry(entryid,b,pid,date1){
	var temp="unreported.asp?b=" + b + "&entryid="+entryid+"pid="+pid+"&date1="+date1+"&action=edit"
	parent.document.location=temp
}
</script>
<body bgcolor="#FFFFFF">
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF">
  <tr>
    <td>
      <div align="center"> <font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif">THERE 
        ARE NO ENTRIES FOR EXPENSES OR REVENUE</font></b></font></div>
    </td>
  </tr>
</table>
<%
else
%>
<script>
function loadentry(entryid,b,pid,date1){
	var temp="unreported.asp?building=" + b + "&entryid="+entryid+"&pid="+pid+"&date1="+date1+"&action=edit"
	parent.document.location=temp
}
</script>
<body bgcolor="#FFFFFF">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td> 
      <div align="center"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td width="11%" bgcolor="#CCCCCC"> 
              <div align="left"><font face="Arial, Helvetica, sans-serif">date</font></div>
            </td>
            <td width="9%" bgcolor="#CCCCCC"> 
              <div align="center"><font face="Arial, Helvetica, sans-serif">Exp</font></div>
            </td>
            <td width="11%" bgcolor="#CCCCCC"> 
              <div align="center"><font face="Arial, Helvetica, sans-serif">Rev</font></div>
            </td>
            <td width="21%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Description</font></td>
            <td width="5%" bgcolor="#CCCCCC"> 
              <div align="center"><font face="Arial, Helvetica, sans-serif">Period</font></div>
            </td>
            <td width="25%" bgcolor="#CCCCCC"> 
              <div align="center"><font face="Arial, Helvetica, sans-serif">Total 
                Amount</font></div>
            </td>
          </tr>
          <%while not rs.EOF 
		%>
          <form name="form1" method="post" action="">
            <tr valign="top" onmouseover="this.style.backgroundColor = 'lightgreen'" onmouseout="this.style.backgroundColor = 'white'" onclick="javascript:loadentry('<%=rs("id")%>','<%=b%>','<%=pid%>','<%=date1%>')"> 
              <td width="11%"> <font size="2"> 
                <font face="Arial, Helvetica, sans-serif"><%=rs("date")%></font></font> 
              </td>
              <td width="9%"> 
                <% if not rs("type") then %>
                <div align="center"><img src="images/greencheck.gif" width="13" height="15"> 
                </div>
                <%end if%>
              </td>
              <td width="5%"> 
                <% if rs("type") then %>
                <div align="center"><img src="images/greencheck.gif" width="13" height="15"></div>
                <%end if%>
              </td>
              <td width="25%"><font face="Arial, Helvetica, sans-serif" size="2"><%=rs("description")%></font></td>
              <td width="3%"> 
                <div align="center"><font face="Arial, Helvetica, sans-serif" size="2"><%=rs("period")%></font></div>
              </td>
              <td width="13%"> 
                <div align="right"><font face="Arial, Helvetica, sans-serif" size="2"><%=FormatCurrency(rs("amt"))%></font></div>
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

<%end if
set cnn1=nothing
rs.close
%>