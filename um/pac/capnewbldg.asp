<%Option Explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
</head>
<%
dim msg
msg = request("msg")
%>
<body bgcolor="#FFFFFF" text="#000000">
<%if msg <> "" then%>
	<div align="center">
	  <font color="#FF0000" size="2"> <strong><%=msg%></strong></font> 
	</div>
<%end if%>
<form name="form2" method="post" action="capaddbldg.asp">
  <table width="100%" border="0">
    <tr bgcolor="#CCCCCC"> 
      <td width="6%" height="2">&nbsp;</td>
      <td width="22%" height="10" colspan=2><font face="Arial, Helvetica, sans-serif">Select 
        Building</font></td>
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Square 
        Feet </font></td>
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Revision</font></td>
    </tr>
    <tr> 
      <td width=6%> <input type="submit" name="choice2"  value="SAVE"> </td>
      <td width="22%" height="19" colspan=2> <font face="Arial, Helvetica, sans-serif"> 
        <%
		Dim cnn1, rst1,sql
		Set cnn1 = Server.CreateObject("ADODB.connection")
		Set rst1 = Server.CreateObject("ADODB.recordset")
		
		cnn1.Open getConnect(0,0,"engineering")
		
		sql="select b.bldgnum, strt as address, portfolioid from ["&Application("CoreIP")&"].dbCore.dbo.buildings b where bldgnum not in (select bldgnum from tlbldg tl) and offline=0 order by strt"
		rst1.Open sql, cnn1, 0, 1, 1
		if not rst1.eof then
		%>
        <select name="bldgnum" size="1">
          <option>========</option>
          <%
		  do until rst1.eof
		  %>
          <option value="<%=Trim(rst1("bldgnum"))%>|<%=Trim(rst1("address"))%>"><font face="Arial, Helvetica, sans-serif"><%=rst1("address")%> (PID:<%=rst1("portfolioid")%>) 
          </font> 
          <%
		  rst1.movenext
		  loop
		  %>
        </select>
        </font>
		<%end if%></td>
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="sqft" >
        </font></td>
      <td width="11%" height="19" > <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="rev">
        </font></td>
    </tr>
  </table>
</form>
</body>
</html>
