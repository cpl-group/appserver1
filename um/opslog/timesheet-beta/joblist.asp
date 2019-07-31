<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<script lanaguage="JavaScript" type="text/javascript">
<%
		if isempty(getKeyValue("name")) then
%>
top.location="../index.asp"
<%
		end if	

Dim cnn1, rst1
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
cnn1.Open getConnect(0,0,"intranet")

dim sqlstr, search
search = request("search")
//sqlstr = "select distinct id, description from master_job where status != 'Closed' order by id desc"
%>
</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"><title>QuickSearch - Job Log</title></head>
<body bgcolor="#dddddd" text="#000000">
<form name="searchform" method="post">
<table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#eeeeee">
<tr>
  <td bgcolor="#6699cc"><span class="standardheader">Quick Job Search</span></td>
</tr>
<tr>
  <td>
  <input type="text" name="search" value="<%=search%>">
  <input type="submit" name="submit" value="Search" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
  </td>
</tr>
<tr>
  <td>
  <%if trim(search)<>"" then
  sqlstr = "select * from Master_job where (job like '%" & search & "%' or description like '%" & search & "%' or Address_1 like '%" & search & "%' or Address_2 like '%" & search & "%'  or pm_last like '%" & search & "%' or pm_first like '%" & search & "%' or customer_name like '%" & search & "%' or status like '%" & search & "%') and status <> 'Closed' and company in('EM','NE') order by id desc"
  rst1.Open sqlstr, cnn1, 0, 1, 1
  if not rst1.eof then
  %>
  Click row to insert job into timesheet<br>
  <% end if %>
  <% end if %>
  <div id="searchdiv" style=overflow:auto;height:210px;width:100%;background-color:#eeeeee;" class="innerbody">
  <% if trim(search)<>"" then %>
  <table border=0 cellpadding="3" cellspacing="1" bgcolor="#cccccc" width="100%">
  <% if not rst1.eof then %>
  <%
    do until rst1.eof
    %>
  <tr bgcolor="#ffffff" onclick="opener.jobPicked('<%=rst1("id")%>');try{opener.setDesc('<%=rst1("id")%>','','<%=request("name")%>')}catch(exception){};self.close();" onMouseOver="this.style.backgroundColor = 'lightgreen'" onMouseOut="this.style.backgroundColor = 'white'" style="cursor:hand">
    <td><%=rst1("id")%></td>
    <td><%=rst1("Company")%></td>
    <td><%=rst1("description")%></td>
  </tr>
    <%
    rst1.movenext
    loop
  %>
  <% else %>
  <tr>
    <td bgcolor="#ffffff">&quot;<%=search%>&quot; was not found.</td>
  </tr>
  <% end if %>
  </table>
  <% end if%>
  </div>
  </td>
</tr>
<tr bgcolor="#dddddd">
  <td align="center"><input type="button" value="Close Window" onclick="self.close();" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;"></td>
</tr>
</table>
</form>

<% set cnn1 = nothing %>
</body>
</html>