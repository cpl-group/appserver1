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
cnn1.open  getConnect(0,0,"dbCore")

dim sqlstr, search
search = request("search")
//sqlstr = "select distinct id, description from master_job where status != 'Closed' order by id desc"
%>
</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"><title>QuickSearch - Buildings in UM</title></head>
<body bgcolor="#dddddd" text="#000000">
<form name="searchform" method="post">
<table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#eeeeee">
<tr>
      <td bgcolor="#6699cc"><span class="standardheader">Quick Building Search</span></td>
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
	sqlstr = "SELECT * FROM buildings b, portfolio p WHERE p.id=b.portfolioid AND (bldgnum like '%"& search &"%' or strt like '%"& search &"%' or bldgname like '%"& search &"%') and offline=0 order by strt"
  rst1.Open sqlstr, cnn1, 0, 1, 1
  if not rst1.eof then
  %>
  Click row to select Building Number<br>
  <% end if %>
  <% end if %>
  <% if trim(search)<>"" then %>
  <table border=0 cellpadding="3" cellspacing="1" bgcolor="#cccccc" width="100%">
  <tr>
            <td width="20%">ID #</td>
  <td width="80%">Building</td>
  </tr>
  </table>
  <div id="searchdiv" style=overflow:auto;height:210px;width:100%;background-color:#eeeeee;" class="innerbody">
  <table border=0 cellpadding="3" cellspacing="1" bgcolor="#cccccc" width="100%">
  <% if not rst1.eof then %>
  <%
    do until rst1.eof
    %>
  <tr bgcolor="#ffffff" onclick="opener.BuildingPicked('<%=rst1("bldgnum")%>');try{opener.setDesc('<%=rst1("bldgnum")%>','','<%=request("name")%>')}catch(exception){};self.close();" onMouseOver="this.style.backgroundColor = 'lightgreen'" onMouseOut="this.style.backgroundColor = 'white'" style="cursor:hand">
    <td width="20%"><%=rst1("bldgnum")%></td>
    <td width="80%"><%=rst1("strt")%></td>
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