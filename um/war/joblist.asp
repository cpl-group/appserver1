<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<script lanaguage="JavaScript" type="text/javascript">
<%
		if isempty(Session("name")) then
%>
top.location="../index.asp"
<%
		else
			if Session("ts") < 4 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				Response.Redirect "../main.asp"
			end if	
		end if	

Dim cnn1, rst1
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
cnn1.Open getConnect(0,0,"Intranet")

dim sqlstr, search, tgt
search = request("search")
tgt = request("tgt")
//sqlstr = "select distinct id, description from master_job where status != 'Closed' order by id desc"
%>

function toggleHelp(){
  if (document.all.quickhelp.style.display == "none") {
    document.all.quickhelp.style.display = "inline"
  } else {
    document.all.quickhelp.style.display = "none"
  }
}

</script>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
</head>
<body bgcolor="#dddddd" text="#000000">
<form name="searchform" method="post">
<table border=0 cellpadding="3" cellspacing="0" width="100%" bgcolor="#eeeeee">
<tr bgcolor="#6699cc">
  <td><span class="standardheader">Quick Job Search</span></td>
  <td align="right"><a href="javascript:toggleHelp();" style="text-decoration:none;"><img src="/images/q-fff-69c.gif" alt="?" align="absmiddle" border="0"></a>&nbsp;</td>  
</tr>
<tr>
  <td colspan="2">
  <input type="text" name="search" value="<%=search%>">
  <input type="submit" name="submit" value="Search" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;">
  </td>
</tr>
<tr>
  <td colspan="2">
  <!-- begin quickhelp -->
  <div id="quickhelp" style="display:none;">
  Your search term can be:<br>
  <ul>
  <li>a partial or complete <b>job number</b>;
  <li>a word or phrase that may appear in a <b>job description</b> or <b>address</b> (i.e., "330 Madison");
  <li>a project manager's <b>first</b> or <b>last name</b> (but not both, as these are separate fields in the jobs database);
  <li>job <b>status</b> ("Unstarted", "In progress", "Closed").
  </ul>
  
  Use the Job Log on the intranet for more advanced searching.
  </div>
  <!-- end quickhelp -->
  <%if trim(search)<>"" then
  sqlstr = "select * from Master_job where (job like '%" & search & "%' or description like '%" & search & "%' or Address_1 like '%" & search & "%' or Address_2 like '%" & search & "%'  or pm_last like '%" & search & "%' or pm_first like '%" & search & "%' or customer_name like '%" & search & "%' or status like '%" & search & "%') order by id desc"
  rst1.Open sqlstr, cnn1, 0, 1, 1
  if not rst1.eof then
  %>
  Click a row to insert job into search field<br>
  <% end if %>
  <% end if %>
  <div id="searchdiv" style=overflow:auto;height:210px;width:100%;background-color:#eeeeee;" class="innerbody">
  <% if trim(search)<>"" then %>
  <table border=0 cellpadding="3" cellspacing="1" bgcolor="#cccccc" width="100%">
  <% if not rst1.eof then %>
  <%
    do until rst1.eof
    %>
  <tr bgcolor="#ffffff" onclick="opener.form1.<%=tgt%>.value='<%=rst1("id")%>';opener.form1.<%=tgt%>.style.backgroundColor='66ff66';self.close();" onMouseOver="this.style.backgroundColor = 'lightgreen'" onMouseOut="this.style.backgroundColor = 'white'" style="cursor:hand">
    <td><%=rst1("id")%></td>
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
  <td colspan="2" align="center"><input type="button" value="Close Window" onclick="self.close();" style="background-color:#eeeeee;border:1px outset #ffffff;color:336699;"></td>
</tr>
</table>
</form>

<% set cnn1 = nothing %>
</body>
</html>