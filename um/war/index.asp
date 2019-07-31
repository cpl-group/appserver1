<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
'COMMENTS
'1/18/2008 N.Ambo removed other company options and defaulted to "GY" as the only option

dim cid,rid
'cid = request.querystring("c")
rid = request.querystring("rid")

cid = "GY" '1/18/2008 N.Ambo defaulted to Genergy because we dont use the other companies anymore

dim rst1, cnn1
set rst1 = server.createobject("ADODB.Recordset")
set cnn1 = server.createobject("ADODB.Connection")
cnn1.open getConnect(0,0,"Intranet")
%>
<html>
<head>
<title>Reports Viewer</title>
<script language="JavaScript" type="text/javascript">
function loadcompany()
{	var frm = document.forms['form1'];
	var newhref = "index.asp?c="+frm.c.value;
	document.location.href=newhref;
}

function loadreport()
{	var frm = document.forms['form1'];
	var newhref = "index.asp?c="+frm.c.value+"&rid="+frm.rid.value;
	document.location.href=newhref;
}

function loadyear()
{	var frm = document.forms['form1'];
	var newhref = "index.asp?c="+frm.c.value+"&building="+frm.building.value+"&byear="+frm.byear.value;
	document.location.href=newhref;
}

function loadperiod()
{	var frm = document.forms['form1'];
	if((frm.building.value!='')&&(frm.byear.value!='')&&(frm.bperiod.value!=''))
	{	var newhref = "bill_validation.asp?c="+frm.c.value+"&building="+frm.building.value+"&byear="+frm.byear.value+"&bperiod="+frm.bperiod.value;
		document.frames['mainval'].location=newhref;
	}
}
function setformaction(act)
{
	document.form1.action = act
}

function openwin(url,mwidth,mheight){
  window.name="opener";
  popwin = window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth);
  popwin.focus();
}

</script>
<script language="JavaScript" type="text/javascript">
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
</script>
</head>
<body bgcolor="#eeeeee">
<table width="100%" border="0" cellpadding="3" cellspacing="0">
<tr>
  <td bgcolor="#228866"><span class="standardheader">Report Viewer</span></td>
</tr>
<tr>
  <td style="border-top:1px solid #ffffff;">
  <form name="form1" target="mainval">
  <select name="c" onchange="loadcompany()">
  <%'1/18/2008 N.Ambo removed<option value="">Select Company</option> %>
  <%rst1.open "SELECT distinct company FROM report WHERE company = 'GY' ORDER BY company", cnn1
  do until rst1.eof%>
    <option value="<%=trim(rst1("company"))%>"<%if trim(rst1("company"))=trim(cid) then response.write " SELECTED"%>><%=rst1("company")%></option>
  <%	rst1.movenext
  loop
  rst1.close%>
  </select>
  <%if trim(cid)<>"" then%>
  <select name="rid" onchange="loadreport()">
  <option value="">Select Report</option>
  <%
  rst1.open "SELECT [name], id, link, security FROM report WHERE enable=1 and company='"&cid&"' ORDER BY [name]", cnn1
  Dim rname
  do until rst1.eof
    if allowGroups("Genergy_Corp,AR_Admin,gAccounting,IT Services") then 
  %>
    <option value="<%=trim(rst1("id"))%>" <%if trim(rst1("id"))=trim(rid) then 
      rname = rst1("link")%> SELECTED <% end if %>><%=trim(rst1("name"))%></option>
  <%	
    end if 
    rst1.movenext
  loop
  rst1.close
  %>
  </select>
  <%end if
  if trim(rid)<>"" then
  %>
  <script> setformaction('<%=rname%>') </script>
  <%
  
    rst1.open "SELECT prm FROM report WHERE id="&rid, cnn1
    if rst1.eof then
      response.write "<option value="""">No Parameters Defined</option>"
    else
      Dim PrmArray, y
      
      PrmArray = split(rst1("prm"),"_")
    
    rst1.close
    Dim OptArray, pArray, oArray, sqlArray, desc,prm_options,inputtype, issql
    for each pArray in PrmArray
      if pArray <> "c" then 'start parray
      rst1.open "SELECT prm_desc, prm_option, sql,inputtype FROM prm WHERE prm_name='" & parray &"'", cnn1
      desc = rst1("prm_desc")
      prm_options = rst1("prm_option")
      inputtype = rst1("inputtype")
      issql = rst1("sql")		
      rst1.close
      select case trim(inputtype) 'start inputtype
      case "select"  
        if issql then 'start sql
        sqlArray = split(prm_options,";")
        rst1.open sqlArray(0), cnn1
        %>
        <select name="<%=parray%>">
        <option value="">Select <%=desc%></option>
        <%
        do until rst1.eof%>
        <option value="<%=rst1(trim(sqlArray(1)))%>"><%=rst1(trim(sqlArray(1)))%></option></option>
        <%	rst1.movenext
        loop
        rst1.close%>
        </select>
        <%			
        else
        OptArray = split(prm_options, "_")
        %>
        <select name="<%=parray%>">
        <option value="">Select <%=desc%></option>
        <%
        for each oArray in OptArray
        %>
        <option value="<%=trim(oArray)%>"><%=trim(oArray)%></option>
        <%	
        next 
        %>
        </select>
        <%
        end if ' End SQL
      case "input"
        %>
        <input name="<%=parray%>" type="text" style="background-color:pink;" onclick="document.form1.<%=parray%>.value='';document.form1.<%=parray%>.style.backgroundColor='pink';" value="Input <%=desc%> Here" size="10">
        <%			
      end Select 'End Inputtype
      dim tgt
      tgt = parray
      end if 'End pArray 
      next 
    %>
      <input name="Submit" type="submit" value="Go!">
    <%	
      if trim(desc) = "Job Number" then
        %>
        &nbsp;&nbsp;<img src="/genergy2/setup/images/aro-rt.gif" border="0">&nbsp;<a href="javascript:openwin('joblist.asp?tgt=<%=tgt%>',260,320);">Quick job search</a>
        
        <%
      end if 'End check for job number
    end if
  end if
    
  %>
  </form>
  </td>
</tr>
</table>

<iframe src="../blank.htm" name="mainval" id="mainval" width="100%" height="90%" marginwidth="0" marginheight="0" style="background-color:white;border:1px solid #999999;" frameborder=0 border=0></iframe>
</body>
</html>
