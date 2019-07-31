<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<html>
<head>
<%
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"dbCore")
%>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body bgcolor="#eeeeee" text="#000000" style="border-top:2px outset #ffffff;" class="innerbody">
<form name="form1" method="post" action="systemmodify.asp">

<%
key=cint(Request("key"))
sql2 = "SELECT * FROM ipindex WHERE id =" & key
Set rst2 = Server.CreateObject("ADODB.Recordset")

rst2.Open sql2, cnn1, adOpenStatic, adLockReadOnly
if not rst2.EOF then
        key 		= 	rst2("id")
		ip 			= 	rst2("ip")
        ipname 		=	rst2("ipname")
		systemid	=	rst2("systemid")
		userid		=	rst2("userid")
		assigndate	= 	rst2("adate")
end if
rst2.close
%>
<table width="100%" border=0 cellpadding="2" cellspacing="0" bgcolor="#eeeeee">
<tr>
      <td bgcolor="#dddddd" style="border-bottom:1px solid #999999;">&nbsp; </td>
</tr>
<tr>
  <td>
  <table border=0 cellpadding="0" cellspacing="2" width="100%">
          <tr> 
            <td width="150"><input type="hidden" name="key" value="<%=key%>" size="6">
              Assigned To</td>
            <td width="150">IP Address</td>
            <td width="230">Machine Name (DN)</td>
            <td width="163">Assigned System</td>
            <td width="120">Assign Date</td>
          </tr>
          <tr> 
            <td><input name="userid" type="text" value="<%=userid%>" size="20" maxlength="50"></td>
            <td> <input type="text" name="ip" value="<%=ip%>" size="12"> 
            </td>
            <td><input type="text" name="ipname" value="<%=ipname%>" size="12"> 
            </td>
            <td>
		  <select name="systemid">
              <% 
		if systemid <> "" then 
			strsql = "SELECT  id, serial, systemtype FROM systemsIndex where id=" & systemid
			rst2.Open strsql, cnn1, adOpenStatic
			if not rst2.eof then 
			%>
			<option value="<%=rst2("id")%>" selected><%=rst2("systemtype")%> (<%=rst2("serial")%>)</option>
			<%
			end if
			rst2.close
		end if
		%><option value="0">No System Assigned</option><%
		strsql = "SELECT  id, serial, systemtype FROM systemsIndex where id not in (select systemid from ipindex)" 
		rst2.Open strsql, cnn1, adOpenStatic
		if not rst2.eof then 
		  while not rst2.eof
		  %>
			<option value="<%=rst2("id")%>"><%=rst2("systemtype")%> (<%=rst2("serial")%>)</option>
                <% 
		  rst2.movenext
		  wend
		%>
              </select> 
              <%
		end if
		rst2.close
		%>
            </td>
            <td><input type="text" name="assigndate" size="10%" value="<%=assigndate%>"></td>
          </tr>
          <tr valign="top"> 
            <%
		  if key > 0 then 
			modify = "Update IP"
		  else
		  	modify = "Save IP"	
		  end if
		  %>
            <td colspan="13"> <input type="Submit" name="modify" value="<%=modify%>" style="border:1px outset #ddffdd;background-color:ccf3cc;font-weight:bold;"> 
              <input type="button" name="Button" value="Cancel" onclick="<%=CancelOnclick%>" style="border:1px outset #ddffdd;background-color:ccf3cc;"> 
            </td>
          </tr>
        </table>
  </td>
</tr>
</table>
	</form>
</body>
</html>
