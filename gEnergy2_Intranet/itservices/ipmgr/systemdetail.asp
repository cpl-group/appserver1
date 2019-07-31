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
sql2 = "SELECT * FROM systemsindex WHERE id =" & key
Set rst2 = Server.CreateObject("ADODB.Recordset")

rst2.Open sql2, cnn1, adOpenStatic, adLockReadOnly
if not rst2.EOF then
        key 			= 	rst2("id")
		serial 		= 	rst2("serial")
        systemtype 	=	rst2("systemtype")
		processor	=	rst2("processor")
		memory		=	rst2("memory")
		harddrive	=	rst2("harddrive")
		nic			= 	rst2("nic")
		video		=	rst2("video")
		monitor		=	rst2("monitor")
		note		=	rst2("note")
		
end if
rst2.close
%>
<table width="100%" border=0 cellpadding="2" cellspacing="0" bgcolor="#eeeeee">
<tr>
      <td bgcolor="#dddddd" style="border-bottom:1px solid #999999;">&nbsp; </td>
</tr>
<tr>
  <td>
  <table border=0 cellpadding="0" cellspacing="2">
          <tr>
            <td width="150"><input type="hidden" name="key" value="<%=key%>" size="6">
              Serial</td>
            <td width="150">System Type</td>
            <td width="4">&nbsp;</td>
            <td width="230">Processor</td>
            <td width="163">Memory</td>
            <td width="120">Harddrive</td>
            <td width="185">NIC</td>
            <td width="168">Video</td>
            <td width="166" nowrap>Monitor</td>
          </tr>
          <tr>
            <td><input name="serial" type="text" value="<%=serial%>" size="20" maxlength="50"></td>
            <td> 
              <input type="text" name="systemtype" value="<%=systemtype%>" size="12"> 
            </td>
            <td>&nbsp;</td>
            <td><input type="text" name="processor" value="<%=processor%>" size="12"> 
            </td>
            <td> <input type="text" name="memory" value="<%=memory%>" size="6"></td>
            <td><input type="text" name="harddrive" size="10%" value="<%=harddrive%>"></td>
            <td><input type="text" name="nic" size="10%" value="<%=nic%>"></td>
            <td><input type="text" name="video" size="10%" value="<%=video%>"> 
            </td>
            <td><input type="text" name="monitor" value="<%=monitor%>" size="6"></td>
          </tr>
          <tr valign="top"> 
            <td colspan="13">Note:</td>
          </tr>
          <tr valign="top"> 
            <td colspan="13"><input name="note" type="text" value="<%=note%>" size="100" maxlength="200"></td>
          </tr>
          <tr valign="top"> 
            <%
		  if key > 0 then 
			modify = "Update System"
		  else
		  	modify = "Save System"	
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
