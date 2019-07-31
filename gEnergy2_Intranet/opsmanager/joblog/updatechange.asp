<%option explicit%>
<!-- #include virtual="/genergy2/secure.inc" -->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim cnn,rst,strsql,themode,change_id,jid,description
themode=trim(secureRequest("mode"))
change_id=secureRequest("change_id")
jid=secureRequest("jid")
description=secureRequest("description")


set cnn = server.createobject("ADODB.connection")
cnn.open getConnect(0,0,"intranet")
    
dim sql,amount,accepted,estimate_id
if themode="save" then
	amount=secureRequest("amount")
	if change_id="" then
		sql="insert into CHANGE_ORDER(jobno,description,amount,accepted) values('"&jid&"','"&description&"',"&amount&",0)"
	else
		accepted=secureRequest("accepted")
		if accepted="" then
			accepted="0"
		end if
	sql="update CHANGE_ORDER set description='"&description&"',amount="&amount&",accepted=" & accepted & " where id="&change_id
	'response.write sql
	'response.end
	end if
	cnn.execute sql
	%><script>
	alert("Change Order Updated")
	opener.document.location.reload()
	window.close()
	</script>
	<%
	response.End()
end if

if themode<>"new" then
  set rst = server.createobject("ADODB.recordset")
  rst.open "select description,amount,accepted from CHANGE_ORDER where id="&change_id,cnn
  description=rst("description")
  amount=rst("amount")
  if rst("accepted") then
    accepted="1"
  else
    accepted="0"
  end if
end if

'descriptiondescription
'amountamount
'acceptedaccepted%>
    <html>
    <head>
    <title>Update Change Order</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <script>
    function closepage()
    {
      if (confirm("Cancel changes?")){
        window.close()
      }
    }

    </script>
<link rel="Stylesheet" href="../../styles.css" type="text/css">   
    </head>
    <body bgcolor="#dddddd">
    <form name="form1" method="post" action="updatechange.asp">
      <input type="hidden" name="mode" value="save">
      <input type="hidden" name="change_id" value=<%=change_id%> >
      <input type="hidden" name="jid" value="<%=jid%>">
 
	<table border=0 cellpadding="3" cellspacing="0" width="100%">
    <tr valign="top" bgcolor="#6699cc">
      <td><span class="standardheader">Update Change Order</span></td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee">
      <td style="border-bottom:1px solid #cccccc;">
      <table border=0 cellpadding="3" cellspacing="0" width="100%">
      <tr>
        <td>Description:</td>
        <td><input name="description" type="text" value="<%=description%>" size="50" maxlength="50"></td>
      </tr>
      <tr>
        <td>Amount:</td>
    <%  
       Dim access
       if  allowGroups("Genergy_Corp,AR_Admin,gAccounting,IT Services") then
 	   access=1 
		%>
       <td>$<input name="amount" type="text" value="<%=amount%>" size="8" maxlength="8"></td>
        
     <%else %> 
		<td><div> $<%=amount%> </div></td> 
		<%end if%>
	 </tr>
      <tr> 
        <td>Accepted:</td>
        <td><input name="accepted" type="checkbox" value="1" ></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>
       <%if access then%> <input type="submit" value="Update">&nbsp;<%end if%><input type="button" value="Cancel" onclick="closepage();">
        </td>
      </tr>
      </table>
      </td>
    </tr>
    </table>
    <br>
    
    </form>
    </body>
    </html>