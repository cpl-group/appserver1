<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>

function processpo(poid,action, shipping) {
	if (action=="Sumbit PO for Review") {
		var poaction="submit"
	} else {
		var poaction="withdrawl"
	}	
	
	var temp = "processpo.asp?poid=" + poid + "&poaction=" + poaction
	if (shipping <= 0) {
		if(confirm("Shipping is $0, continue submission?")){
			document.location=temp	
		}
	}else {
		document.location=temp
	}
function printpo(poid){
	var temp = "poreport.asp?processpoid="+poid
	alert(temp)
	document.location=temp
}
}
function checkshipping(shipping) {

	if (shipping <= 0) {
	
	 	alert("Shipping is currently $0, please double check.")
	}

}
</script>
<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<%
PO= Request.Querystring("po")
POID = Request.Querystring("poid")
user=Session("login")
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=main;"

If POID < 1 then 

sqlstr = "select * from po where po.jobnum=substring('" & PO & "',1,4) and po.ponum=substring('" & PO & "',6,3)"

else

sqlstr = "select * from po where po.id=" & POID

end if

rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.eof then
%>

<table width="100%" border="0">
  <tr>
    <td>
      <div align="center">
        <p><font face="Arial, Helvetica, sans-serif"><i>PO<%=job%> not found 
          - please resubmit query or contact your system administrator </i></font></p>
        <p><font face="Arial, Helvetica, sans-serif"><i>
          <input type="button" name="Button" value="BACK" onclick="Javascript:history.back()">
          </i></font></p>
      </div>
    </td>
  </tr>
</table>
<%
else
%>
<form name="poform" method="post" action="poupdate.asp">
  <table width="100%" border="0">
    <tr> 
      <td bgcolor="#3399CC" height="2"> 
        <table width="100%" border="0">
          <tr> 
            <td height="2"><i><b><font color="#FFFFFF"><font face="Arial, Helvetica, sans-serif">Details 
              for PO # : <%=ponum%> 
              
              <%=rst1("jobnum")%>.<%=rst1("ponum")%> </font></font></b></i></td>
            <td height="2"> 
              <div align="right"><i><b><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">
                <input type="hidden" name="processpoid" value="<%=rst1("id")%>">
				<%
				if not rst1("submitted") and not rst1("accepted") then
				%>
				<input type="button" name="processpo" value="Sumbit PO for Review" onClick="processpo(processpoid.value,this.value,ship_amt.value)">
				<% else
				%>
				<input type="button" name="printpo" value="Print PO" onClick="printpo(processpoid.value)">
				<input type="button" name="processpo" value="Withdraw Submitted PO" onClick="processpo(processpoid.value,this.value)">
				<%
				end if%>
                <input type="button" name="goback" value="BACK" onClick="Javascript:history.back()">
                </font></b></i></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td height="2"> 
        <div align="left"> 
          <table width="100%" border="0">
            <tr bgcolor="#CCCCCC"> 
              <td width="33%"><font face="Arial, Helvetica, sans-serif">Date</font></td>
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Vendor:</font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">Job Address:</font></td>
            </tr>
            <tr> 
              <td width="33%" height="31"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="podate" value="<%=rst1("podate")%>">
                </font></td>
              <td width="37%" height="31"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="vendor" value="<%=rst1("vendor")%>" size="40" maxlength="40">
                </font></td>
              <td width="30%" height="31"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="jobaddr" value="<%=rst1("jobaddr")%>" size="40" maxlength="40">
                </font></td>
            </tr>
            <tr bgcolor="#CCCCCC"> 
              <td width="33%"><font face="Arial, Helvetica, sans-serif">Shipping 
                Address:</font></td>
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Requistioner</font></td>
              <td width="30%"><font face="Arial, Helvetica, sans-serif">PO Description</font></td>
            </tr>
            <tr valign="top"> 
              <td width="33%" height="85"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="shipaddr" value="<%=rst1("shipaddr")%>" >
                </font></td>
              <td width="37%" height="85"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="req">
                  <%Set rst2 = Server.CreateObject("ADODB.recordset")
			sqlstr = "select [first name]+' '+ [last name] as name, substring(username,9,20) as user1 from employees order by [last name]"
   			rst2.Open sqlstr, cnn1, 0, 1, 1
			if not rst2.eof then
					do until rst2.eof
					If rst1("requistioner")= rst2("user1") then	
		%>
                  <option value="<%=rst2("user1")%>"selected><font face="Arial, Helvetica, sans-serif"><%=rst2("name")%></font></option>
                  <%else
				  %>
                  <option value="<%=rst2("user1")%>"><font face="Arial, Helvetica, sans-serif"><%=rst2("name")%></font></option>
                  <%
				  end if
					rst2.movenext
					loop
					end if
					rst2.close
				%>
                </select>
                </font></td>
              <td width="30%" height="85"> <font face="Arial, Helvetica, sans-serif"> 
                <textarea name="description" cols="25" rows="3" wrap="PHYSICAL"><%=rst1("description")%></textarea>
                </font></td>
            </tr>
            <tr bgcolor="#CCCCCC"> 
              <td width="37%"><font face="Arial, Helvetica, sans-serif">Shipping 
                Amount</font></td>
              <td width="33%"><font face="Arial, Helvetica, sans-serif">PO Total</font></td>
            </tr>
            <tr> 
              <td width="37%"> <font face="Arial, Helvetica, sans-serif"> $ 
                <input type="text" name="ship_amt" value="<%=rst1("ship_amt")%>" onChange="checkshipping(this.value)">
                </font></td>
              <td width="33%"> <font face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst1("po_total"))%> 
                <input type="hidden" name="total" value="<%=FormatCurrency(rst1("po_total"))%>">
                </font></td>
              
            </tr>
          </table>
          <font face="Arial, Helvetica, sans-serif"><i> 
          <% if not rst1("submitted") and not rst1("accepted") then	%>
          <input type="submit" name="Button5" value="Update">
		  <% end if %>
          <input type="button" name="Button22" value="BACK" onClick="Javascript:history.back()">
          </i></font></div>
      </td>
    </tr>
  </table>

</form>
<IFRAME name="poitem" width="100%" height="150" src=<%="poitems.asp?poid="& rst1("id") & "&submitted=" & rst1("submitted") & "&accepted=" & rst1("accepted")%> scrolling="auto" marginwidth="8" marginheight="16"></iframe> 
<IFRAME name="podetail" width="100%" height="150" src=<%="podetail.asp?poid="& rst1("id")& "&submitted=" & rst1("submitted") & "&accepted=" & rst1("accepted")%> scrolling="auto" marginwidth="8" marginheight="16"></iframe> 
<%
end if
rst1.close
%>
</body>
</html>
