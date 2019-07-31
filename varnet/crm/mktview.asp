<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<script>

function rfpopen(frm){
    var temp = "../opslog/newcustomer.asp?mkid="+ frm.id1.value +"&cust="+ frm.cust.value
	
	window.open(temp,"", "scrollbars=no,width=820, height=400, status=no" );
}
function reload(cust,mkid){
	var temp = "mktview.asp?mkid="+mkid+"&cust="+cust
	document.location.href=temp;
}
</script>
<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<%
mkid= Request.Querystring("mkid")
cust=Request.Querystring("cust")
'response.write cust
'response.write mkid
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_main")

if isempty(cust) then
sqlstr = "select manager, contact as cust,* from mktlog m join contacts c on m.contact=c.id join salesmanagers on salesmanagers.id = m.salesmanager where m.id=" & mkid

else
sqlstr="select salesmanagers.manager as manager, c.id as cust,c.first_name+c.last_name as name,c.title as title,c.phone as phone,c.fax as fax,c.email as email,c.referredby,c.otherref as otherref,m.* from mktlog m join salesmanagers on m.salesmanager=salesmanagers.id,contacts c  where m.id=" & mkid&" and c.id="&cust
end if

'response.write sqlstr
'response.end
rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.eof then
%>


<table width="100%" border="0">
  <tr>
    <td>
      <div align="center">
        <p><font face="Arial, Helvetica, sans-serif"><i>MKT Contact<%=job%> not found 
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
<form name="form1" method="post" action="mktupdate.asp">
  <table width="100%" border="0">
    <tr> 
      <td bgcolor="#3399CC" height="2"> 
        <table width="100%" border="0">
          <tr> 
            <td height="2"><i><b><font color="#FFFFFF"><font face="Arial, Helvetica, sans-serif">Details for Contact  CRM Number:<%=mkid%></font></font></b></i></td>
            <td height="2"> 
              <div align="right"><i><b><font face="Arial, Helvetica, sans-serif"><i>
                <input type="button" name="rfp" value="OPEN AN RFP" onClick="rfpopen(document.forms['form1'])">
                </i></font><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"> 
                <input type="hidden" name="id1" value="<%=mkid%>">
                <input type="hidden" name="Button" value="Print Marketing Contact" onClick="printpo(id1.value)">
                <input type="button" name="Button2" value="BACK" onClick="Javascript:history.back()">
                </font></b></i></div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td height="143"> 
        <div align="left"> 
          <table width="100%" border="0">
            <tr bgcolor="#CCCCCC"> 
              <td width="25%"><font face="Arial, Helvetica, sans-serif">Customer</font></td>
              <td width="25%"><font face="Arial, Helvetica, sans-serif">Title</font></td>
              <td width="19%"><font face="Arial, Helvetica, sans-serif">Phone 
                Numer </font></td>
              <td width="31%"><font face="Arial, Helvetica, sans-serif">Fax Number</font></td>
            </tr>
            <tr> 
              <td width="25%" height="31"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="cust" onchange="reload(cust.value,id1.value)">
                  <%Set rst3 = Server.CreateObject("ADODB.recordset")
			sqlstr = "Select first_name,last_name,company,id from contacts order by last_name"
   			rst3.Open sqlstr, cnn1, 0, 1, 1
			if not rst3.eof then
					do until rst3.eof
					if rst1("cust")=rst3("id") then
		%>
                  <option value="<%=rst3("id") %>"selected><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><%=rst3("last_name") %>, 
                  <%=rst3("first_name")%> (<%=rst3("company")%>)</font></option>
                  <%else%>
                  <option value="<%=rst3("id") %>"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"><%=rst3("last_name") %>, 
                  <%=rst3("first_name")%> (<%=rst3("company")%>)</font></option>
                  <%
				 	end if
					rst3.movenext
					loop
					end if
					rst3.close
				%>
                </select>
                </font></td>
              <td width="25%" height="31"> <font face="Arial, Helvetica, sans-serif">
                <%=rst1("title")%>
                </font></td>
              <td width="19%" height="31"> <font face="Arial, Helvetica, sans-serif">
               <%=rst1("phone")%>
                </font></td>
              <td width="31%"><font face="Arial, Helvetica, sans-serif">
               <%=rst1("fax")%>
                </font></td>
            </tr>
            <tr bgcolor="#CCCCCC"> 
              <td width="25%"><font face="Arial, Helvetica, sans-serif">Email</font></td>
              <td width="25%"><font face="Arial, Helvetica, sans-serif">Referred 
                By:</font></td>
              <td width="19%"><font face="Arial, Helvetica, sans-serif">Start 
                Date</font></td>
              <td width="31%"><font face="Arial, Helvetica, sans-serif">Enter 
                By</font></td>
            </tr>
            <tr> 
              <td width="25%"> <font face="Arial, Helvetica, sans-serif">
                <%=rst1("email")%>
                </font></td>
              <td width="25%"> <font face="Arial, Helvetica, sans-serif"><%=rst1("referredby")%> 
               </font></td>
              <td width="19%"> <font face="Arial, Helvetica, sans-serif"><%=rst1("recordingdate")%> </font></td>
              <td width="31%"><font face="Arial, Helvetica, sans-serif"><%=rst1("enteredby")%></font></td>
            </tr>
            <tr bgcolor="#CCCCCC"> 
              <td width="25%"><font face="Arial, Helvetica, sans-serif">Situation</font></td>
			  <td width="25%"><font face="Arial, Helvetica, sans-serif">Status</font></td>
  			  <td width="25%">&nbsp;</td>
  			  <td width="25%"><font face="Arial, Helvetica, sans-serif">Sales 
                Manager </font></td>
            </tr>
            <tr> 
              <font face="Arial, Helvetica, sans-serif">
              <td width="25%" valign="top"> 
                <textarea name="sit" cols="25" rows="3" wrap="PHYSICAL"><%=rst1("situation")%> </textarea>
              </td>
              </font>
              <td width="25%">
                <textarea name="status" cols="25" rows="3" wrap="PHYSICAL"><%=rst1("status")%></textarea>
              </td>
			  <td></td>
			  <td valign="top"><font face="Arial, Helvetica, sans-serif">
<select name="manager">
                  <%Set rst3 = Server.CreateObject("ADODB.recordset")
			sqlstr = "Select * from  salesmanagers order by manager"
   			rst3.Open sqlstr, cnn1, 0, 1, 1
			if not rst3.eof then
					do until rst3.eof
		%>
                  <option value="<%=rst3("id")%>" <%if trim(rst3("id"))= trim(rst1("salesmanager")) then %> selected <% else response.write rst1("salesmanager") end if %> ><font face="Arial, Helvetica, sans-serif"><i><b><font color="#FFFFFF"><%=rst3("manager")%></font></b></i></font></option>
                  <%
				 
					rst3.movenext
					loop
					end if
					rst3.close
				%>
                </select>
                </font></td>
            </tr>
          </table>
          <font face="Arial, Helvetica, sans-serif"><i> 
          <input type="submit" name="Button5" value="UPDATE">
          <input type="button" name="Button22" value="BACK" onClick="Javascript:history.back()">
		   
          </i></font></div>
		 </td>
    </tr>
  </table>

</form>
<IFRAME name="mktitem" width="100%" height="150" src=<%="mktitems.asp?mkid="&mkid%> scrolling="auto" marginwidth="8" marginheight="16"></iframe> 
<IFRAME name="mktdetail" width="100%" height="150" src=<%="mktdetail.asp?mkid="&mkid%> scrolling="auto" marginwidth="8" marginheight="16"></iframe> 
<%
end if
rst1.close
set cnn1=nothing
%>
</body>
</html>
