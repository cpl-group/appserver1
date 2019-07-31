<html>
<head>
<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("UM") < 5 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
		user=Session("name")
%>

<title>Billl Processor - Delete Tenant Bill</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function viewbill(ypid,bldg,luid) {
		var temp
		temp="invoiceview4del.asp?ypid=" + ypid + "&bldg="+bldg+"&luid="+luid
		document.frames.admin.location=temp
} 
function loadlist(bldg,ypid) {
	if (ypid == null){
	var temp = "tenantbilllist.asp?bldg=" + bldg 
	}else{
	var temp = "tenantbilllist.asp?bldg=" + bldg +"&ypid="+ypid
	}
	document.location = temp
}
</script>
</head>
<%
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"

		
%>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC"> 
      <div align="center"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif">Bill 
        Processor - Delete Tenant Bill</font></b></font></div>
    </td>
  </tr>
</table><table width="100%" border="0">
  <tr>
    <td width="48%" height="2"> 
	<% if isempty(Request.Querystring("bldg")) then %>
      <p> 
        <select name="bldg" onchange="loadlist(this.value)">
          <OPTGROUP label='Select Building'> 
          <%Set rst2 = Server.CreateObject("ADODB.recordset")
			sqlstr = "select bldgnum, strt from buildings order by strt"
   			rst2.Open sqlstr, cnn1, 0, 1, 1
			if not rst2.eof then
					do until rst2.eof
		%>
          <option value="<%=rst2("bldgnum")%>"><font face="Arial, Helvetica, sans-serif"><%=rst2("strt")%>,<%=rst2("bldgnum")%> </font></option>
          <%
					rst2.movenext
					loop
					end if
					rst2.close
				%>
        </select>
        <input type="button" name="Button2" value="View Building Period List" onClick="loadlist(bldg.value)">
        <% else 
	  	if isempty(Request.Querystring("ypid")) then%>
        <select name="ypid" onchange="loadlist(bldgnum.value,this.value)">
          <OPTGROUP label='Select Bill Period'> 
          <%Set rst2 = Server.CreateObject("ADODB.recordset")
			sqlstr = "SELECT DISTINCT billyrperiod.*, tblbillbyperiod.ypid FROM BillYrPeriod JOIN tblbillbyperiod ON tblbillbyperiod.ypid = billyrperiod.ypid where billyrperiod.bldgnum = '" & Request.Querystring("bldg") & "' ORDER BY tblbillbyperiod.ypid DESC"
   			rst2.Open sqlstr, cnn1, 0, 1, 1
			if not rst2.eof then
					do until rst2.eof
		%>
          <option value="<%=rst2("ypid")%>"><font face="Arial, Helvetica, sans-serif">Bill 
          Period : <%=rst2("billperiod")%>/<%=rst2("billyear")%>, <%=rst2("datestart")%> 
          to <%=rst2("dateend")%> </font></option>
          <%
					rst2.movenext
					loop
					end if
					rst2.close
				%>
        </select>
        <input type="hidden" name="bldgnum" value="<%=Request.QueryString("bldg")%>">
        <input type="button" name="Button2" value="View Tenants" onClick="loadlist(bldgnum.value,ypid.value)">
        <font face="Arial, Helvetica, sans-serif"><i> 
        <input type="button" name="Button22" value="BACK" onClick="Javascript:history.back()">
        </i></font> 
        <% 
	  else 
	  %>
        <select name="luid" >
          <optgroup label='Select Tenant'> 
          <%Set rst2 = Server.CreateObject("ADODB.recordset")
			strsql = "SELECT DISTINCT tenantnum, Billingname, leaseutilityid FROM tblBillByPeriod where bldgnum= N'" & Request.Querystring("bldg") & "' order by Billingname"

   			rst2.Open strsql, cnn1, adOpenStatic
			if not rst2.eof then
					do until rst2.eof
		%>
          <option value="<%=rst2("leaseutilityid")%>"><font face="Arial, Helvetica, sans-serif"><%=rst2("BillingName")%>(<%=rst2("TenantNum")%>)</font></option>
          <%
					rst2.movenext
					loop
					end if
					rst2.close
				%>
        </select>
        <input type="hidden" name="bldgnum" value="<%=Request.QueryString("bldg")%>">
		<input type="hidden" name="ypid" value="<%=Request.QueryString("ypid")%>">
        <input type="button" name="Button23" value="View Bill" onClick="viewbill(ypid.value,bldgnum.value,luid.value)">
        <%
	  end if 
	  end if%>
      </p>
      </td>
    <td width="52%" height="2"> 
      <div align="right">
        <input type="button" name="Submit" value="Print Building Bill" onClick='javascript:document.frames.admin.focus();document.frames.admin.print()'>
      </div>
    </td>
  </tr>
</table>
<p><IFRAME name="admin" width="100%" height="100%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME></p>
</body>
</html>