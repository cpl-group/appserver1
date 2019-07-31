<html>
<head>
<title>MKT</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function fillup(typebox){
	document.location="MKTindex.asp?typebox=" + typebox
}

function searchjob(typebox,findvar, findvar2) {
findvar2=""
if (document.forms[0].findvar2!=null) findvar2=document.forms[0].findvar2.value
//alert(typebox)
//alert(findvar2)
//alert(document.forms[0].findvar2==null)
	var temp
	if (typebox=="mktnum")
    {    temp="mktview.asp?mkid=" + findvar +"&cust="+ findvar2
    }else if (findvar2=="update con")
    {    temp="updatecontact.asp?select=" + typebox + "&findvar="+findvar + "&findvar2=" + findvar2
	}else
    {    temp="mktsearch.asp?select=" + typebox + "&findvar="+findvar + "&findvar2=" + findvar2
	}
	//alert(temp);
	document.frames.mkt.location=temp
}
function updatecontact(typebox,findvar) {
	var temp
    temp="contactview.asp?select=" + typebox + "&cid="+findvar
	//alert(temp);
	document.frames.mkt.location=temp
}

function report(spec, job) {
document.frames.mkt.location=spec
}
</script>
</head>
<%@Language="VBScript"%>
<%
ReDim Ay(3)
ReDim By(3)
Ay(0) = "mktnum"
Ay(1) = "contact"
Ay(2) = "type"
By(0) = "MKT Number"
By(1) = "Contact"
By(2) = "Contact Type"


msg = Request.querystring("msg")
typebox = Request("typebox")
			if isempty(msg) then
				msg="Please enter search and click the FIND button to begin"
			end if
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_main")

		
%>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" bgcolor="#3399CC">
  <tr>
    <td>
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF">CRM 
        Log</font></b></font></div>
    </td>
  </tr>
</table>
<form name="form1" method="post">
<table width="100%" border="0" align="center">
<tr><td align="left"><font face="Arial, Helvetica, sans-serif">Search for Contact by: 
    <select name="typebox" size="1" onChange=fillup(this.value)>
    <%
    for i=0 to 2
	if Ay(i) = typebox then
	    %><option value="<%=typebox%>" selected><%=By(i)%></option><%
    else
	    %><option value="<%=Ay(i)%>" ><%=By(i)%></option><%
    end if
	Next
	%>
    </select>
    <%
	if typebox="type" then%>
<select name="findvar">
<%
rst1.open "select [id], org from mkt_organizations order by org", cnn1
do until rst1.eof
    response.write "<option value="""& rst1("id") &""">"& rst1("org") &"</option>"
    rst1.movenext
loop
rst1.close
%>
</select>
<select name="findvar2">
    <option value="1">All Members</option>
    <option value="2">Priciple Members</option>
    <option value="3">Associate Members</option>
</select>
        <input type="button" name="Submit" value="Find" onClick="searchjob(typebox.value,findvar.value)">
        <%
    elseif typebox="contact" then
        response.write "<select name=""findvar2"">"
        rst1.open "SELECT contacts.id, Last_Name, First_Name, Company, mkt_organizations.org from contacts inner join mkt_organizations on mkt_organizations.id=contacts.org order by Last_Name", cnn1
        do until rst1.eof
            response.write "<option value="""& rst1("id") &""">"& rst1("Last_Name") &", "& rst1("First_Name") &" ("& rst1("company") &", "& rst1("org") &")"&"</option>"
            rst1.movenext
        loop
        response.write "<input type=""button"" name=""Submit3"" value=""Find"" onClick=""searchjob(typebox.value,findvar2.value)"">"
        response.write "&nbsp;<input type=""button"" name=""Submit3"" value=""Update"" onClick=""updatecontact(typebox.value,findvar2.value)"">"
    else
        response.write "<input type=""text"" name=""findvar2"" size=""50"" maxlength=""50"">"
        response.write "<input type=""button"" name=""Submit"" value=""Find"" onClick=""searchjob(typebox.value,findvar2.value)"">"
    end if
		%>
        </font></td>
    </tr>
    <tr>
      <td align="center"><%=msg%></td>
    </tr>
    <tr> 
      <td align="left">
        
		  <input type="hidden" name="nc" value="newcontact.asp">
            <input type="hidden" name="np" value="newmkt.asp">
		  <input type="button" name="job" value="New Interaction"  onclick="report(np.value)">
        <input type="button" name="job" value="New Contact"  onclick="report(nc.value)">
      </td>
      <td> 
        <div align="right">
          <input type="button" name="print" value="Print Current View" onClick="javascript:document.frames.mkt.focus();document.frames.mkt.print()">
        </div>
      </td>
    </tr>
  </table>
  <font face="Arial, Helvetica, sans-serif"> </font> 
</form>
<IFRAME name="mkt" width="100%" height="100%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME> 
</body>
</html>