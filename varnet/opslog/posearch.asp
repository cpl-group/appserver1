<%@Language="VBScript"%>
<%

item= Request.Querystring("select")
var= Request.Querystring("findvar")

	if isempty(var) then
				msg="Please enter search and click the FIND button to begin"
				 'Write a browser-side script to update another frame (named
				 'detail) within the same frameset that displays this page.
				Response.Write "<script>" & vbCrLf
			    Response.Write "parent.location = " & _
                Chr(34) & "poindex.asp?msg=" & msg & Chr(34) & vbCrLf
				Response.Write "</script>" & vbCrLf
	end if

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_main")

	if item="jobnum" then
	sqlstr = "select employees.[first name]+ ' ' + employees.[last name] as req,ltrim(str(po.Jobnum)) + '.' + ltrim(str(po.POnum)) as ponumber, po.vendor,po.po_total,po.podate,po.requistioner from employees join po on substring(employees.username,7,20)=po.requistioner where jobnum = " & var & "order by podate desc "
	
	else
	if item = "vendor" and not isempty(var) then
		sqlstr= "select employees.[first name]+ ' ' + employees.[last name] as req,ltrim(str(po.Jobnum)) + '.' + ltrim(str(po.POnum)) as ponumber, po.vendor,po.po_total,po.podate,po.requistioner from employees join po on substring(employees.username,7,20)=po.requistioner where vendor like '%" & var & "%'order by podate desc "
		
	else
	if item="description" and not isempty(var) then
		sqlstr= "select employees.[first name]+ ' ' + employees.[last name] as req,ltrim(str(po.Jobnum)) + '.' + ltrim(str(po.POnum)) as ponumber, po.vendor,po.po_total,po.podate,po.requistioner from employees join po on substring(employees.username,7,20)=po.requistioner where description like '%" & var & "%'order by podate desc "
		
	else
	if item="requistioner" and not isempty(var) then
		sqlstr= "select employees.[first name]+ ' ' + employees.[last name] as req,ltrim(str(po.Jobnum)) + '.' + ltrim(str(po.POnum)) as ponumber, po.vendor,po.po_total,po.podate,po.requistioner from employees join po on substring(employees.username,7,20)=po.requistioner where employees.[first name]+ ' ' + employees.[last name] like '%" & var & "%'order by podate desc "
	end if
	end if
	end if
	end if
		


rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic



rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.EOF then 
	msg="Last search not found...please try again"
	Response.Write "<script>" & vbCrLf
	Response.Write "parent.location = " & _
   Chr(34) & "poindex.asp?msg=" & msg & Chr(34) & vbCrLf
	Response.Write "</script>" & vbCrLf
Else
x=0
%>
<body bgcolor="#FFFFFF">
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="30"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF" size="4">PURCHASE ORDER SEARCH RESULTS </font></b></font></div>
    </td>
  </tr>
  <tr> 
    <td> 
      <table width="100%" border="0">
        <tr bgcolor="#CCCCCC"> 
		  <td bgcolor="#CCCCCC" width="8%"><font face="Arial, Helvetica, sans-serif" color="#000000">PO 
            Number</font></td>   	
          <td bgcolor="#CCCCCC" width="20%"><font face="Arial, Helvetica, sans-serif" color="#000000">Vendor</font></td>
          <td width="13%"><font face="Arial, Helvetica, sans-serif" color="#000000">Requistioner</font></td>
          <td width="21%"><font face="Arial, Helvetica, sans-serif" color="#000000"> 
            Amount</font></td>
            
          <td width="38%"><font face="Arial, Helvetica, sans-serif" color="#000000"> 
            PO Date</font></td>
        </tr>
        <% While not rst1.EOF %>
        <tr> 
          <td width="8%"><font face="Arial, Helvetica, sans-serif"><a href=<%="poview.asp?po=" & rst1("ponumber")%> ><%=rst1("ponumber")%></a></font></td>
		  
          <td width="20%"><font face="Arial, Helvetica, sans-serif"><%=rst1("vendor")%></font></td>
          <td width="13%"><font face="Arial, Helvetica, sans-serif"><%=rst1("req")%>
		  <input type="hidden" name="job" value="<%=rst1("requistioner")%>">
		  </font></td>
		  <td width="21%"><font face="Arial, Helvetica, sans-serif"><%=FormatCurrency(rst1("po_total"))%></font></td>
		  <td width="38%"><font face="Arial, Helvetica, sans-serif"><%=rst1("podate")%></font></td>
        </tr>
        <%
		x=x+1
		rst1.movenext
		Wend
		%>
      </table>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#3399CC"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF"><%=x%> POs 
        Found </font></b></font></div>
    </td>
  </tr>
</table>
<%
end if
rst1.close
%>

