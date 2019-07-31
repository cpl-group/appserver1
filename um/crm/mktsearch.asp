<%@Language="VBScript"%>
<%

SOption= Request.Querystring("select")
var= Request.Querystring("findvar")
var2= Request.Querystring("findvar2")
orgtype= Request.Querystring("orgtype")
org= Request.Querystring("org")

	if isempty(var) then
				msg="Please enter search and click the FIND button to begin"
				 'Write a browser-side script to update another frame (named
				 'detail) within the same frameset that displays this page.
				Response.Write "<script>" & vbCrLf
			    Response.Write "parent.location = " & _
                Chr(34) & "mktindex.asp?msg=" & msg & Chr(34) & vbCrLf
				Response.Write "</script>" & vbCrLf
	end if

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"Intranet")

	if SOption="contact" then
	    sqlstr = "select First_name, Last_name, m.id,m.recordingdate from mktlog m join contacts c on m.contact=c.id where c.id="&var
	elseif SOption="type" then
        sqlstr = "Select First_name, Last_name, m.id, m.recordingdate from mktlog m join contacts c on m.contact=c.id where org="& var
        if var2<>"1" then sqlstr = sqlstr & " and orgtype=" & var2
'	    if var = "rebny" then'
'		    sqlstr= "select m.id,m.contact_name,m.recordingdate from mktlog m join contacts c on m.contact=c.id where rebny=1"
'	    elseif var="boma" and not isempty(var) then
'		    if var2="assc_members" then
'    		    sqlstr= "select m.id,m.contact_name,m.recordingdate from mktlog m join contacts c on m.contact=c.id where assc_member=1"
'		    elseif var2="princ_members" then
'		        sqlstr= "select m.id,m.contact_name,m.recordingdate from mktlog m join contacts c on m.contact=c.id where princ_member=1"
'		    elseif var2="all" then
'		        sqlstr= "select m.id,m.contact_name,m.recordingdate from mktlog m join contacts c on m.contact=c.id where assc_member=1 or princ_member=1"
'		    end if
'		end if
'	'else
'		'sqlstr= "select m.id,m.contact_name,m.recordingdate from mktlog m join contacts c on m.contacts=c.company where assc_members=1 'or princ_members=1"
	end if
	'response.write sqlstr
	'response.end
rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic
rst1.Open sqlstr, cnn1', 0, 1, 1

if rst1.EOF then 
	msg="Last search not found...please try again"
	Response.Write "<script>" & vbCrLf
	Response.Write "parent.location = " & _
    Chr(34) & "mktindex.asp?msg=" & msg & Chr(34) & vbCrLf
	Response.Write "</script>" & vbCrLf
Else
x=0
%>
<body bgcolor="#FFFFFF">
<table width="100%" border="0">
<tr><td bgcolor="#3399CC" align="center"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF" size="4"><b>MKT CONTACT SEARCH RESULTS </b></font></td></tr>
<tr><td> 
        <table width="100%" border="0">
        <tr bgcolor="#CCCCCC"> 
            <td bgcolor="#CCCCCC" width="10%"><font face="Arial, Helvetica, sans-serif" color="#000000">MKT Number</font></td>
            <td bgcolor="#CCCCCC" width="18%"><font face="Arial, Helvetica, sans-serif" color="#000000">Contact Name </font></td>
            <td width="38%"><font face="Arial, Helvetica, sans-serif" color="#000000">Date</font></td>
        </tr>
        <%While not rst1.EOF%>
            <tr> 
                <td width="10%"><font face="Arial, Helvetica, sans-serif"><a href=<%="mktview.asp?mkid=" & rst1("id")%> ><%=rst1("id")%></a></font></td>
                <td width="18%"><font face="Arial, Helvetica, sans-serif"><%=rst1("First_name")%>&nbsp;<%=rst1("Last_name")%></font></td>
                <td width="38%"><font face="Arial, Helvetica, sans-serif"><%=rst1("recordingdate")%></font></td>
            </tr>
            <%
    		x=x+1
    		rst1.movenext
    		Wend%>
        </table></td>
</tr>
<tr><td bgcolor="#3399CC" align="center"><font face="Arial, Helvetica, sans-serif"><b><font color="#FFFFFF"><%=x%> MKTs Found </font></b></font></td></tr>
</table>
<%
end if
rst1.close
%>

