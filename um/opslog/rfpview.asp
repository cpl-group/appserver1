
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script> 
function jobfolder(job){
	var jobid = new String(job)
	var dir = "data" + jobid.substr(0,1)
	var temp = "\\\\10.0.7.2\\genergy\\operations\\operations_log\\" + dir + "\\" + job

	window.open(temp,"JobFolder", "scrollbars=yes, width=500, height=300, resizeable, status" );
}
</script> 
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
rfp= Request.Querystring("rfp")
'response.write job

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rstC = Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"intranet")

sqlstr = "select Distinct customers.companyname, [employees].[First Name] + ' ' + [employees].[Last Name] AS projmanager, rfplog.* ,case when rfplog.[entry id] > 6283 then left([entry type],2)+'-00'+convert(varchar(4),[entry id]) else '00-00'+convert(varchar(4),[entry id]) end as  tjob from employees join rfplog on (employees.id=rfplog.salesmanager) join customers on (rfplog.customer=customers.customerid ) where [entry id]=" & rfp

'response.write sqlstr
'response.end
rst1.Open sqlstr, cnn1, 0, 1, 1
if rst1.eof then
%>

<table width="100%" border="0">
  <tr>
    <td>
      <div align="center">
        <p><font face="Arial, Helvetica, sans-serif"><i>RFP <%=rfp%> not found 
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

dim contactname, PhoneNumber, FaxNumber, custnum
if trim(request("cid"))<>"" then
	custnum = cINT(trim(request("cid")))
else

	if custnum="" then custnum = cINT(trim(rst1("customer")))
end if

rstC.open "select * from customers where customerid="&custnum&" order by companyname", cnn1
if not rstC.eof then
	contactname = rstC("ContactFirstName")&" "&rstC("ContactLastName")
	PhoneNumber = rstC("PhoneNumber")
	FaxNumber = rstC("FaxNumber")
else
	contactname = rst1("contact name")
	PhoneNumber = rst1("Phone Number")
	FaxNumber = rst1("Fax Number")
end if



if rst1("current status")="Proposal Inprogress" then
%>
<form name="form1" method="post" action="rfpupdate.asp">
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="2"> 
      <table width="100%" border="0">
        <tr> 
            <td height="2" width="19%"><i><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Details 
              for RFP # : <%=rst1("tjob")%>
              <input type="hidden" name="rfp" value="<%=rfp%>">
              <% if not isempty(rst1("mkid")) then%>
              MKT Ref # : <%=rst1("mkid")%> 
              <% end if%>
              <input type="hidden" name="mkid" value="<%=rst1("mkid")%>">
              </font></b></i></td>
            <td height="2" width="22%"> 
              <div align="right"><font face="Arial, Helvetica, sans-serif"><i>
                <input type="button" name="Button3" value="JOB FOLDER" onClick="jobfolder(rfp.value)">
                <input type="button" name="Button2" value="BACK" onClick="Javascript:history.back()">
                </i></font></div>
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
              <td width="26%"><font face="Arial, Helvetica, sans-serif">Customer:</font></td>
              <td width="37%"><font face="Arial, Helvetica, sans-serif" size="3">Proposal 
                Bill Type:</font></td>
              <td width="37%"><font face="Arial, Helvetica, sans-serif" size="3">Secondary 
                Proposal Bill Type:</font></td>
            </tr>
            <tr> 
              <td width="26%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="cid" onChange="document.location.href='rfpview.asp?rfp=<%=rfp%>&cid='+this.value">
                  <%Set rst6 = Server.CreateObject("ADODB.recordset")
			  str6="select distinct customerid, companyname from customers order by companyname"
			  rst6.Open str6, cnn1, 0, 1, 1
			  do until rst6.eof%>
                  <option value="<%=trim(rst6("customerid"))%>"<%if cint(trim(rst6("customerid")))=custnum then%> selected<%end if%>><font face="Arial, Helvetica, sans-serif"><%=rst6("companyname")%></font></option>
                  <%
			  rst6.movenext
			  loop
			  rst6.close
			  %>
                </select>
                </font></td>
              <td width="37%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="cost">
                  <%Set rst5 = Server.CreateObject("ADODB.recordset")
			  str5="select rfptype from rfptype "
			  rst5.Open str5, cnn1, 0, 1, 1
			  do until rst5.eof 
			  	  if rst5("rfptype")=rst1("proposal") then
			  %>
                  <option value="<%=rst5("rfptype")%>" selected><%=rst5("rfptype")%></option>
                  <%
			  	  else
			  %>
                  <option value="<%=rst5("rfptype")%>"><%=rst5("rfptype")%></option>
                  <%
			      end if
			  rst5.movenext
			  loop
			  rst5.close
			  %>
                </select>
                <%if rst1("amt")="0" then%>
                $ 
                <input type="text" name="amt" size="5" maxlength="10" value="0" >
                <%else%>
                $ 
                <input type="text" name="amt" value="<%=rst1("amt")%>" size="5" maxlength="10">
                <%end if%>
                </font></td>
              <td width="37%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="cost2">
                  <%
			  str5="select rfptype from rfptype "
			  rst5.Open str5, cnn1, 0, 1, 1
			  do until rst5.eof 
			  	  if rst5("rfptype")=rst1("proposal2") then
			  %>
                  <option value="<%=rst5("rfptype")%>" selected><font face="Arial, Helvetica, sans-serif"><%=rst5("rfptype")%></font></option>
                  <%
			  	  else
			  %>
                  <option value="<%=rst5("rfptype")%>"><font face="Arial, Helvetica, sans-serif"><%=rst5("rfptype")%></font></option>
                  <%
			      end if
			  rst5.movenext
			  loop
			  rst5.close
			  %>
                </select>
                <%if rst1("amt2")="0" then%>
                $ 
                <input type="text" name="amt2" size="5" maxlength="10" value="0" >
                <%else%>
                $ 
                <input type="text" name="amt2" value="<%=rst1("amt2")%>" size="5" maxlength="10">
                <%end if%>
                </font></td>
            </tr>
            
          </table>
				 
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="21%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif" size="3">Type:</font></td>
              <td width="27%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Contact 
                Name:</font></td>
              <td width="32%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Customer 
                Phone Number</font></td>
              <td width="20%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Customer 
                Fax Number</font></td>
            </tr>
            <tr> 
              <td width="21%"><font face="Arial, Helvetica, sans-serif">
                <select name="entrytype">
                  <% 
			  Set rst4 = Server.CreateObject("ADODB.recordset")
			  str4="SELECT [Type ID]FROM [Genergy Entry Types]where [job] =0 ORDER BY [Type ID]"
			  rst4.Open str4, cnn1, 0, 1, 1
			  do until rst4.eof
			  	  if rst4("Type ID")=rst1("entry type") then
			  %>
                  <option value="<%=rst4("Type ID")%>" selected><font face="Arial, Helvetica, sans-serif"><%=rst4("Type ID")%></font></option>
                  <%
			      else
			  %>
                  <option value="<%=rst4("Type ID")%>"><font face="Arial, Helvetica, sans-serif"><%=rst4("Type ID")%></font></option>
                  <%
			      end if
			  rst4.movenext
			  loop
			  rst4.close%>
                </select>
                </font></td>
              <td width="27%"><font face="Arial, Helvetica, sans-serif"><%=contactname%> 
                <input type="hidden" name="cname" value="<%=contactname%> ">
                </font></td>
              <td width="32%"> 
                <input type="hidden" name="customerphone" value="<%=PhoneNumber%>">
                <font face="Arial, Helvetica, sans-serif"> <%=PhoneNumber%> 
                </font></td>
              <td width="20%"> 
                <input type="hidden" name="customerfax" value="<%=FaxNumber%>">
                <font face="Arial, Helvetica, sans-serif"> <%=FaxNumber%> 
                </font></td>
            </tr>
            <tr> 
              <td width="21%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Referred 
                By</font></td>
              <td width="27%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Requested 
                By Name:</font></td>
              <td width="32%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Requested 
                By Phone Number:</font></td>
              <td width="20%">&nbsp;</td>
            </tr>
            <tr> 
              <td width="21%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="refby" value="<%=rst1("referred by")%>">
                </font></td>
              <td width="27%" height="32"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="reqname" value="<%=rst1("Requested By Name")%>">
                </font></td>
              <td width="32%" height="32"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="reqphone" value="<%=rst1("Requested By Phone")%>">
                </font></td>
              <td width="20%">&nbsp;</td>
            </tr>
          </table>
          <table width="100%" border="0">
            <tr bgcolor="#CCCCCC"> 
              <td width="19%"><font face="Arial, Helvetica, sans-serif">Floor 
                / Room</font></td>
            </tr>
            <tr> 
              <td width="19%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="floorroom">
                  <%
			  Set rst7 = Server.CreateObject("ADODB.recordset")
			  str7 = "select * from floors "
   			  rst7.Open str7, cnn1, 0, 1, 1
			  if not rst7.eof then
			      do until rst7.eof	
		       	  if rst7("floor")=rst1("floor/room") then
			  %>
                  <option value="<%=rst1("floor/room")%>" selected ><%=rst1("floor/room")%></option>
                  <%
			      else
			  %>
                  <option value="<%=rst7("floor")%>"><%=rst7("floor")%></option>
                  <%
			  	  end if
				  rst7.movenext
				  loop
			  end if
			  %>
                </select>
                </font></td>
            </tr>
          </table>
        <table width="100%" border="0">
          <tr> 
            <td><font face="Arial, Helvetica, sans-serif">Description / Comments</font></td>
          </tr>
          <tr> 
              <td valign="top"> <font face="Arial, Helvetica, sans-serif"> 
                <textarea name="description" rows="5" cols="75" ><%=rst1("description")%></textarea>
                </font></td>
          </tr>
        </table>
        <table width="100%" border="0">
          <tr bgcolor="#CCCCCC"> 
              <td width="15%" height="22"><font face="Arial, Helvetica, sans-serif">Entered 
                By</font></td>
              <td width="18%" height="22"><font face="Arial, Helvetica, sans-serif">Sales 
                Manager</font></td>
              <td width="22%" height="22"><font face="Arial, Helvetica, sans-serif">Recording 
                Date</font></td>
				
              <td width="23%"><font face="Arial, Helvetica, sans-serif">Start 
                Date (mm/dd/yyyy)</font></td>
			   <%if rst1("current status")<>"Proposal Accepted" then %>
              <td width="22%"><font face="Arial, Helvetica, sans-serif">Estimated 
                Completion Date (mm/dd/yyyy)</font></td>
				<%else%>
				<td width="22%"><font face="Arial, Helvetica, sans-serif">Actual 
                Completion Date (mm/dd/yyyy)</font></td>
				<%end if%>
          </tr>
          <tr> 
              <td width="15%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="hidden" name="EnteredBy" value="<%=rst1("Entered By")%>">
                <%=rst1("Entered By")%> </font></td>
			  <td width="18%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="mid">
                  <%
				Set rst8 = Server.CreateObject("ADODB.recordset")

				sqlstr = "select * from Managers order by lastname, firstname"
				rst8.Open sqlstr, cnn1, 0, 1, 1
				do until rst8.eof%>
					<option value="<%=rst8("mid")%>" <%If trim(rst1("salesmanager"))=trim(rst8("mid")) then%>selected<%end if%>><%=rst8("lastname")%>, <%=rst8("firstname")%></option><%
					rst8.movenext
				loop
				rst8.close
				%>
                </select>
                </font></td>
              <td width="22%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="hidden" name="recdate" value="<%=rst1("recording date")%>">
                <%=rst1("recording date")%> </font></td>
              <td width="23%"><font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="stdate" value="<%=rst1("scheduled date")%>">
                </font></td>
				 <%if rst1("current status")<>"Proposal Accepted" then %>
              <td width="22%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="enddate" value="<%=rst1("estcdate")%>">
			 
              </font></td>
			  <%else%>
				<td width="22%"><font face="Arial, Helvetica, sans-serif"> <input type="text" name="enddate" value="<%=rst1("actualcdate")%>"></font></td>
				<%end if%>
          </tr>
          <tr bgcolor="#CCCCCC"> 
              <td width="15%" height="18"><font face="Arial, Helvetica, sans-serif">Current 
                Status</font></td>
              <td width="18%" height="18"><font face="Arial, Helvetica, sans-serif">Probability</font></td>
			   <%if rst1("current status")<>"Proposal Accepted" then %>
              <td width="22%" height="18"><font face="Arial, Helvetica, sans-serif">Estimated 
                Sales Cycle (weeks)</font></td>
				 <%else%>
				
              <td width="22%"><font face="Arial, Helvetica, sans-serif">Actual 
                Sales Cycle (weeks)</font></td>
				<%end if%>
              <td width="23%" height="18"><font face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
			  <td width="22%" height="18"><font face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
          </tr>
          <tr> 
            
               
              <td width="15%"> <font face="Arial, Helvetica, sans-serif"> 
                 <select name="status" >
                  <%Set rst9 = Server.CreateObject("ADODB.recordset")
			  str9="select status from status where job=0 order by id"
			  rst9.Open str9, cnn1, 0, 1, 1
			  if not rst9.eof then
			  do until rst9.eof
			  if rst9("status")=rst1("current status") then
			  %>
                  <option value="<%=rst9("status")%>" selected ><font face="Arial, Helvetica, sans-serif"><%=rst9("status")%></font></option>
                  <%else%>
                  <option value="<%=rst9("status")%>"><font face="Arial, Helvetica, sans-serif"><%=rst9("status")%></font></option>
                  <%
				  end if
			  rst9.movenext
			  loop
			  end if
			  rst9.close%>
                </select>
                </font></td>
              <td width="18%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="prob" value="<%=rst1("probability")%>">
              </font></td>
			  <%if rst1("current status")<>"Proposal Accepted" then %>
              <td width="22%"> <font face="Arial, Helvetica, sans-serif"><%=rst1("esalescycle")%> </font></td><%else%>
				<td width="22%"><font face="Arial, Helvetica, sans-serif"> <%=rst1("asalescycle")%></font></td>
				<%end if%>
			  
			  <td width="23%"> <font face="Arial, Helvetica, sans-serif">&nbsp; </font></td>
			  <td width="22%"> <font face="Arial, Helvetica, sans-serif">&nbsp; </font></td>
          </tr>
        </table>
        <table width="100%" border="0">
          <tr> 
            <td bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Comments</font></td>
          </tr>
          <tr> 
            <td> <font face="Arial, Helvetica, sans-serif"> 
              <textarea name="comments" rows="5" cols="75"><%=rst1("comments")%></textarea>
              </font></td>
          </tr>
        </table>
          <font face="Arial, Helvetica, sans-serif"><i> </i></font> 
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="48%"> 
                <input type="submit" name="choice" value="UPDATE">
                <font face="Arial, Helvetica, sans-serif"><i> 
                <input type="button" name="Button22" value="BACK" onClick="Javascript:history.back()">
                </i></font></td>
              <td width="52%"> 
                <div align="right"> </div>
              </td>
            </tr>
          </table>
        </div>
    </td>
  </tr>
</table>

</form>
<%else%>
<form name="form1" method="post" action="rfpupdate.asp">
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="2"> 
      <table width="100%" border="0">
        <tr> 
            <td height="2" width="19%"><i><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Details 
              for RFP # : <%=rst1("tjob")%>
              <input type="hidden" name="rfp" value="<%=rfp%>">
              <% if not isempty(rst1("mkid")) then%>
              MKT Ref # : <%=rst1("mkid")%> 
              <% end if%>
              <input type="hidden" name="mkid" value="<%=rst1("mkid")%>">
              </font></b></i></td>
            <td height="2" width="22%"> 
              <div align="right"><font face="Arial, Helvetica, sans-serif"><i>
                <input type="button" name="Button3" value="JOB FOLDER" onClick="jobfolder(rfp.value)">
                <input type="button" name="Button2" value="BACK" onClick="Javascript:history.back()">
                </i></font></div>
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
              <td width="26%"><font face="Arial, Helvetica, sans-serif">Customer:</font></td>
              <td width="37%"><font face="Arial, Helvetica, sans-serif" size="3">Proposal 
                Bill Type:</font></td>
              <td width="37%"><font face="Arial, Helvetica, sans-serif" size="3">Secondary 
                Proposal Bill Type:</font></td>
            </tr>
            <tr> 
              <td width="26%"> <font face="Arial, Helvetica, sans-serif"> <%=rst1("companyname")%></font></td>
              <td width="37%"> <font face="Arial, Helvetica, sans-serif"> <%=rst1("proposal")%> 
                $ 
              <%=rst1("amt")%>
           
                </font></td>
               <td width="37%"> <font face="Arial, Helvetica, sans-serif"> <%=rst1("proposal2")%> 
                $ 
              <%=rst1("amt2")%>
           
                </font></td>
            </tr>
            
          </table>
				 
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr> 
              <td width="21%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif" size="3">Type:</font></td>
              <td width="27%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Contact 
                Name:</font></td>
              <td width="32%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Customer 
                Phone Number</font></td>
              <td width="20%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Customer 
                Fax Number</font></td>
            </tr>
            <tr> 
              <td width="21%"><font face="Arial, Helvetica, sans-serif">Please 
                Use Job Log-This is an Open Job </font></td>
              <td width="27%"><font face="Arial, Helvetica, sans-serif"><%=contactname%> 
                <input type="hidden" name="cname" value="<%=contactname%> ">
                </font></td>
              <td width="32%"> 
                <input type="hidden" name="customerphone" value="<%=PhoneNumber%>">
                <font face="Arial, Helvetica, sans-serif"> <%=PhoneNumber%> 
                </font></td>
              <td width="20%"> 
                <input type="hidden" name="customerfax" value="<%=FaxNumber%>">
                <font face="Arial, Helvetica, sans-serif"> <%=FaxNumber%> 
                </font></td>
            </tr>
            <tr> 
              <td width="21%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Referred 
                By</font></td>
              <td width="27%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Requested 
                By Name:</font></td>
              <td width="32%" bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Requested 
                By Phone Number:</font></td>
              <td width="20%">&nbsp;</td>
            </tr>
            <tr> 
              <td width="21%"> <font face="Arial, Helvetica, sans-serif"> 
               <%=rst1("referred by")%>
                </font></td>
              <td width="27%" height="32"> <font face="Arial, Helvetica, sans-serif"> 
        <%=rst1("Requested By Name")%>
                </font></td>
              <td width="32%" height="32"> <font face="Arial, Helvetica, sans-serif"> 
                <%=rst1("Requested By Phone")%>
                </font></td>
              <td width="20%">&nbsp;</td>
            </tr>
          </table>
          <table width="100%" border="0">
            <tr bgcolor="#CCCCCC"> 
              <td width="19%" height="15"><font face="Arial, Helvetica, sans-serif">Floor 
                / Room</font></td>
            </tr>
            <tr> 
              <td width="19%"> <font face="Arial, Helvetica, sans-serif"> <%=rst1("floor/room")%> 
                </font></td>
            </tr>
          </table>
        <table width="100%" border="0">
          <tr bgcolor="#CCCCCC"> 
            <td><font face="Arial, Helvetica, sans-serif">Description / Comments</font></td>
          </tr>
          <tr> 
              <td valign="top"> <font face="Arial, Helvetica, sans-serif"> 
             <%=rst1("description")%>
                </font></td>
          </tr>
        </table>
        <table width="100%" border="0">
          <tr bgcolor="#CCCCCC"> 
              <td width="15%" height="22"><font face="Arial, Helvetica, sans-serif">Entered 
                By</font></td>
              <td width="18%" height="22"><font face="Arial, Helvetica, sans-serif">Sales 
                Manager</font></td>
              <td width="22%" height="22"><font face="Arial, Helvetica, sans-serif">Recording 
                Date</font></td>
				
              <td width="23%"><font face="Arial, Helvetica, sans-serif">Start 
                Date (mm/dd/yyyy)</font></td>
			   <%if rst1("current status")<>"Proposal Accepted" then %>
              <td width="22%"><font face="Arial, Helvetica, sans-serif">Estimated 
                Completion Date (mm/dd/yyyy)</font></td>
				<%else%>
				<td width="22%"><font face="Arial, Helvetica, sans-serif">Actual 
                Completion Date (mm/dd/yyyy)</font></td>
				<%end if%>
          </tr>
          <tr> 
              <td width="15%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="hidden" name="EnteredBy" value="<%=rst1("Entered By")%>">
                <%=rst1("Entered By")%> </font></td>
			  <td width="18%"> <font face="Arial, Helvetica, sans-serif"> <%=rst1("projmanager")%> 
                </font></td>
              <td width="22%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="hidden" name="recdate" value="<%=rst1("recording date")%>">
                <%=rst1("recording date")%> </font></td>
              <td width="23%"><font face="Arial, Helvetica, sans-serif"> 
               <%=rst1("scheduled date")%>
                </font></td>
				 <%if rst1("current status")<>"Proposal Accepted" then %>
              <td width="22%"> <font face="Arial, Helvetica, sans-serif"> 
               <%=rst1("estcdate")%>
			 
              </font></td>
			  <%else%>
				<td width="22%"><font face="Arial, Helvetica, sans-serif"> <%=rst1("actualcdate")%></font></td>
				<%end if%>
          </tr>
          <tr bgcolor="#CCCCCC"> 
              <td width="15%" height="18"><font face="Arial, Helvetica, sans-serif">Current 
                Status</font></td>
              <td width="18%" height="18"><font face="Arial, Helvetica, sans-serif">Probability</font></td>
			   <%if rst1("current status")<>"Proposal Accepted" then %>
              <td width="22%" height="18"><font face="Arial, Helvetica, sans-serif">Estimated 
                Sales Cycle (weeks)</font></td>
				 <%else%>
				
              <td width="22%"><font face="Arial, Helvetica, sans-serif">Actual 
                Sales Cycle (weeks)</font></td>
				<%end if%>
              <td width="23%" height="18"><font face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
			  <td width="22%" height="18"><font face="Arial, Helvetica, sans-serif">&nbsp;</font></td>
          </tr>
          <tr> 
            
               
              <td width="15%"> <font face="Arial, Helvetica, sans-serif"> <%=rst1("current status")%> 
                </font></td>
              <td width="18%"> <font face="Arial, Helvetica, sans-serif"> 
             <%=rst1("probability")%>
              </font></td>
			  <%if rst1("current status")<>"Proposal Accepted" then %>
              <td width="22%"> <font face="Arial, Helvetica, sans-serif"><%=rst1("esalescycle")%> </font></td><%else%>
				<td width="22%"><font face="Arial, Helvetica, sans-serif"> <%=rst1("asalescycle")%></font></td>
				<%end if%>
			  
			  <td width="23%"> <font face="Arial, Helvetica, sans-serif">&nbsp; </font></td>
			  <td width="22%"> <font face="Arial, Helvetica, sans-serif">&nbsp; </font></td>
          </tr>
        </table>
        <table width="100%" border="0">
          <tr> 
            <td bgcolor="#CCCCCC"><font face="Arial, Helvetica, sans-serif">Comments</font></td>
          </tr>
          <tr> 
            <td> <font face="Arial, Helvetica, sans-serif"> 
             <%=rst1("comments")%>
              </font></td>
          </tr>
        </table>
          <font face="Arial, Helvetica, sans-serif"><i> </i></font> 
          <table width="100%" border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="48%"> 
                <input type="submit" name="choice" value="UPDATE">
                <font face="Arial, Helvetica, sans-serif"><i> 
                <input type="button" name="Button22" value="BACK" onClick="Javascript:history.back()">
                </i></font></td>
              <td width="52%"> 
                <div align="right"> </div>
              </td>
            </tr>
          </table>
        </div>
    </td>
  </tr>
</table>

</form>

<%
end if
end if
%>
</body>
</html>
