
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script> 
function jobfolder(job){
	var jobid = new String(job)
	var dir = "data" + jobid.substr(0,1)
	var temp = "g:/operations/operations_log/" + dir + "/" + job

	window.open(temp,"JobFolder", "scrollbars=yes, width=500, height=300, resizeable, status" );
}
</script> 
</head>

<body bgcolor="#FFFFFF" text="#000000">
<%@Language="VBScript"%>
<%
job= Request.Querystring("job")
'response.write job

Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")



cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=sa;pwd=!general!;database=ilite;"

sqlstr = "select * from " & "[job log]" & " where " & "[entry id]" & "=" & job 

'sqlstr = "select Distinct customers.companyname, [employees].[First Name] + ' ' + [employees].[Last Name] AS projmanager, [job log].[entry type],[job log].* from main.dbo.employees join [job log] on (main.dbo.employees.id=[job log].manager) join main.dbo.customers on ([job log].customer=customers.customerid ) where [entry id]=" & job

'response.write sqlstr
'response.end

rst1.Open sqlstr, cnn1, 0, 1, 1

if rst1.eof then

%>

<table width="100%" border="0">
  <tr>
    <td>
      <div align="center">
        <p><font face="Arial, Helvetica, sans-serif"><i>Job <%=job%> not found 
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

if instr(rst1("entry type"),"RFP")=0 then 
%>
<form name="form1" method="post" action="opslogupdate.asp">
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="2"> 
      <table width="100%" border="0">
        <tr> 
            <td height="2" width="19%"><i><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Details 
              for Job # : <%=job%> 
              <input type="hidden" name="job" value="<%=job%>">
              </font></b></i></td>
			  
            <td height="2" width="29%"><i><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif"> 
              <% Set rst2 = Server.CreateObject("ADODB.recordset")
			  str="select sum(po_total) as sum1 from main.dbo.po where jobnum=" & job & " and accepted = 1"
			  rst2.Open str, cnn1, 0, 1, 1
			  if rst2("sum1") > 0 then %>
              PO Totals: <%=Formatcurrency(rst2("sum1"))%> 
              <% else%>
              PO Totals: $0 
              <% end if 
			  rst2.close
			  %>
              </font></b></i></td>
			  
            <td height="2" width="30%"><i><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif"> 
              <% Set rst3 = Server.CreateObject("ADODB.recordset")
			  str="select sum(hours) as sum1 from main.dbo.invoice_submission where jobno=" & job & " and submitted=1"
			  rst3.Open str, cnn1, 0, 1, 1
			  %>
              Total Hours Invoiced: <%=rst3("sum1")%> 
              <% rst3.close
			  %>
              </font></b></i></td>
            <td height="2" width="22%"> 
              <div align="right"><font face="Arial, Helvetica, sans-serif"><i>
                <input type="button" name="Button3" value="JOB FOLDER" onClick="jobfolder(job.value)">
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
              <td width="35%"><font face="Arial, Helvetica, sans-serif">Type: 
                </font></td>
			  <td width="35%"><font face="Arial, Helvetica, sans-serif">Billing 
                Type :</font></td>
				
              <td width="30%"><font face="Arial, Helvetica, sans-serif">Secondary 
                Billing Type:</font></td>
              
             
          </tr>
          <tr> 
              <td width="35%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="entrytype">
              <% 
			  Set rst4 = Server.CreateObject("ADODB.recordset")
			  str4="SELECT [Type ID]FROM main.dbo.[Genergy Entry Types]where job=1 ORDER BY [Type ID] "
			  rst4.Open str4, cnn1, 0, 1, 1
			  do until rst4.eof
			  	  if rst4("Type ID")=rst1("entry type") then
			  %>
			  <option value="<%=rst4("Type ID")%>" selected><%=rst4("Type ID")%></option>
			  <%
			      else
			  %>
			  <option value="<%=rst4("Type ID")%>"><%=rst4("Type ID")%></option>
			  <%
			      end if
			  rst4.movenext
			  loop
			  rst4.close%>
			  </select>
			  </font></td>
			  
			  <td width="35%"><font face="Arial, Helvetica, sans-serif"> 
                <select name="cost">
                  <%Set rst5 = Server.CreateObject("ADODB.recordset")
			  str5="select jobtype from main.dbo.jobtypes "
			  rst5.Open str5, cnn1, 0, 1, 1
			  do until rst5.eof 
			  	  if rst5("jobtype")=rst1("jobtype") then
			  %>
                  <option value="<%=rst5("jobtype")%>" selected><%=rst5("jobtype")%></option>
                  <%
			  	  else
			  %>
                  <option value="<%=rst5("jobtype")%>"><%=rst5("jobtype")%></option>
                  <%
			      end if
			  rst5.movenext
			  loop
			  rst5.close
			  %>
                </select>
                <%if rst1("amt")="0" then%>
                $ 
                <input type="text" name="amt" size="7" maxlength="7" value="0" >
                <%else%>
                $ 
                <input type="text" name="amt" value="<%=rst1("amt")%>" size="7" maxlength="7">
                <%end if%>
                </font> </td>
			  <td width="30%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="seccost">
                  <%Set rst5 = Server.CreateObject("ADODB.recordset")
			  str5="select jobtype from main.dbo.jobtypes "
			  rst5.Open str5, cnn1, 0, 1, 1
			  do until rst5.eof 
			  	  if rst5("jobtype")=rst1("sectype") then
			  %>
                  <option value="<%=rst5("jobtype")%>" selected><font face="Arial, Helvetica, sans-serif"><%=rst5("jobtype")%></font></option>
                  <%
			  	  else
			  %>
                  <option value="<%=rst5("jobtype")%>"><font face="Arial, Helvetica, sans-serif"><%=rst5("jobtype")%></font></option>
                  <%
			      end if
			  rst5.movenext
			  loop
			  rst5.close
			  %>
                </select>
                <%if rst1("amt")="0" then%>
                $ 
                <input type="text" name="secamt" size="7" maxlength="7" value="<%=rst1("secamt")%>" >
                <%else%>
                $ 
                <input type="text" name="secamt" value="<%=rst1("secamt")%>" size="7" maxlength="7">
                <%end if%>
                </font></td>
             
             
          </tr>
		  </table>
		 <table width="100%" border="0">
            <tr bgcolor="#CCCCCC"> 
              <td width="21%"><font face="Arial, Helvetica, sans-serif">Customer:</font></td>
              <td width="24%"><font face="Arial, Helvetica, sans-serif">Contact 
                Name:</font></td>
              <td width="25%"><font face="Arial, Helvetica, sans-serif">Customer 
                Phone #</font></td>
			  <td width="30%"><font face="Arial, Helvetica, sans-serif">Fax Number</font></td>
            </tr>
            <tr> 
              <td width="21%" height="32"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="cid">
                  <%Set rst6 = Server.CreateObject("ADODB.recordset")
			  str6="select distinct customerid, companyname from main.dbo.customers order by companyname"
			  rst6.Open str6, cnn1, 0, 1, 1
			  do until rst6.eof 
			  	  if rst6("customerid")=rst1("customer") then
			  %>
                  <option value="<%=rst6("customerid")%>" selected><font face="Arial, Helvetica, sans-serif"><%=rst6("companyname")%></font></option>
                  <%
			  	  else
			  %>
                  <option value="<%=rst6("customerid")%>"><font face="Arial, Helvetica, sans-serif"><%=rst6("companyname")%></font></option>
                  <%
			      end if
			  rst6.movenext
			  loop
			  rst6.close
			  %>
                </select>
                </font></td>
              <td width="24%" height="32"> <font face="Arial, Helvetica, sans-serif"> 
                <%=rst1("contact name")%> 
                <input type="hidden" name="cname" value="<%=rst1("contact name")%> ">
                </font></td>
              <td width="25%" height="32"><font face="Arial, Helvetica, sans-serif"><%=rst1("Phone Number")%></font> 
              </td>
				 <td width="30%" height="32"> <font face="Arial, Helvetica, sans-serif"><%=rst1("Fax Number")%> 
                </font></td>
            </tr>
            <tr bgcolor="#CCCCCC"> 
              <td width="21%"><font face="Arial, Helvetica, sans-serif">Requested 
                By Name:</font></td>
              <td width="24%"><font face="Arial, Helvetica, sans-serif">Requested 
                By Phone #:</font></td>
              <td width="25%"><font face="Arial, Helvetica, sans-serif">Referred 
                By</font></td>
				 
              <td width="30%"><font face="Arial, Helvetica, sans-serif">Floor 
                / Room</font></td>
            </tr>
            <tr> 
              <td width="21%"><font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="reqname" value="<%=rst1("Requested By Name")%>">
                </font> </td>
              <td width="24%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="reqphone" value="<%=rst1("Requested By Phone")%>">
                </font></td>
              <td width="25%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="refby" value="<%=rst1("referred by")%>">
                </font></td>
				 
              <td width="30%"><font face="Arial, Helvetica, sans-serif"> 
                <select name="floorroom">
                  <%
			  Set rst7 = Server.CreateObject("ADODB.recordset")
			  str7 = "select * from main.dbo.floors "
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
              <td width="18%" height="22"><font face="Arial, Helvetica, sans-serif">Project 
                Manager</font></td>
              <td width="22%" height="22"><font face="Arial, Helvetica, sans-serif">Recording 
                Date</font></td>
				
              <td width="23%"><font face="Arial, Helvetica, sans-serif">Start 
                Date (mm/dd/yyyy)</font></td>
			   
              <td width="22%"><font face="Arial, Helvetica, sans-serif">End Date (mm/dd/yyyy)</font></td>
          </tr>
          <tr> 
              <td width="15%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="hidden" name="EnteredBy" value="<%=rst1("Entered By")%>">
                <%=rst1("Entered By")%> </font></td>
			  <td width="18%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="mid">
                  <%
				Set rst8 = Server.CreateObject("ADODB.recordset")

				sqlstr = "select * from main.dbo.Managers order by lastname, firstname"
				rst8.Open sqlstr, cnn1, 0, 1, 1
				do until rst8.eof%>
					<option value="<%=rst8("mid")%>" <%If trim(rst1("manager"))=trim(rst8("mid")) then%>selected<%end if%>><%=rst8("lastname")%>, <%=rst8("firstname")%></option><%
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
				
              <td width="22%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="text" name="enddate" value="<%=rst1("requested target date")%>">
			 
              </font></td>
          </tr>
          <tr bgcolor="#CCCCCC"> 
              <td width="15%" height="18"><font face="Arial, Helvetica, sans-serif">Current 
                Status</font></td>
              <td width="18%" height="18"><font face="Arial, Helvetica, sans-serif">% 
                completed </font></td>
              <td width="22%" height="18"><font face="Arial, Helvetica, sans-serif">Last 
                Bill Date  </font></td>
				 
              <td width="23%" height="18"><font face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif">Ref. Job 
                #</font></font></td>
			  <td width="22%" height="18"><font face="Arial, Helvetica, sans-serif"></font></td>
          </tr>
          <tr> 
            
               
              <td width="15%"> <font face="Arial, Helvetica, sans-serif"> 
                <select name="status">
                  <%Set rst9 = Server.CreateObject("ADODB.recordset")
			  str9="select distinct status from main.dbo.status where job=1 order by status desc"
			  rst9.Open str9, cnn1, 0, 1, 1
			  if not rst9.eof then
			  do until rst9.eof
			  if rst9("status")=rst1("current status") then
			  %>
                  <option value="<%=rst9("status")%>" selected ><font face="Arial, Helvetica, sans-serif"><%=rst9("status")%></font></option>
                  <%else
					  if rst9("status") ="Closed" and Session("opslog") =5 then
						%>
					  <option value="<%=rst9("status")%>"><font face="Arial, Helvetica, sans-serif"><%=rst9("status")%></font></option>
					  <%
					  else 
						  if rst9("status") <> "Closed" and rst1("current status") <> "Closed" then 
						  %>
						  <option value="<%=rst9("status")%>"><font face="Arial, Helvetica, sans-serif"><%=rst9("status")%></font></option>
						  <%
						  else
						  	if rst1("current status")="Closed" and session("opslog")=5 then
							 %>
						  <option value="<%=rst9("status")%>"><font face="Arial, Helvetica, sans-serif"><%=rst9("status")%></font></option>
						  <%
							end if 
						  end if
					  end if
			end if
			  rst9.movenext
			  loop
			  end if
			  rst9.close%>
                </select>
                </font></td>
           ss
				<td width="35%"><font face="Arial, Helvetica, sans-serif"> 
                <select name="percentcomp">
                  <%Set rst10 = Server.CreateObject("ADODB.recordset")
			  str10="select [percent] from main.dbo.percents order by id "
			  rst10.Open str10, cnn1, 0, 1, 1
			  do until rst10.eof 
			  	  if rst10("percent")=rst1("% completed")then
			  %>
                  <option value="<%=rst10("percent")%>" selected><%=rst10("percent")%></option>
                  <%
			  	  else
			  %>
                  <option value="<%=rst10("percent")%>"><%=rst10("percent")%></option>
                  <%
			      end if
			  rst10.movenext
			  loop
			  rst10.close
			  %>
                </select>
               
              </font></td>
              <td width="22%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="hidden" name="billdate" value="<%=rst1("billdate")%>">
			  <%=rst1("billdate")%>
              </font></td>
			  
			  <td width="23%"> <font face="Arial, Helvetica, sans-serif">
                <input type="text" name="refnum" value="<%=rst1("ChgOrderRefNum")%>" maxlength="5" size="5">
                </font></td>
			  <td width="22%"> <font face="Arial, Helvetica, sans-serif"> </font></td>
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
                <input type="submit" name="choice" value="Update">
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

<% else%>

<form name="form1" method="post" action="opslogupdate.asp">
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="2"> 
      <table width="100%" border="0">
        <tr> 
            <td height="2" width="19%"><i><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif">Details 
              for Job # : <%=job%> 
              <input type="hidden" name="job" value="<%=job%>">
              </font></b></i></td>
			  
            <td height="2" width="29%"><i><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif"> 
              <% Set rst2 = Server.CreateObject("ADODB.recordset")
			  str="select sum(po_total) as sum1 from main.dbo.po where jobnum=" & job & " and accepted = 1"
			  rst2.Open str, cnn1, 0, 1, 1
			  if rst2("sum1") > 0 then %>
              PO Totals: <%=Formatcurrency(rst2("sum1"))%> 
              <% else%>
              PO Totals: $0 
              <% end if 
			  rst2.close
			  %>
              </font></b></i></td>
			  
            <td height="2" width="30%"><i><b><font color="#FFFFFF" face="Arial, Helvetica, sans-serif"> 
              <% Set rst3 = Server.CreateObject("ADODB.recordset")
			  str="select sum(hours) as sum1 from main.dbo.invoice_submission where jobno=" & job & " and submitted=1"
			  rst3.Open str, cnn1, 0, 1, 1
			  %>
              Total Hours Invoiced: <%=rst3("sum1")%> 
              <% rst3.close
			  %>
              </font></b></i></td>
            <td height="2" width="22%"> 
              <div align="right"><font face="Arial, Helvetica, sans-serif"><i>
                <input type="button" name="Button3" value="JOB FOLDER" onClick="jobfolder(job.value)">
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
              <td width="35%"><font face="Arial, Helvetica, sans-serif">Type:</font></td>
			  <td width="35%"><font face="Arial, Helvetica, sans-serif">Billing 
                Type :</font></td>
				
              <td width="30%"><font face="Arial, Helvetica, sans-serif">Secondary 
                Billing Type:</font></td>
              
             
          </tr>
          <tr> 
              <td width="35%"> <font face="Arial, Helvetica, sans-serif"> Please 
                Use RFP LOG - This is Currently an RFP </font></td>
			  <td width="35%"><font face="Arial, Helvetica, sans-serif"> <%=rst1("jobtype")%>
               
                <%if rst1("amt")="0" then%>
                $ 0
                
                <%else%>
                $ 
              <%=rst1("amt")%>
                <%end if%>
                </font> </td>
			  <td width="30%"> <font face="Arial, Helvetica, sans-serif"> <%=rst1("sectype")%>
                <%if rst1("secamt")="0" then%>
                $ 0 
                <%else%>
                $ <%=rst1("secamt")%> 
                <%end if%>
                </font></td>
             
             
          </tr>
		  </table>
		 <table width="100%" border="0">
            <tr bgcolor="#CCCCCC"> 
              <td width="21%"><font face="Arial, Helvetica, sans-serif">Customer:</font></td>
              <td width="24%"><font face="Arial, Helvetica, sans-serif">Contact 
                Name:</font></td>
              <td width="25%"><font face="Arial, Helvetica, sans-serif">Customer 
                Phone #</font></td>
			  <td width="30%"><font face="Arial, Helvetica, sans-serif">Fax Number</font></td>
            </tr>
            <tr> 
              <td width="21%" height="32"> <font face="Arial, Helvetica, sans-serif">               
                 <%=rst1("companyname")%>
                </font></td>
              <td width="24%" height="32"> <font face="Arial, Helvetica, sans-serif"> 
                <%=rst1("contact name")%> 
                <input type="hidden" name="cname" value="<%=rst1("contact name")%> ">
                </font></td>
              <td width="25%" height="32"><font face="Arial, Helvetica, sans-serif"><%=rst1("Phone Number")%></font> 
              </td>
				 <td width="30%" height="32"> <font face="Arial, Helvetica, sans-serif"><%=rst1("Fax Number")%> 
                </font></td>
            </tr>
            <tr bgcolor="#CCCCCC"> 
              <td width="21%"><font face="Arial, Helvetica, sans-serif">Requested 
                By Name:</font></td>
              <td width="24%"><font face="Arial, Helvetica, sans-serif">Requested 
                By Phone #:</font></td>
              <td width="25%"><font face="Arial, Helvetica, sans-serif">Referred 
                By</font></td>
				 
              <td width="30%"><font face="Arial, Helvetica, sans-serif">Floor 
                / Room</font></td>
            </tr>
            <tr> 
              <td width="21%"><font face="Arial, Helvetica, sans-serif"> 
               <%=rst1("Requested By Name")%>
                </font> </td>
              <td width="24%"> <font face="Arial, Helvetica, sans-serif"> 
               <%=rst1("Requested By Phone")%>
                </font></td>
              <td width="25%"> <font face="Arial, Helvetica, sans-serif"> 
                <%=rst1("referred by")%>
                </font></td>
				 
              <td width="30%"><font face="Arial, Helvetica, sans-serif"> 
                
                 <%=rst1("floor/room")%> </font></td>
            </tr>
				</table>
        <table width="100%" border="0">
          <tr> 
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
              <td width="18%" height="22"><font face="Arial, Helvetica, sans-serif">Project 
                Manager</font></td>
              <td width="22%" height="22"><font face="Arial, Helvetica, sans-serif">Recording 
                Date</font></td>
				
              <td width="23%"><font face="Arial, Helvetica, sans-serif">Start 
                Date (mm/dd/yyyy)</font></td>
			   
              <td width="22%"><font face="Arial, Helvetica, sans-serif">End Date (mm/dd/yyyy)</font></td>
          </tr>
          <tr> 
              <td width="15%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="hidden" name="EnteredBy" value="<%=rst1("Entered By")%>">
                <%=rst1("Entered By")%> </font></td>
			  <td width="18%"> <font face="Arial, Helvetica, sans-serif"> <%=rst1("projmanager")%></font></td>
              <td width="22%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="hidden" name="recdate" value="<%=rst1("recording date")%>">
                <%=rst1("recording date")%> </font></td>
              <td width="23%"><font face="Arial, Helvetica, sans-serif"> 
                <%=rst1("scheduled date")%>
                </font></td>
				
              <td width="22%"> <font face="Arial, Helvetica, sans-serif"> 
                <%=rst1("requested target date")%>
			 
              </font></td>
          </tr>
          <tr bgcolor="#CCCCCC"> 
              <td width="15%" height="18"><font face="Arial, Helvetica, sans-serif">Current 
                Status</font></td>
              <td width="18%" height="18"><font face="Arial, Helvetica, sans-serif">% 
                completed </font></td>
              <td width="22%" height="18"><font face="Arial, Helvetica, sans-serif">Last 
                Bill Date  </font></td>
				 
              <td width="23%" height="18"><font face="Arial, Helvetica, sans-serif"><font face="Arial, Helvetica, sans-serif">Ref. Job 
                #</font></font></td>
			  <td width="22%" height="18"><font face="Arial, Helvetica, sans-serif"></font></td>
          </tr>
          <tr> 
            
               
              <td width="15%"> <font face="Arial, Helvetica, sans-serif"> 
         
                <%=rst1("Current Status")%> 
                
                </font></td>
            
              <td width="18%"> <font face="Arial, Helvetica, sans-serif"> 
			  <input type="hidden" name="percentcomp" value="<%=rst1("% completed")%>">
               <%=rst1("% completed")%>
              </font></td>
              <td width="22%"> <font face="Arial, Helvetica, sans-serif"> 
                <input type="hidden" name="billdate" value="<%=rst1("billdate")%>">
			  <%=rst1("billdate")%>
              </font></td>
			  
			  <td width="23%"><input type="hidden" name="refnum" value="<%=rst1("ChgOrderRefNum")%>">
               <%=rst1("ChgOrderRefNum")%>
                </td>
			  <td width="22%"> <font face="Arial, Helvetica, sans-serif"> </font></td>
          </tr>
        </table>
        <table width="100%" border="0">
          <tr> 
              <td bgcolor="#CCCCCC" height="18"><font face="Arial, Helvetica, sans-serif">Comments</font></td>
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
                <input type="hidden" name="choice" value="Update">
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
