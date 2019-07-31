<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->

<%
'12.14.2007 N.Ambo made change for caclulating admin/fee. Adminfee value in the line detail part of the bill was being calulated differntly from the right-hand side. 
'Admin fee in line detail was calculated as : (Energy + demand –credit) * Admin Fee percentage; it has now been changed to remove the -credit in the calculation
'It now uses value admindollar from the query which represents (energy+demand)*admin fee
function getNumber(number)
'	response.write "|"&number&"|"
	if not(isNumeric(number)) then number = 0
	getNumber = number
end function

dim bperiod, building, byear, pid,rpt, pdf, Genergy_Users, utilityid,demo, sql,strt, utilityname


bperiod = request("bperiod")
building = request("building")
byear = request("byear")
pid = request("pid")
'if pid = "" then pid = getpid(building) end if
Genergy_Users = request("Genergy_Users")
utilityid = trim(request("utilityid"))
utilityname = request.querystring("utilityname")
strt=request.querystring("strt")
demo = request("demo")
if demo = "" then demo = false end if
dim pdfsession
pdfsession = request("fdp")
if request.servervariables("HTTP_REFERER")="Webster://Internal/315" and isempty(session("xmlUserObj")) or ( pdfsession ="pdffdp" ) then 'this is for pdf sessions
  loadNewXML("activepdf")
  loadIps(0)
	if Genergy_Users="True" then setGroup("Genergy Users")
	pdf = true
end if
%>
			<link rel='Stylesheet' href='/genergy2/styles1.css' type='text/css'>
			<table width='100%' border='0' bgcolor='#FFFFFF'>
				<tr>
					<td width='90%' valign='top' align='center'><font size="-1"><b><%=strt%></b><br>Submetering Summary Report</td>
					<td width='10%' valign='bottom'>&nbsp;<br>&nbsp;
						<table border='0' cellspacing='2' cellpadding='3' bgcolor='#000000'>
							<tr bgcolor='white'>
								<td align='center'><font size='-4'>Bill&nbsp;Year</td>
								<td align='center'><font size='-4'>Bill&nbsp;Period</td>
								<%if utilityname<>"" then %> <td align='center'><font size='-4'>Utility</td> <%else%></tr> <% end if %>
							<tr bgcolor='white'>
								<td align='center'><font size='-4'><%=byear%></td>
								<td align='center'><font size='-4'><%=bperiod%></td>
								<%if utilityname<>"" then %> <td align='center' nowrap><font size='-4'><%=utilityname%></td> <%else%></tr><% end if %>
						</table>
						<table border='0' cellspacing='0' cellpadding='0'>
							<tr height='8'>
								<td>&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>