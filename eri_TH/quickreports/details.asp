<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<title>Cost Details Report</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">		

</head>
<%
dim storedproc, by1, by2, bldg, FunctionLabel,bldglabel, rgblist,y
dim rs, cnn, cmd,prm, PerChange(12,2),rpt,NameStart,rpttype, vardollar, varpercent, applyvar, varlabel,tempArray
tempArray 		= split(trim(request("storedproc")),"|")
storedproc 		= tempArray(0)
FunctionLabel 	= tempArray(1)
bldg 		= trim(request("bldg"))
bldglabel	= ucase(trim(request("bldgname")))
by1 		= trim(request("by1"))
by2 		= trim(request("by2"))
applyvar 	= trim(request("applyvariant"))

if by2 = "" then 
	by2 = null
end if 
rpt 	= trim(request("rpt"))

NameStart = Len(rpt) + 4 

if applyvar = "true" then 
	varpercent	= trim(request("varpercent"))
	vardollar	= trim(request("vardollar"))
	if vardollar = "" then 
		vardollar = 1
		varlabel = " ADJUSTED BY " & FORMATPERCENT(varpercent,0)
	else
		if varpercent = "" then 
			varpercent = 0
			varlabel = " WITH A COST OF " & FORMATCURRENCY(VARDOLLAR) & " APPLIED."
		else
			varlabel = " ADJUSTED BY " & FORMATPERCENT(varpercent,0) & " WITH A COST OF " & FORMATCURRENCY(VARDOLLAR) & " APPLIED"
		end if
	end if
	FunctionLabel = "DETAIL VIEW FOR " & ucase(bldglabel) & " [ " & FunctionLabel & varlabel & " ]"
else
	varpercent = 0 
	vardollar = 0
	FunctionLabel = "DETAIL VIEW FOR " & ucase(bldglabel) & " [ " & FunctionLabel & " ]"
end if

%>
<body bgcolor="eeeeee">
<%
Set cnn = Server.CreateObject("ADODB.Connection")
Set rs = Server.CreateObject("ADODB.recordset")
set cmd = server.createobject("ADODB.Command")

cnn.Open getLocalConnect(bldg)
cnn.CursorLocation = adUseClient
cmd.CommandType = adCmdStoredProc
cmd.Name = "get"
if trim(by1)<>"" then
    cmd.CommandText = storedproc

    Set prm = cmd.CreateParameter("bldg", adChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("by1", adChar, adParamInput, 10)
    cmd.Parameters.Append prm
    Set prm = cmd.CreateParameter("by2", adinteger, adParamInput)
    cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("vardollar", addouble, adParamInput,18,4)
    cmd.Parameters.Append prm
	Set prm = cmd.CreateParameter("varpercent", addouble, adParamInput,18,2)
    cmd.Parameters.Append prm	
	Set prm = cmd.CreateParameter("rpttype", adVarChar, adParamOutput, 50)
	cmd.Parameters.Append prm
    Set cmd.ActiveConnection = cnn
    'return set to recordset rs
	 cmd.Parameters("bldg") = bldg
	 cmd.Parameters("by1") = by1
	 cmd.Parameters("by2") = by2
	 cmd.Parameters("vardollar") = vardollar
	 cmd.Parameters("varpercent") = varpercent
	 'response.write cmd.commandtext
	 dim i
	 for i = 0 to cmd.parameters.count-1
	'response.write ", '"&cmd.parameters(i)&"'"
	 next
	 cnn.get rs
end if
x=0
if rs.state = adstateopen then
if not rs.eof then 
rpttype = cmd.Parameters("rpttype")

%>
<table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#999999">
  <tr> 
    <td colspan=15 bgcolor="#6699cc"><span class="standardheader"> <%=ucase(FunctionLabel)%><span></span></td>
  </tr>
  <tr> 
    <td><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;</font> </td>
    <td><div align="center"><font size="1" face="Arial, Helvetica, sans-serif">Jan 
        </font></div></td>
    <td><div align="center"><font size="1" face="Arial, Helvetica, sans-serif">Feb 
        </font></div></td>
    <td><div align="center"><font size="1" face="Arial, Helvetica, sans-serif">Mar 
        </font></div></td>
    <td><div align="center"><font size="1" face="Arial, Helvetica, sans-serif">Apr 
        </font></div></td>
    <td><div align="center"><font size="1" face="Arial, Helvetica, sans-serif">May 
        </font></div></td>
    <td><div align="center"><font size="1" face="Arial, Helvetica, sans-serif">June 
        </font></div></td>
    <td><div align="center"><font size="1" face="Arial, Helvetica, sans-serif">July 
        </font></div></td>
    <td><div align="center"><font size="1" face="Arial, Helvetica, sans-serif">Aug 
        </font></div></td>
    <td><div align="center"><font size="1" face="Arial, Helvetica, sans-serif">Sept 
        </font></div></td>
    <td><div align="center"><font size="1" face="Arial, Helvetica, sans-serif">Oct 
        </font></div></td>
    <td><div align="center"><font size="1" face="Arial, Helvetica, sans-serif">Nov 
        </font></div></td>
    <td><div align="center"><font size="1" face="Arial, Helvetica, sans-serif">Dec 
        </font></div></td>
    <td><div align="center"><font size="1" face="Arial, Helvetica, sans-serif">Total 
        </font></div></td>
	<td><div align="center"><font size="1" face="Arial, Helvetica, sans-serif">&nbsp;
        </font></div></td>
  </tr>
<%			


			while not rs.EOF 
				dim label, tempdata,currentyear,totalamt
				totalamt = 0
				
					currentyear = rs("billyear")
%>
  <tr> 
    <td> <font size="1" face="Arial, Helvetica, sans-serif"><%=currentyear%></font></td>
    <%
				for y=1 to 12 
				Select Case rpttype
					Case "cost"
						if rs("p" & y) = "0" then 
							tempdata = "-"
						else
							if cdbl(rs("p"&y)) > 1000 then 
								tempdata = formatcurrency(cdbl(rs("p"&y)),0)
							else
								if rs("p" & y) < 1 and cdbl(rs("p"&y)) > 0 then 
								tempdata = formatcurrency(cdbl(rs("p"&y)),4)
								else
								tempdata = formatcurrency(cdbl(rs("p"&y)),0)
								end if
							end if
							totalamt = totalamt + tempdata
						end if
					Case "forecast"
						if rs("p" & y) = "0" then 
							tempdata = "-"
							totalamt = totalamt + PerChange(y,1)
						else
							if cdbl(rs("p"&y)) > 1000 then 
								tempdata = formatcurrency(cdbl(rs("p"&y)),0)
							else
								if rs("p" & y) < 1 and cdbl(rs("p"&y)) > 0 then 
								tempdata = formatcurrency(cdbl(rs("p"&y)),4)
								else
								tempdata = formatcurrency(cdbl(rs("p"&y)),2)
								end if
							end if
							totalamt = totalamt + tempdata
						end if
					Case Else
						if cdbl(rs("p" & y)) <= "0" then 
							tempdata = "-"
						else
							if cdbl(rs("p"&y)) > 1000 then 
								tempdata = formatnumber(cdbl(rs("p"&y)),0)
							else
								if rs("p" & y) < 1 and cdbl(rs("p"&y)) > 0 then 
								tempdata = formatnumber(cdbl(rs("p"&y)),4)
								else
								tempdata = formatnumber(cdbl(rs("p"&y)),2)
								end if
							end if
							totalamt = totalamt + tempdata
						end if 
				End Select
				PerChange(y,x) = cdbl(rs("p"&y))
%>
    <td><div align="<%if tempdata="-" then%>center<%else%>right<%end if%>"><font size="1" face="Arial, Helvetica, sans-serif"><%=tempdata%></font></div></td>
    <%				next 
				Select Case rpttype 
					case "cost"
						totalamt = formatcurrency(totalamt,0)
					case "forecast"
						totalamt = formatcurrency(totalamt,0)
					case else
						totalamt = formatnumber(totalamt,0)
				end select
	%>    <td><div align="right"><font size="1" face="Arial, Helvetica, sans-serif"><%=totalamt%></font></div></td>

    <td> <font size="1" face="Arial, Helvetica, sans-serif"><%=currentyear%></font></td>
  </tr>
  <% 
				if x = 0 then 
					by1 = rs("billyear")
				end if
				rs.movenext
				if not rs.eof then
						x=x+1
						by2 = rs("billyear")
						currentyear = rs("billyear")
				end if						
			wend 

  %>
<% if x > 0 and trim(rpttype) <> "forecast" then %>
  <tr>
    <td> <div align="center">Diff</div></td>
<%				dim amt1, amt2
				amt1 = 0
				amt2 = 0
				for y=1 to 12 
				if PerChange(y,0) = 0 then 
					tempdata = "1"
				else
					tempdata = (PerChange(y,1) - PerChange(y,0)) / PerChange(y,0) 
				end if

%>
    <td><div align="right"><font <% if tempdata < 0 then %>color="#FF0000" <% end if %>size="1" face="Arial, Helvetica, sans-serif"><%=FormatPercent(tempdata,1)%></font></div></td>
<%		
				amt1 = amt1 + PerChange(y,0)
				amt2 = amt2 + PerChange(y,1)
				next 
				if amt1=0 then totalamt=0 else totalamt = (amt2 - amt1) / amt1 
%>
	<td><%=formatpercent(totalamt,1)%></td>
    <td><div align="center">Diff</div></td>
  </tr>
<% end if %>
</table>
<% end if 
end if%>
</body>
</html>
