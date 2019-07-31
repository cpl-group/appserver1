<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if request.servervariables("HTTP_REFERER")="Webster://Internal/315" and isempty(session("xmlUserObj")) then 'this is for pdf sessions
  loadNewXML("activepdf")
  loadIps(0)
end if
%>
<HTML>
<head>
<title>Usage / Cost Allocation</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head><style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #0099FF;
SCROLLBAR-HIGHLIGHT-COLOR: #0099FF;
SCROLLBAR-SHADOW-COLOR: #333333;
SCROLLBAR-3DLIGHT-COLOR: #333333;
SCROLLBAR-ARROW-COLOR: #333333;
SCROLLBAR-TRACK-COLOR: #333333;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}
-->
</style>
<style type=3D"text/css"><!--A {text-decoration: none}--></style>
<%
dim bldgnum, bp, by,sql,cnn,rs,buildingname,pdf

bldgnum	=	Request("bldgnum")
bp		=	Request("bp")
by		=	Request("by")
pdf		= 	Request("pdf")

if pdf = "" then pdf = false else pdf=true

if bldgnum = "" then 
	response.write "SYSTEM ERROR: Building code not provided."
	response.end
end if

set cnn 	= server.createobject("ADODB.Connection")
set rs 		= server.createobject("ADODB.Recordset")

cnn.Open getLocalConnect(bldgnum)

sql = "select strt from buildings where bldgnum = '" & bldgnum &"'"
rs.open sql, cnn
if rs.eof then 
	rs.close
	set cnn = nothing
	response.write "SYSTEM ERROR: Building not found."
	response.end
end if

buildingname = rs("strt")
rs.close

if bp = "" or by = "" then 
	sql = "select top 1 billyear, billperiod from tblmetersbyperiod where bldgnum = '"&bldgnum&"' order by billyear desc, billperiod desc"
	
	rs.open sql, cnn
	
	if not rs.eof then 
		bp = rs("billperiod")
		by = rs("billyear")
	else
		Response.write "SYSTEM ERROR: No bill data found"
	end if
	rs.close
end if
dim link
link = "http://pdfmaker.genergyonline.com/pdfmaker/pdfReport_v2.asp?buildpdf=true&landscape=false&devIP="& request.servervariables("SERVER_NAME")&"&sn=/genergy2/eri_TH/lmp/usgallocation.asp&qs="&server.urlencode("bldgnum="&bldgnum&"&by="&by&"&bp="&bp)
'response.write link
'response.end
Dim cmd , prm
set cmd = server.createobject("ADODB.Command")

cnn.CursorLocation = adUseClient

' set up stored proc
cmd.CommandType = adCmdStoredProc
Set cmd.ActiveConnection = cnn
' specify stored procedure to run
cmd.CommandText = "sp_tenant_usage_allocation_download"

' set parameter type and append
Set prm = cmd.CreateParameter("building", adVarChar, adParamInput, 5)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("by", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("bp", adInteger, adParamInput)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("percent", adVarChar, adParamInput, 5)
cmd.Parameters.Append prm
Set prm = cmd.CreateParameter("file", adVarChar, adParamOutput, 50)
cmd.Parameters.Append prm

cmd.Parameters("building") = bldgnum
cmd.Parameters("by") = by
cmd.Parameters("bp") = bp
cmd.Parameters("percent") = 0
cmd.execute()

%>
<script>
function closeLoadBox(name)
{   document.all[name].style.visibility="hidden";
}
function openLoadBox(name)
{   var x=Math.floor(document.body.clientWidth/2-50)
    document.all[name].style.left=x
    document.all[name].style.visibility="visible";
}
function track(e)
{   mousey = event.clientX
    mousex = event.clientY
  return true
}

function PeakCmoveprev()
{   var year = document.forms['form1'].by.value;
    var period = document.forms['form1'].bp.value
    period--;
    if(period<1)
    {   period=12;
        year--;
    }
 document.location.href='./usgallocation.asp?bldgnum=<%=bldgnum%>&bp='+period+'&by='+year;
}

function PeakCmovenext()
{   var year = document.forms['form1'].by.value;
    var period = document.forms['form1'].bp.value
    period++;
    if(period>12)
    {   period=1;
        year++;
    }
 document.location.href='./usgallocation.asp?bldgnum=<%=bldgnum%>&bp='+period+'&by='+year;
}

function PeakCnow()
{  
 document.location.href='./usgallocation.asp?bldgnum=<%=bldgnum%>';
}
</script>
<body bgcolor="#FFFFFF" text="#000000" link="#FFFFFF" vlink="#FFFFFF" alink="#0099FF">
<form name="form1" method="post" action="">
  <div align="center">
    <input type="hidden" name="bldgnum" value="<%=bldgnum%>">
    <input type="hidden" name="by" value="<%=by%>">
    <input type="hidden" name="bp" value="<%=bp%>">
    <strong><font size="3" face="Arial, Helvetica, sans-serif"><u><%=buildingname%></u><br>Usage / Cost Allocation 
    Report </font></strong> </div><br>
  <table width="650" border="0" cellspacing="0" cellpadding="0" align="center">
<tr><td width="650"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr><td width="650" height="2">
                <table width="650" cellpadding="0" cellspacing="0" border="0">
                <tr> 
                  <td bgcolor="#000000"><% if not pdf then %><b><font size="2" face="Arial, Helvetica, sans-serif" color="#FFFFFF"> 
                    <a href="javascript:PeakCmoveprev()" style="text-decoration:none;" onMouseOver="this.style.color='lightblue';" onMouseOut="this.style.color='white'">Previous 
                    Period</a> | <a href="javascript:PeakCmovenext()" style="text-decoration:none;" onMouseOver="this.style.color='lightblue';" onMouseOut="this.style.color='white';">Next 
                    Period</a> | <a href="javascript:PeakCnow()" style="text-decoration:none;" onMouseOver="this.style.color='lightblue';" onMouseOut="this.style.color= 'white';">Go 
                    To Current Period</a> </font></b><% end if%> </td>
                </tr>
		<% 
		if not pdf then %>		
                <tr>
                  <td><div align="right"><font size="2" face="Arial, Helvetica, sans-serif"><strong> <a href="#" onclick="javascript:window.open('<%=link%>','','')" style="text-decoration:none;color:black" onMouseOver="this.style.color='lightblue';" onMouseOut="this.style.color='Black'">Print to PDF</a> | <a href="<%="/eri_th/sqldownload/" & cmd.Parameters("file")%> " style="text-decoration:none;color:black" onMouseOver="this.style.color='lightblue';" onMouseOut="this.style.color='Black'">Download Data</a></strong></font></div></td>
                </tr>
	   <% end if %>
              </table>
            </td>
        </tr></table>
    </td>
</tr></table>
  &nbsp; 
 <%	set cmd = nothing %>
 <table align="center">
    <tr align="center"><td align="center">
<img name="pie" src="/genergy2/eri_th/lmp/usgallocationpie.asp?bldgnum=<%=bldgnum%>&by=<%=by%>&bp=<%=bp%>" alt="Pie Chart"><br>
<!--#INCLUDE VIRTUAL="/genergy2/eri_th/lmp/usgallocationtbl.asp"-->
</td></tr></table>
</form>
</body>
</html>