<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim id,  acctid, vendor, vendorname, serviceaddr, utility, bldgnum, Esco, locked, Escoref
acctid = Request("acctid")
bldgnum = Request("building")
utility = Request("utility")

Dim cnn1, rst1, rst2, sqlstr, pid
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
Set rst2 = Server.CreateObject("ADODB.recordset")

cnn1.Open getLocalConnect(bldgnum)
rst1.open "SELECT portfolioid FROM buildings WHERE bldgnum='"&bldgnum&"'", cnn1
if not rst1.eof then pid = rst1("portfolioid")
rst1.close
dim DBMainmodIP
DBMainmodIP = ""

if trim(acctid)<>"" then
  sqlstr= "select * from tblacctsetup where acctid='"&acctid&"' and bldgnum = '"&bldgnum&"'"
  rst1.ActiveConnection = cnn1
  rst1.Cursortype = adOpenStatic
  rst1.Open sqlstr, cnn1, 0, 1, 1
  if not rst1.eof then
    id = rst1("id")
    acctid = rst1("acctid")
    vendor = rst1("vendor")
    vendorname = rst1("vendorname")
    serviceaddr = rst1("serviceaddr")
    utility = rst1("utility")
    bldgnum = rst1("bldgnum")
    Esco = rst1("Esco")
    Escoref = rst1("Escoref")
    locked = rst1("locked")
  end if
end if
%>

<html>
<head>
<title>Edit Account Information</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/GENERGY2_INTRANET/styles.css" type="text/css">		
</head>
<script>
function meter(acctid,utility,bldg){
	var temp="meterframes.asp?acctid="+acctid+"&utility="+utility+"&bldg="+bldg
	//alert(temp)
	parent.document.frames.entry.location=temp
}
function checkescoBox(option)
{	if(option.value==0)
	{	document.forms['form1'].escoRef.disabled=false;
		document.all['escoRef'].style.backgroundColor='#FFFFFF';
	}else
	{	document.forms['form1'].escoRef.disabled=true;
		document.all['escoRef'].style.backgroundColor='#6699cc';
		document.all['escoRef'].selectedIndex=0
	}
}
</script>
<body bgcolor="#eeeeee">
<form name="form1" method="get" action="updateacct.asp">
      
  <table width="100%" border="0">
    <tr bgcolor="#6699cc"> 
      <td bgcolor="#6699cc" width="7%" height="20"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">AccountID</font></td>
      <td bgcolor="#6699cc" width="12%" height="20"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">Vendor</font></td>
      <td width="17%" height="20"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">Vendor 
        name</font></td>
      <td width="19%" height="20"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">Service 
        Address</font></td>
      <td width="18%" height="20"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">Utility</font></td>
      <td width="20%" height="20"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">Building</font></td>
    </tr>
    <tr> 
      <td width="7%"><font face="Arial, Helvetica, sans-serif"><nobr>
        <%if trim(acctid)<>"" then%>
        <%=acctid%>
        <input type="hidden" name="acctid" value="<%=acctid%>">
        <%else%>
        <input type="text" name="acctid" value="">
        <%end if%>
        </nobr></font> </td>
      <td width="12%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="vendor" value="<%=vendor%>">
        </font></td>
      <td width="17%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="name1" value="<%=vendorname%>">
        </font></td>
      <td width="19%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="addr2" value="<%=serviceaddr%>">
        </font></td>
      <td width="18%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="hidden" name="utility" value="<%=utility%>">
        <%
	sqlstr = "SELECT * FROM tblutility WHERE utilityid="&utility
    rst2.open sqlstr, cnn1
    if not rst2.eof then response.write rst2("utilitydisplay")
    rst2.close
    %>
        </font></td>
      <td width="20%"><font face="Arial, Helvetica, sans-serif"> 
        <input type="hidden" name="bldg" value="<%=bldgnum%>">
        <%=bldgnum%> </font></td>
    </tr>
    <tr> 
      <td bgcolor="#6699cc"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">Account&nbsp;Type</font></td>
      <td bgcolor="#6699cc"><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">Esco&nbsp;Reference</font></td>
      <td bgcolor="#6699cc" colspan=4><font face="Arial, Helvetica, sans-serif" color="#FFFFFF">Locked&nbsp;Account</font></td>
    </tr>
    <tr> 
      <td><select name="accounttype" onchange="checkescoBox(this)">
          <option value="0">T&nbsp;&amp;&nbsp;D</option>
          <option value="1"<%if Esco then response.write " SELECTED"%>>Esco</option>
        </select></td>
      <td> <select id="escoRef" name="escoRef" style="background-color:#white;">
          <option value="0">none</option>
          <%
			  dim str2
			  str2="SELECT * FROM tblacctsetup where Esco=1 and bldgnum='"&trim(bldgnum)&"'"
			  rst2.Open str2, cnn1, 0, 1, 1
			  do until rst2.eof
			  	response.write "<option value="""&rst2("AcctID")&""""
				if trim(rst2("AcctID"))=trim(Escoref) then response.write " SELECTED"
				response.write ">"&rst2("Vendorname")&"</option>"
				rst2.movenext
			  loop
			  
			  rst2.close
			%>
        </select> </td>
      <td colspan=4><input type="checkbox" name="locked" value="1" <%if locked="True" then response.write " CHECKED"%>></td>
    </tr>
    <tr>
      <td>
	  <%if not(isBuildingOff(bldgnum)) then%>
        <%if trim(acctid)<>"" then%>
        <input type="submit" name="action" value="UPDATE"> <input type="hidden" name="id" value="<%=id%>"> 
        <input type="hidden" name="acctid2" value="<%=acctid%>"> 
        <%else%>
        <input type="submit" name="action" value="SAVE"> 
        <%end if%>
	  <%end if%>
      </td>
      <td>&nbsp;</td>
      <td colspan=4>&nbsp;</td>
    </tr>
  </table>
	  
  
    <%
set cnn1=nothing
%></form>
<i><font face="Arial, Helvetica, sans-serif">*If this Account should be removed, 
please contact <a href ="mailto:george_nemeth@genergy.com">George Nemeth</a></font></i> 

<script>
checkescoBox(document.forms['form1'].accounttype)
</script>
</html>  
