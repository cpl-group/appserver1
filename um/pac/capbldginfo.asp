<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<html>
<head>
<%
		if isempty(Session("name")) then
'			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("opslog") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
		user=Session("name")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")

cnn1.Open getConnect(0,0,"engineering")
bldgnum=request("bldgnum")

sql="select * from tlbldg where bldgnum='"& bldgnum &"'"

rst1.Open sql, cnn1, 0, 1, 1	
if not rst1.eof then
	address=rst1("address")
	sqft=rst1("sqft")
	revdate=rst1("rev_date")
end if
%>

<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000">

<form name="form2" method="post" action="capbldgupdate.asp">
  <table width="100%" border="0" bgcolor="#3399CC">
    <tr >
      <td height="2" width="71%"><font face="Arial, Helvetica, sans-serif"><i><b> 
        <font color="#FFFFFF">Building Information</font></b></i></font> </td>
    </tr>
  </table> 

  <table width="100%" border="0">
    <tr bgcolor="#CCCCCC"> 
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Address</font></td>
      <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Building 
        SQFT</font></td>
		<td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Revision Date 
        </font></td>
		 <td width="11%" height="10"><font face="Arial, Helvetica, sans-serif">Revision #
        </font></td>
    </tr>
    <tr> 
	  <input type="hidden" name="bldgnum" value="<%=bldgnum%>">
      
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="hidden" name="address" value="<%=address%>"><%=address%>
        </font></td>
      <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="sqft" value="<%=sqft%>">
        </font></td>
		<td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
        <input type="text" name="revdate" value="<%=revdate%>">
       </font></td>
	   <td width="11%" height="19"> <font face="Arial, Helvetica, sans-serif"> 
	   <input type="text" name="revnum" value="<%=rst1("revision")%>">
       </font></td>
    </tr>
  </table>
  <p> 
  	<%if not(isBuildingOff(bldgnum)) then%>
    <input type="submit" name="submit" value="Update">
    <input type="button" name="submit" value="Add Floor" onclick='javascript:detail.location="capdetail.asp?item=floor&bldgnum=<%=bldgnum%>"'>
    <input type="button" name="submit" value="Add Riser" onClick='javascript:detail.location="capdetail.asp?bldgnum=<%=bldgnum%>&item=riser"'>
	<%end if%>
    <input type="button" name="submit" value="List All Risers" onClick='javascript:riser.location="capriser.asp?bldgnum=<%=bldgnum%>";floor.location="capfloor.asp?bldgnum=<%=bldgnum%>"'>
    <input type="button" name="PrintReport" value="Print Report" onClick='javascript:window.open("/eri_th/plp/padr.asp?bldgid=<%=bldgnum%>","Report")'>
    <input type="button" name="submit2" value="Back" onclick='javascript:history.back()'>
  </p>
  </form>
<IFRAME name="riser" width="100%" height="150" src="capriser.asp?bldgnum=<%=bldgnum%>" scrolling="auto" marginwidth="8" marginheight="16"></iframe>
<IFRAME name="floor" width="100%" height="150" src="capfloor.asp?bldgnum=<%=bldgnum%>" scrolling="auto" marginwidth="8" marginheight="16"></iframe> 
<IFRAME name="detail" width="100%" height="150" src="null.htm" scrolling="auto" marginwidth="8" marginheight="16"></iframe> 
</body>
</html>
