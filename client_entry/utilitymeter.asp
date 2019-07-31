<html>
<head>
<title>Utility Meters</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<%@Language="VBScript"%>

<%
acctid=Request.querystring("acctid")
bldg=Request.querystring("bldg")
utility=Request.querystring("utility")
id1=Request.querystring("meterid")
flag=Request.querystring("flag")
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open application("cnnstr_genergy1")


sqlstr= "select * from meters1 where meterid='" &id1& "' "

rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic

rst1.Open sqlstr, cnn1, 0, 1, 1
if not rst1.EOF and isempty(flag) then%>
<body bgcolor="#FFFFFF">
<form name="detail" method="post" action=" updmeter.asp">
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="30" width="89%"> 
      <div align="left"><i><b><font face="Arial, Helvetica, sans-serif" size="3"><font color="#FFFFFF">Account 
        Number <%=Request.querystring("acctid")%>: 
        <input type="hidden" name="bldg2" value="<%=rst1("bldgnum")%>">
        <input type="hidden" name="id2" value="<%=rst1("meterid")%>">
        <input type="hidden" name="acct2" value="<%=rst1("acctid")%>">
        Meter: <%=rst1("meternum")%></font><font face="Arial, Helvetica, sans-serif" size="3" color="#FFFFFF">, 
        Utility: 
        <input type="hidden" name="utility2" value="<%=Request.querystring("utility")%>">
        <% =Request.querystring("utility")%>
        </font><font face="Arial, Helvetica, sans-serif" size="2"><font color="#FFFFFF"><font size="3"> 
        </font></font></font></font></b></i></div>
    </td>
  </tr>
</table>
 
  <table width="100%" border="0">
    <tr> 
      <td width="21%"> <font color="#000000" face="Arial, Helvetica, sans-serif" size="2">Start 
        Date:</font><font face="Arial, Helvetica, sans-serif" size="2"></font></td>
      <td width="36%"><font color="#000000" face="Arial, Helvetica, sans-serif" size="2">
        <input type="text" name="sd2" value="<%=rst1("datestart")%>" maxlength="10" size="11">
        </font><font face="Arial, Helvetica, sans-serif" size="2"> </font></td>
      <td width="43%"> 
        <div align="left"><font color="#000000" face="Arial, Helvetica, sans-serif" size="2">Date 
          Off:</font> <font color="#000000" face="Arial, Helvetica, sans-serif" size="2"> 
          <input type="text" name="dof2" value="<%=rst1("dateoffline")%>" size="11" maxlength="10">
          </font></div>
      </td>
    </tr>
    <tr> 
      <td width="21%"><font color="#000000" face="Arial, Helvetica, sans-serif" size="2">Location:</font><font face="Arial, Helvetica, sans-serif" size="2"></font><font face="Arial, Helvetica, sans-serif" size="2"> 
        </font></td>
      <td width="36%"> <font color="#000000" face="Arial, Helvetica, sans-serif" size="2"> 
        </font><font face="Arial, Helvetica, sans-serif" size="2"></font><font face="Arial, Helvetica, sans-serif" size="2"> 
        <input type="text" name="loc2" value="<%=rst1("location")%>">
        </font></td>
      <td width="43%"> 
        <div align="left"><font color="#000000" face="Arial, Helvetica, sans-serif" size="2">Online: 
          <input type="checkbox" name="online2" value="1" <%if rst1("online") then %>checked<%end if%> >
          </font><font face="Arial, Helvetica, sans-serif" size="2"></font></div>
      </td>
    </tr>
  </table>
   
	    
  <table width="100%" border="0">
    <tr> 
      <td width="36%"><font face="Arial, Helvetica, sans-serif" size="2">Meter 
        Comment</font></td>
      
    </tr>
    <tr> 
      <td width="36%"> 
        <textarea name="textarea" rows="5" cols="30" wrap="PHYSICAL" ><%=rst1("metercomments")%></textarea>
      </td>
      
    </tr>
  </table>
  <table width="8%" border="0">
    <tr>
    <td>
        <input type="submit" name="upd" value="UPDATE">
      </td>
  </tr>
</table>

 
</form>
<%
else%>
<form name="savedetail" method="post" action=" savemeter.asp">     
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC" height="30"> 
        <div align="left"><font face="Arial, Helvetica, sans-serif"><font size="3" color="#FFFFFF"><i><b>New 
          Meter, Utility:
          <% =Request.querystring("utility")%>
          </b></i></font><font face="Arial, Helvetica, sans-serif" color="#000000" size="2">
          <input type="hidden" name="utility" value="<% =Request.querystring("utility")%>">
          </font></font></div>
    </td></tr>  
</table>

 
  <table width="100%" border="0">
    <tr> 
      <td width="6%" height="30"> <font size="2"> 
        <input type="hidden" name="bldg" value=" <%=Request.querystring("bldg")%>">
        <input type="hidden" name="acct"value=" <%=Request.querystring("acctid")%>">
        <font face="Arial, Helvetica, sans-serif" color="#000000">Meter:</font></font></td>
      <td width="12%" height="30"> <font size="2"> 
        <input type="text" name="meter" >
        </font></td>
      <td width="11%" height="30"> <font face="Arial, Helvetica, sans-serif" color="#000000" size="2">Start 
        Date:</font> </td>
      <td width="10%" height="30"><font face="Arial, Helvetica, sans-serif" color="#000000" size="2"> 
        <input type="text" name="sd" size="11" maxlength="10" >
        </font></td>
      <td width="10%" height="30"><font face="Arial, Helvetica, sans-serif" size="2">Online:</font> 
      </td>
      <td width="11%" height="30"><font face="Arial, Helvetica, sans-serif" color="#000000"> 
        <input type="checkbox" name="online" value="1" >
        </font><font face="Arial, Helvetica, sans-serif" color="#000000" size="2"> 
        </font> </td>
     
    </tr>
    <tr> 
      <td width="6%"><font face="Arial, Helvetica, sans-serif" color="#000000" size="2">Location:</font></td>
      <td width="12%"> <font size="2"> 
        <input type="text" name="loc" >
        </font></td>
      <td width="11%"><font face="Arial, Helvetica, sans-serif" color="#000000" size="2">Date 
        Off:</font> </td>
      <td width="10%"><font face="Arial, Helvetica, sans-serif" color="#000000" size="2"> 
        <input type="text" name="dof" size="11" maxlength="10" >
        </font></td>
      <td width="10%"><font face="Arial, Helvetica, sans-serif" color="#000000" size="2"> 
        </font></td>
      <td width="11%"><font size="2"></font></td>
    </tr>
  </table>   
	    <table width="100%" border="0">
          <tr>
      <td width="37%"><font face="Arial, Helvetica, sans-serif" size="2">Meter 
        Comment</font></td>
	 
	  </tr>
	  
	  <tr>
      <td width="37%"> 
        <textarea name="description" rows="5" cols="30" ></textarea>
      </td>
     
	  </tr>
	  </table>
	  <table width="8%" border="0">
    <tr>
    <td>
        <input type="submit" name="save" value="SAVE">
      </td>
  </tr>
</table>
  <%
end if

set cnn1=nothing
%>
</form>
</body>
</html>