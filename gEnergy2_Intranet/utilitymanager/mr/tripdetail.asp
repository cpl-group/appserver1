<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<html>
<head>
<%

Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_palmserver")

%>
<link rel="Stylesheet" href="/gEnergy2_Intranet/styles.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head>
<body bgcolor="#eeeeee" text="#000000" style="border-top:2px outset #ffffff;" class="innerbody">
<form name="form1" method="post" action="tripmodify.asp">

<%
key=cint(Request("key"))
sql2 = "SELECT * FROM Tripcodeindex WHERE autoid =" & key
Set rst2 = Server.CreateObject("ADODB.Recordset")

rst2.Open sql2, cnn1, adOpenStatic, adLockReadOnly
if not rst2.EOF then
        id 		= 	rst2("id")
		bldgnum = 	rst2("bldgnum")
        utility =	rst2("utilityid")
		tripdate=	rst2("tripdate")
		billyear=	rst2("billyear")
		billperiod=	rst2("billperiod")
end if
rst2.close
%>
<table width="100%" border=0 cellpadding="2" cellspacing="0" bgcolor="#eeeeee">
<tr>
      <td bgcolor="#dddddd" style="border-bottom:1px solid #999999;">&nbsp; </td>
</tr>
<tr>
  <td>
  <table border=0 cellpadding="0" cellspacing="2">
          <tr> 
            <td width="103">Trip Code</td>
            <td width="1">&nbsp;</td>
            <td width="210">Building </td>
            <td width="60">Utility</td>
            <td width="82">Trip Date</td>
            <td width="127">Bill Year</td>
            <td width="115">Bill Period</td>
            <td width="115" nowrap>Apply To All In Trip</td>
          </tr>
          <tr> 
            <td> <input type="hidden" name="key" value="<%=key%>" size="6"> <input type="text" name="tripid" value="<%=id%>" size="6"> 
            </td>
            <td>&nbsp;</td>
            <td> 
              <% 
		strsql = "SELECT  bldgnum, bldgname FROM BuildingIndex"
		rst2.Open strsql, cnn1, adOpenStatic
		if not rst2.eof then 
		%>
              <select name="bldgnum">
                <% 
		  while not rst2.eof
		  %>
                <option value="<%=rst2("bldgnum")%>" <%if rst2("bldgnum") = trim(bldgnum) then%> selected <%end if%>><%=rst2("bldgname")%>, <%=rst2("bldgnum")%></option>
                <% 
		  rst2.movenext
		  wend
		%>
              </select> 
              <%
		end if
		rst2.close
		%>
            </td>
            <td> 
              <% 
		strsql = "SELECT  * FROM UtilitiesIndex"
		rst2.Open strsql, cnn1, adOpenStatic
		if not rst2.eof then 
		%>
              <select name="utility">
                <% 
		  while not rst2.eof
		  %>
                <option value="<%=rst2("id")%>" <%if trim(rst2("id")) = trim(utility) then%> selected <%end if%>><%=rst2("utility")%></option>
                <% 
		  rst2.movenext
		  wend
		%>
              </select> 
              <%
		end if
		rst2.close
		%>
            </td>
            <td><input type="text" name="tripdate" size="10%" value="<%=tripdate%>"></td>
            <td><input type="text" name="billyear" size="10%" value="<%=billyear%>"></td>
            <td><input type="text" name="billperiod" size="10%" value="<%=billperiod%>"> 
            </td>
            <td><input name="applytoall" type="checkbox" value="1" checked></td>
          </tr>
          <tr valign="top"> 
            <%
		  if key > 0 then 
			modify = "Update"
		  else
		  	modify = "Save"	
		  end if
		  %>
            <td colspan="12"> <input type="Submit" name="modify" value="<%=modify%>" style="border:1px outset #ddffdd;background-color:ccf3cc;font-weight:bold;">
              <input type="Submit" name="modify_sub" value="Delete Trip" style="border:1px outset #ddffdd;background-color:ccf3cc;"> 
              <input type="button" name="Button" value="Cancel" onclick="<%=CancelOnclick%>" style="border:1px outset #ddffdd;background-color:ccf3cc;"> 
            </td>
          </tr>
        </table>
  </td>
</tr>
</table>
	</form>
</body>
</html>
