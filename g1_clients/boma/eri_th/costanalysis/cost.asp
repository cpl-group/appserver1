<!-- #include file="../lmp/./adovbs.inc" -->
<% 
bldg= Request("bldg")
%> 
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function yearfill(bldg,pid){
    document.location="cost.asp?bldg="+bldg+"&pid="+pid
}
function createImg(cost){
    var bldg=document.forms[0].bldg.value
	var y1=document.forms[0].y1.value
	var y2=document.forms[0].y2.value
	var ip="sqlserverg1"
	var temp= "http://www.genergy.com/cgi-bin/billbar.cgi?bldg=" + bldg + "&y1=" + y1 +"&y2=" + y2+"&cost="+cost+"&ip="+ip
	document.frames.graph.location=temp;
}

</script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#0099FF" height="21"> 
      <div align="center"><font face="Arial, Helvetica, sans-serif" size="4" color="#FFFFFF">Cost 
        Analysis </font></div>
    </td>
  </tr>
</table>
  
<form method="post" action="" name="list">
<input type="hidden" name="ip" value="10.0.7.16">
  <table width="306" border="0" cellspacing="0" cellpadding="0" align="center" height="55">
    <tr valign="top"> 
      <td height="9" width="103"> 
        <div align="left"> 
          <p><font face="Arial, Helvetica, sans-serif" size="3">Building</font></p>
        </div>
      </td>
      <td height="9" width="47"> 
        <div align="left"><font face="Arial, Helvetica, sans-serif" size="3">Year 
          1</font></div>
      </td>
      <td height="9" width="45"> 
        <div align="left"><font face="Arial, Helvetica, sans-serif" size="3">Year 
          2</font></div>
      </td>
    </tr>
    <tr> 
      <td height="27" width="103"> 
        <%
		Set cnn1 = Server.CreateObject("ADODB.Connection")
		Set rst1 = Server.CreateObject("ADODB.recordset")
		cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
      %>
        <input type="hidden" name="profiletype" value="<%=profiletype %>">
        <input type="hidden" name="prev" value="prev">
        <input type="hidden" name="next" value="next">
        <select name="bldg" onChange="yearfill(this.value,pid.value)">
          <%
		strsql = "SELECT strt, bldgnum FROM buildings where portfolioid='" & request.querystring("pid") & "' order by strt"
    	rst1.Open strsql, cnn1, adOpenStatic
		if bldg <> "" then
    	    do until rst1.EOF 
			if(rst1("bldgnum") = bldg) then
	  %>
          <option value=<%=rst1("bldgnum")%> selected> <%=rst1("strt")%></option>
          <%
	   	    else
			%>
          <option value=<%=rst1("bldgnum")%>> <%=rst1("strt")%></option>
          <%
			end if
			rst1.movenext
			loop
		else

		    %>
          <option value=<%=request("pid")%> selected>Portfolio</option>
          <%
			do until rst1.EOF
			%>
          <option value=<%=rst1("bldgnum")%>> <%=rst1("strt")%></option>
          <%
		    rst1.movenext
		    loop
    		end if
			rst1.close
			%>
        </select>
      </td>
      <%
		  if bldg <> "" then
		%>
      <td height="27" width="47"> 
        <select name="y1">
          <%
			  strsql = "SELECT distinct billyear FROM billyrperiod WHERE (bldgnum = '" & bldg & "') order by billyear desc"
			  rst1.Open strsql, cnn1, adOpenStatic
			  if not rst1.EOF then		
			      do until rst1.EOF
		%>
          <option value=<%=rst1("billyear")%>><%=rst1("billyear") %> </option>
          <%
		          rst1.movenext
		          loop
			  rst1.movefirst	  
		      end if
		  %>
        </select>
      </td>
      <td width="45" height="27"> 
        <select name="y2">
          <%
     		  if not rst1.EOF then
	        	  do until rst1.EOF
		  %>
          <option value=<%=rst1("billyear")%>><%=rst1("billyear") %> </option>
          <%
		          rst1.movenext
		          loop
		          rst1.close
		      end if
		%>
        </select>
      </td>
      <%
	    else
	  %>
      <td height="27" width="50"> 
        <select name="y1">
          <option>------</option>
        </select>
      </td>
      <td height="27" width="50"> 
        <select name="y2">
          <option>------</option>
        </select>
      </td>
      <%
	    end if
	  %>
    </tr>
    <tr> 
      <td width="103" height="70"> 
        <p> 
          <input type="hidden" name="pid" value="<%=Request("pid")%>" >
        </p>
        <p align="left"> 
          <select name="sp_procedure">
            <%
			  strsql = "select name, substring(name, 7, 25) as Fieldname from sysobjects where left(name,2) ='g1' and (substring(name,4,2)='ge' or  substring(name, 4,3)='" & Request("pid") & "')"
			  rst1.Open strsql, cnn1, adOpenStatic
			  if not rst1.EOF then
			  do until rst1.eof
		  %>
            <option value="<%=rst1("name")%>"><%=rst1("fieldname") %> </option>
            <%
		  	  rst1.movenext
			  loop
			  end if
		  %>
          </select>
        </p>
      </td>
      <td height="70" width="47"> 
        <div align="center"> 
          <input type="button" name="Button" value="View" onClick="createImg(sp_procedure.value)">
        </div>
      </td>
      <td height="70" width="45">&nbsp; </td>
    </tr>
  </table>
</form>
<p align="left"><IFRAME name="graph" src="null.htm" width="100%" height="100%" scrolling="auto" marginwidth="8" marginheight="16"></IFRAME></p>
</body>
</html>
