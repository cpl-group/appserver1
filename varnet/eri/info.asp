<%@Language="VBScript"%>
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'			Response.Redirect "http://www.genergyonline.com"
		end if	
		
		If Session("eri")  >  2 then
			
%>
<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<title>New Page 1</title>

<meta name="Microsoft Theme" content="none, default">
</head>

<%
Tenant = Request("qcatnr")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"


strsql = "SELECT * FROM tenant_history WHERE (tenant_no='" & Tenant & "') "  
 
rst1.Open strsql, cnn1, adOpenStatic
%>
<body bgcolor="#FFFFFF">
<div align="center">
 <center>
    <table border="1" width="100%">
      <tr> 
        <td width="42%" align="center" bgcolor="#66CCFF"><font face="Arial, Helvetica, sans-serif" color="#000000">Tenant 
          #</font></td>
        <td width="4%" align="center" bgcolor="#66CCFF"><font face="Arial, Helvetica, sans-serif" color="#000000">Date 
          Event</font></td>
        <td width="6%" align="center" bgcolor="#66CCFF">Sur_KWH</td>
        <td width="2%" align="center" bgcolor="#66CCFF">Sur_KW</td>
        <td width="2%" align="center" bgcolor="#66CCFF">SQFT</td>
        <td width="2%" align="center" bgcolor="#66CCFF"><font face="Arial, Helvetica, sans-serif" color="#000000">% 
          Rate</font></td>
        <td width="4%" align="center" bgcolor="#66CCFF"><font face="Arial, Helvetica, sans-serif" color="#000000">% 
          MAC</font></td>
        <td width="6%" align="center" bgcolor="#66CCFF"><font face="Arial, Helvetica, sans-serif" color="#000000">Charge</font></td>
        <td width="6%" align="center" bgcolor="#66CCFF"><font face="Arial, Helvetica, sans-serif" color="#000000">Note</font></td>
      </tr>
      <tr valign="middle"> 
        <%
Do While Not rst1.EOF
 %>
        <td width="16" align="center" height="6"><form name="form1" method="post" action="infoupdate.asp"> 
          <font face="Arial"> <i> 
          <input type="text" name="tenant_no" value="<%=rst1("Tenant_no")%>" size="15">
          </i></font></td>
        <td width="4%" align="center" height="6"><font face="Arial"><i> 
          <input type="text" name="date_event" value="<%=rst1("date_event")%>" size="6">
          </i></font></td>
        <td width="6%" align="center" height="6"><font face="Arial"><i> 
		<% if isnull(rst1("sur_kwh")) then %>
          <input type="text" name="sur_kwh" value="0" size="6">
		  <%else %>
          <input type="text" name="sur_kwh" value="<%=rst1("sur_kwh")%>" size="6">		  
		  <%end if %>
          </i></font></td>
        <td width="2%" align="center" height="6"><font face="Arial"><i> 
		<% if isnull(rst1("sur_kw")) then %>
          <input type="text" name="sur_kw" value="0" size="6">
		  <% else %>
          <input type="text" name="sur_kw" value="<%=rst1("sur_kw")%>" size="6">		  
		  <% end if %>
          </i></font></td>
        <td width="2%" align="center" height="6"><font face="Arial"><i>
		<% if isnull(rst1("sqft")) then %>
          <input type="text" name="sqft" value="<%=rst1("sqft")%>" size="6">
		  <% else %>
          <input type="text" name="sqft" value="<%=rst1("sqft")%>" size="6">		  
		  <% end if %>
          </i></font></td>
        <%if IsNull(rst1("rate")) Then %>
        <td width="2%" align="center" height="6"><i> 
          <input type="text" name="rate" value="<% =FormatPercent(0,2) %>" size="6">
          </i></font></td>
        <%else %>
        <td width="4%" align="center" height="6"><font face="Arial"><i> 
          <input type="text" name="rate" value="<%=FormatPercent(rst1("rate"))%>" size="6">
          </i></font></td>
        <% end if %>
        <%if IsNull(rst1("fuel")) Then %>
        <td width="6%" align="center" height="6"><i> 
          <input type="text" name="fuel" value="<% =FormatPercent(0,2) %>" size="10">
          </i></font></td>
        <%else %>
        <td width="6%" align="center" height="6"><font face="Arial"><i> 
          <input type="text" name="fuel" value="<%=FormatPercent(rst1("fuel"),2)%>" size="10">
          </i></font></td>
        <% end if %>
        <%if IsNull(rst1("Charge")) Then %>
        <td width="6%" align="center" height="6"><i> 
          <input type="text" name="charge" value="<%=FormatCurrency(0,2)%>" size="10">
          </i></font></td>
        <%else %>
        <td width="6%" align="center" height="6"><font face="Arial"><i> 
          <input type="text" name="charge" value=" <%=FormatCurrency(rst1("charge"),2)%>"  size="10">
          </i></font></td>
        <% end if 
		%>
        <td width="12%" align="center" height="6"><font face="Arial"><i> 
          <input type="text" name="note" value="<%=rst1("note")%>">
          </i></font></td>
        <td width="6%" align="center" height="6"><font face="Arial"> 
          <input type="hidden" name="id" value="<%=rst1("id") %>">
          <input type="submit" name="Submit" value="Update">
          </font> </form></td>
      </tr>
      <%
rst1.MoveNext  
Loop

rst1.Close
Set rst1 = Nothing
cnn1.Close
Set cnn1 = Nothing
 %>
 
 	    <td width="16" align="center" height="6"><form name="form1" method="post" action="infoadd.asp"> 
          <font face="Arial"> <i> 
          <input type="text" name="tenant_no" value="<%=Tenant%>" size="15">
          </i></font></td>
        <td width="4%" align="center" height="6"><font face="Arial"><i> 
          <input type="text" name="date_event" value="" size="6">
          </i></font></td>
        <td width="6%" align="center" height="6"><font face="Arial"><i> 
		    <input type="text" name="sur_kwh" value="0" size="6">
		  </i></font></td>
        <td width="2%" align="center" height="6"><font face="Arial"><i> 
		   <input type="text" name="sur_kw" value="0" size="6">
	      </i></font></td>
        <td width="2%" align="center" height="6"><font face="Arial"><i>
		  <input type="text" name="sqft" value="0" size="6">
		 </i></font></td>
        <td width="2%" align="center" height="6"><i> 
          <input type="text" name="rate" value="0" size="6">
          </i></font></td>
         <td width="6%" align="center" height="6"><i> 
          <input type="text" name="fuel" value="0" size="10">
          </i></font></td>
        <td width="6%" align="center" height="6"><i> 
          <input type="text" name="charge" value="0" size="10">
          </i></font></td>
        <td width="12%" align="center" height="6"><font face="Arial"><i> 
          <input type="text" name="note" value="">
          </i></font></td>
        <td width="6%" align="center" height="6"><font face="Arial"> 
          <input type="submit" name="Submit" value="Add">
          </font> </form></td>
      </tr>
    </table>
 </center>
</div>


</body>

</html>
<% else %>
 <html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">

<title>New Page 1</title>

<meta name="Microsoft Theme" content="none, default">
</head>

<%
Tenant = Request("qcatnr")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"


strsql = "SELECT * FROM tenant_history WHERE (tenant_no='" & Tenant & "') "  
 
rst1.Open strsql, cnn1, adOpenStatic
%>
<body bgcolor="#FFFFFF">
<div align="center">
 <center>
    <table border="1" width="100%">
      <tr> 
        <td width="7%" align="center" bgcolor="#66CCFF"><font face="Arial, Helvetica, sans-serif" color="#000000">Tenant 
          #</font></td>
        <td width="2%" align="center" bgcolor="#66CCFF"><font face="Arial, Helvetica, sans-serif" color="#000000">Date 
          Event</font></td>
        <td width="4%" align="center" bgcolor="#66CCFF"><font face="Arial, Helvetica, sans-serif" color="#000000">% 
          Rate</font></td>
        <td width="4%" align="center" bgcolor="#66CCFF"><font face="Arial, Helvetica, sans-serif" color="#000000">% 
          MAC</font></td>
        <td width="5%" align="center" bgcolor="#66CCFF"><font face="Arial, Helvetica, sans-serif" color="#000000">Charge</font></td>
        <td width="28%" align="center" bgcolor="#66CCFF"><font face="Arial, Helvetica, sans-serif" color="#000000">Note</font></td>
      </tr>
      <tr> 
        <%
Do While Not rst1.EOF
 %>
        <td width="7%" align="center"><font face="Arial"><i><%=rst1("Tenant_no")%></i></font></td>
        <td width="2%" align="center"><font face="Arial"><i><%=rst1("date_event")%></i></font></td>
        <%if IsNull(rst1("rate")) Then %>
        <td width="4%" align="center">&nbsp;</td>
        <%else %>
        <td width="4%" align="center"><font face="Arial"><i><%=FormatPercent(rst1("rate"),2)%></i></font></td>
        <% end if %>
        <%if IsNull(rst1("fuel")) Then %>
        <td width="5%" align="center">&nbsp;</td>
        <%else %>
        <td width="28%" align="center"><font face="Arial"><i><%=FormatPercent(rst1("fuel"),2)%></i></font></td>
        <% end if %>
        <%if IsNull(rst1("Charge")) Then %>
        <td width="4%" align="center">&nbsp;</td>
        <%else %>
        <td width="4%" align="center"><font face="Arial"><i><%=Formatcurrency(rst1("charge"),2)%></i></font></td>
        <% end if %>
        <td width="42%" align="center"><font face="Arial"><i><%=rst1("note")%></i></font></td>
      </tr>
      <%
rst1.MoveNext  
Loop

rst1.Close
Set rst1 = Nothing
cnn1.Close
Set cnn1 = Nothing
 %>
    </table>
 </center>
</div>


</body>

</html>

<% end if  %>