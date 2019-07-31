<%@Language="VBScript"%>
<!-- #include file="./adovbs.inc" -->
<%
		if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'			Response.Redirect "http://www.genergyonline.com"
		end if
		
if Session("eri") > 2 then 	
		
%>
<html>
<head>
<script language="JavaScript">

function openpopup(type, count){
	//var str2=document.doublecombo.qty.value
	var temp="filter.asp?type1="+type+"&count="+count
	//alert(temp);
    window.open(temp,"", "scrollbars=yes, width=300, height=338" );

}
	
function Desc_alert(){
	alert("Please enter a Description")
}
	
</script>





<title>Survey Item Details</title>

<meta http-equiv="Content-type1" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="FFFFFF" text="#000000" topmargin="0">
<table width="1031" border="1" cellspacing="2" cellpadding="0">
  <tr> 
    <td width="29" height="39" align="center"><font face="Arial" size="1">Qty</font></td>
    <td width="146" height="39" align="center"><font face="Arial" size="1">Type</font></td>
    <td width="148" height="39" align="center"><font face="Arial" size="1">Description</font></td>
    <td width="50" height="39" align="center"><font face="Arial" size="1">Amps</font></td>
    <td width="29" height="39" align="center"><font face="Arial" size="1">Volt</font></td>
    <td width="22" height="39" align="center"><font face="Arial" size="1">PH</font></td>
    <td width="39" height="39" align="center"><font face="Arial" size="1">PF</font></td>
    <td width="32" height="39" align="center">
      <p align="center"><font face="Arial" size="1">MF</font></p>
    </td>
    <td width="50" height="39" align="center"><font face="Arial" size="1">Watt</font></td>
    <td width="44" height="39" align="center"><font face="Arial" size="1">TotKw</font></td>
    <td width="36" height="39" align="center"><font face="Arial" size="1">Adj %</font></td>
    <td width="23" height="39" align="center"> 
      <div align="right"><font face="Arial" size="1">Adj Kw</font></div>
    </td>
    <td width="36" height="39" align="center"><font face="Arial" size="1">HOn</font></td>
    <td width="36" height="39" align="center"><font face="Arial" size="1">HOff</font></td>
    <td width="49" height="39" align="center"><font face="Arial" size="1">KwhOn</font></td>
    <td width="50" height="39" align="center"><font face="Arial" size="1">KwhOff</font></td>
    <td width="16" height="39" align="center"> 
      <div align="center"><font face="Arial" size="1">Int</font></div>
    </td>
    <td width="30" height="39" align="center"> 
      <div align="center"><font face="Arial" size="1">Base</font> </div>
    </td>
    <td width="18" align="center">&nbsp; </td>
  </tr>
  <%
Set cnn1 = Server.CreateObject("ADODB.Connection")

cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=eri_data;"

sql = "SELECT * FROM tblSurveyItem WHERE(surveyid='"& request("survey_id") & "') order by id"


Set rst1 = Server.CreateObject("ADODB.Recordset")

rst1.Open sql, cnn1, adOpenStatic, adLockReadOnly


count = 0 
If not rst1.EOF then 

%>
  <% 
    
    Do until rst1.EOF 
	type1=rst1("type")
	name="forms"&"["&count&"]"
	%>
  <tr> 
    <form name=form1 method="post" action="itemsupdate.asp">
      <input type="hidden" name="id" value="<%=Request("id")%>">
      <input type="hidden" name="orderno" value="<%=Request("orderno")%>">
      <input type="hidden" name="tenant_no" value="<%=Request("tenant_no")%>">
      <input type="hidden" name="location" value="<%=Request("location")%>">
      <input type="hidden" name="key" value="<%=rst1("id")%>">
      <td width=29> 
        <font face="Arial" size="1"> 
        <input type="text" name="qty" value="<%=rst1("qty") %>" size="3" maxlength="3" >
        </font>
      </td>
      <td width=146> 
        <font face="Arial" size="1"> 
        <select name="type1" size="1" >
          <option value="<%=rst1("type")%>" selected><%=rst1("type")%></option>
          <option value="Lighting">Lighting</option>
          <option value="Equipment">Equipment</option>
          <option value="HVAC">HVAC</option>
        </select>
        </font>
      </td>
      <td width="148"> <font face="Arial" size="1"> 
        <input type="text" name="description2" size="" value="<%=rst1("description")%>">
        </font> </td>
      <td width="50" align="right"> 
        <div align="right"> 
          <font face="Arial" size="1"> 
          <input type="text" name="amps" size="6" value="<%=formatnumber(rst1("amps"),2)%>" maxlength="6">
          </font>
        </div>
      </td>
      <td width="29" align="right"> 
        <div align="right"> 
          <p align="right"><font face="Arial" size="1"> 
          <input type="text" name="volt" size="3" value="<%=rst1("volt")%>"  maxlength="3">
          </font>
        </div>
      </td>
      <td width="22" align="right"> 
        <div align="right"> 
          <p align="right"><font face="Arial" size="1"> 
          <input type="text" name="ph" size="2" value="<%=rst1("ph")%>" maxlength="2">
          </font>
        </div>
      </td>
      <td width="39" align="right"> 
        <div align="right"> 
          <p align="right"><font face="Arial" size="1"> 
          <input type="text" name="pf" size="4" value="<%=formatnumber(rst1("pf"),2)%>" maxlength="4">
          </font>
        </div>
      </td>
      <td width="32" align="right"> 
        <div align="right"> 
          <p align="right"><font face="Arial" size="1"> 
          <input type="text" name="mf" size="3" value="<%=rst1("monthfactor")%>" maxlength="3">
          </font>
        </div>
      </td>
      <td width="50" align="right"> 
        <div align="right"> 
          <p align="right"><font face="Arial" size="1"> 
          <input type="text" name="watt" size="6" value="<%=rst1("watt")%>" maxlength="10" >
          </font>
        </div>
      </td>
      <td width="44" align="right">  
        <div align="right"><font face="Arial" size="1"><%=formatnumber(rst1("totalkw"),2)%></font></div>
      </td>
      <td width="36" align="right"> 
        <div align="right"> 
          <font face="Arial" size="1"> 
          <input type="text" name="adj" value="<%=formatnumber(rst1("adjfactor"),2)%>" size="4" maxlength="4" >
          </font>
        </div>
      </td>
      <td width="23" align="right"> 
        <div align="right"><font face="Arial" size="1"><%=formatnumber(rst1("adjkw"),2)%> </font> </div>
      </td>
      <td width="36" align="right"> 
        <div align="right"> 
          <font face="Arial" size="1"> 
          <input type="text" name="hon" value="<%=rst1("houron")%>" size="4" maxlength="4" >
          </font>
        </div>
      </td>
      <td width="36" align="right"> 
        <div align="right"> 
          <font face="Arial" size="1"> 
          <input type="text" name="hoff" value="<%=rst1("houroff")%>" size="4" maxlength="4" >
          </font>
        </div>
      </td>
      <td width="49" align="right">  
        <div align="right"><font face="Arial" size="1"><%=formatnumber(rst1("kwhon"),2)%></font></div>
      </td>
      <td width="50" align="right">  
        <div align="right"><font face="Arial" size="1"><%=formatnumber(rst1("kwhoff"),2)%></font></div>
      </td>
      <td width="16"> <font face="Arial" size="1"> 
        <% 
			if(rst1("intense")=true) then
			%>
        <input type="checkbox" name="ie" size="6" checked>
          <%
			else
			%>
          <input type="checkbox" name="ie" size="6" >
          <%
			end if
			
			%>
          </font>
        </td>
        <td width="18"> 
          <font face="Arial" size="1"> 
          <%
			if(rst1("base")=true) then
			%>
          <input type="checkbox" name="bho" size="6" checked>
          <%
			else
			%>
          <input type="checkbox" name="bho" size="6">
          <%
			end if
			%>
          </font>
        </td>
        <td width="68"> 
          <font face="Arial" size="1"> 
          <input type="Submit" name="Submit" value="Update">
          </font>
        </td>
    </form>
  </tr>
  <%
    count = count +1
    rst1.movenext
    Loop
	rst1.close
    %>
  <%
end if
	
%>
  <tr> 
    <form name="form2" method="post" action="itemsadd.asp">
      <input type="hidden" name="id" value="<%=Request("id")%>">
      <input type="hidden" name="key" value="<%=id%>">
      <td width=29> 
	  <a tabindex=0>
        <font face="Arial" size="1">
        <input type="text" name="qty" value="0" size="3"  maxlength="3">
        </font>
		</a>
      </td>
      <td width=146> 
        <font face="Arial" size="1"> 
        <select name="type1" size="1" >
          <option value="Lighting">Lighting</option>
          <option value="Equipment">Equipment</option>
          <option value="HVAC">HVAC</option>
          <option value="" selected>========</option>
        </select>
        </font>
      </td>
      <td width=148> 
        <font face="Arial" size="1"> 
        <input type="text" name="description" size="20" value="Desc" >
        </font>
      </td>
      <td width="50" align="right"> 
        <div align="right"> 
          <font face="Arial" size="1"> 
          <input type="text" name="amps" size="6" value="0" maxlength="6">
          </font>
        </div>
      </td>
      <td width="29" align="right"> 
        <div align="right"> 
          <font face="Arial" size="1"> 
          <input type="text" name="volt" size="3" value="115" maxlength="3">
          </font>
        </div>
      </td>
      <td width="22" align="right"> 
        <div align="right"> 
          <font face="Arial" size="1"> 
          <input type="text" name="ph" size="2" value="1" maxlength="2">
          </font>
        </div>
      </td>
      <td width="39" align="right"> 
        <div align="right"> 
          <font face="Arial" size="1"> 
          <input type="text" name="pf" size="4" value="0.85" maxlength="4">
          </font>
        </div>
      </td>
      <td width="32" align="right"> 
        <div align="right"> 
          <font face="Arial" size="1"> 
          <input type="text" name="mf" size="3" value="12" maxlength="3">
          </font>
        </div>
      </td>
      <td width="50" align="right"> 
        <div align="right"> 
          <font face="Arial" size="1"> 
          <input type="text" name="watt" size="6" value="0" maxlength="10" >
          </font>
        </div>
      </td>
      <td width="44" align="right"> </td>
      <td width="36" align="right"> 
        <div align="right"> 
          <font face="Arial" size="1"> 
          <input type="text" name="adj" value="1" size="4" maxlength="4" >
          </font>
        </div>
      </td>
      <td width="23" align="right"> 
        <div align="right"></div>
      </td>
      <td width="36" align="right"> 
        <div align="right"> 
          <font face="Arial" size="1"> 
          <input type="text" name="hon" size="4" value="0" maxlength="4" >
          </font>
        </div>
      </td>
      <td width="36" align="right"> 
        <div align="right"> 
          <font face="Arial" size="1"> 
          <input type="text" name="hoff" size="4" value="0" maxlength="4" >
          </font>
        </div>
      </td>
      <td width="49" align="right"> </td>
      <td width="50" align="right"> </td>
      <td width="16"> <font face="Arial" size="1"> 
        <input type="checkbox" name="ie2" size="6">
          </font>
        </td>
        <td width="18"> 
          <font face="Arial" size="1"> 
          <input type="checkbox" name="bho" size="6">
          </font>
        </td>
        
      <td width="68"> <font face="Arial" size="1"> 
        <input type="Submit" name="Submit2" value="Add">
        </font> </td>
    </form>
  </tr>
  <%

end if
'cnn1.close
set cnn1=nothing
%>
</table>
</body>
</html>
