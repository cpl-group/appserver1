<html>
<head>
<title>Load Profile</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<style type=3D"text/css"><!--A {text-decoration: none}--></style>

<%
m=Request.QueryString("m")
d=Request.QueryString("d")
b=Request.QueryString("b")
s=Request.QueryString("s")
e=Request.QueryString("e")
luid = Request.QueryString("luid")
lmp=Request.QueryString("lmp")
if isempty(d) then 
	if luid = "" then 
		Set cnn1 = Server.CreateObject("ADODB.Connection")
		Set rst1 = Server.CreateObject("ADODB.recordset")
		cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
		strsql = "select left(convert(varchar,max(date),101),11) as date from pulse_"&b&" where meterid = " & m
		rst1.Open strsql, cnn1
		if not rst1.eof then 
			d=rst1("date")		
		end if
		rst1.close
	else
		Set cnn1 = Server.CreateObject("ADODB.Connection")
		Set rst1 = Server.CreateObject("ADODB.recordset")
		cnn1.Open "driver={SQL Server};server=10.0.7.16;uid=genergy1;pwd=g1appg1;database=genergy1;"
		strsql = "select left(convert(varchar,max(date),101),11) as date from pulse_"&b&" where meterid in (select meterid from meters where leaseutilityid = "& luid &")" 
		rst1.Open strsql, cnn1
		if not rst1.eof then 
			d=rst1("date")		
		end if
		rst1.close
	end if
end if
if isempty(s) then
	s=15
end if
if isempty(e) then 
	e=2400
end if 
%>
<script>
function chgdate(){

	var b = document.forms[0].b.value
	var m = document.forms[0].m.value
	var d = document.forms[0].d.value
	var s = document.forms[0].s.value
	var e = document.forms[0].e.value
	var l = document.forms[0].luid.value
	var lmp= document.forms[0].lmp.value
	var temp = "calendar.asp?b=" + b + "&m=" + m + "&d=" + d + "&s="+s+"&e="+e+"&luid="+l+"&lmp=" + lmp
	//window.open(temp,"","statusbar=0,menubar=0,scrollbars=yes,HEIGHT=300,WIDTH=300")
	calendar = window.open(temp,'cal','WIDTH=200,HEIGHT=250');	
}
function zoomentry(){

	var b = document.forms[0].b.value
	var m = document.forms[0].m.value
	var d = document.forms[0].d.value
	var s = document.forms[0].s.value
	var e = document.forms[0].e.value
	var l = document.forms[0].luid.value
	var lmp = document.forms[0].lmp.value
	var temp = "zoomentry.asp?b=" + b + "&m=" + m + "&d=" + d + "&s="+s+"&e="+e+"&luid="+l+"&lmp="+lmp
	window.open(temp,"","statusbar=0,menubar=0,scrollbars=yes,HEIGHT=125,WIDTH=300")
}
function lmpmoveprev(){

	var b = document.forms[0].b.value
	var m = document.forms[0].m.value
	var d = document.forms[0].pd.value
	var nd = document.forms[0].d.value
	var s = document.forms[0].s.value
	var e = document.forms[0].e.value
	var l = document.forms[0].luid.value
	var lmp = document.forms[0].lmp.value
	var pd = new Date(d)
	pd.setTime(pd.getTime() - 1 * 24 * 60 * 60 * 1000)
	pd = (pd.getMonth()+1) + "/" + pd.getDate() + "/" + pd.getYear()
		
	document.forms[0].d.value = d
	document.forms[0].pd.value = pd
	document.forms[0].nd.value = nd
	
	var temp="lmpload.asp?m="+m+"&d="+d+"&b="+b+"&s="+s+"&e="+e+"&luid="+l+"&lmp="+lmp
	document.frames.lmp.location=temp;

	
	}
function lmpmovenext(){

	var b = document.forms[0].b.value
	var m = document.forms[0].m.value
	var d = document.forms[0].nd.value
	var pd = document.forms[0].d.value
	var s = document.forms[0].s.value
	var e = document.forms[0].e.value
	var l = document.forms[0].luid.value
	var lmp = document.forms[0].lmp.value
	var nd = new Date(d)
	nd.setTime(nd.getTime() + 1 * 24 * 60 * 60 * 1000)
	nd = (nd.getMonth()+1) + "/" + nd.getDate() + "/" + nd.getYear()
		
	document.forms[0].d.value = d
	document.forms[0].pd.value = pd
	document.forms[0].nd.value = nd
	
	var temp="lmpload.asp?m="+m+"&d="+d+"&b="+b+"&s="+s+"&e="+e+"&luid="+l+"&lmp="+lmp
	document.frames.lmp.location=temp;

	
	}
function lmpnow(){

	var b = document.forms[0].b.value
	var m = document.forms[0].m.value
	var d = document.forms[0].td.value
	var s = document.forms[0].s.value
	var e = document.forms[0].e.value
	var l = document.forms[0].luid.value
	var lmp = document.forms[0].lmp.value
	var nd = new Date(d)
	nd.setTime(nd.getTime() + 1 * 24 * 60 * 60 * 1000)
	nd = (nd.getMonth()+1) + "/" + nd.getDate() + "/" + nd.getYear()
	var pd = new Date(d)
	pd.setTime(pd.getTime() - 1 * 24 * 60 * 60 * 1000)
	pd = (pd.getMonth()+1) + "/" + pd.getDate() + "/" + pd.getYear()

	document.forms[0].d.value = d
	document.forms[0].pd.value = pd
	document.forms[0].nd.value = nd

	var temp="lmpload.asp?m="+m+"&d="+d+"&b="+b+"&s="+s+"&e="+e+"&luid="+l+"&lmp="+lmp
	document.frames.lmp.location=temp;
	}
</script>
<body bgcolor="#FFFFFF" text="#000000" link="#FFFFFF" vlink="#FFFFFF" alink="#0099FF">
<h6></h6>
<table width="784" border="1" cellspacing="0" cellpadding="0" height="277" align="center">
  <form name="form1" method="post" action="">
    <tr> 
      <td width="687" height="4" bgcolor="#000000"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#000000">
          <tr> 
            <td width="90%" height="2" bgcolor="#000000"> <font size="2"><b><font face="Arial, Helvetica, sans-serif" color="#FFFFFF"> 
              <input type="hidden" name="b" value="<%=b%>">
              <input type="hidden" name="m" value="<%=m%>">
              <input type="hidden" name="s" value="<%=s%>">
              <input type="hidden" name="e" value="<%=e%>">
              <input type="hidden" name="d" value="<%=d%>">
              <input type="hidden" name="pd" value=<%=DateAdd("d",-1,d)%>>
              <input type="hidden" name="nd" value=<%=DateAdd("d",1,d)%>>
              <input type="hidden" name="td" value=<%=Date()%>>
              <input type="hidden" name="luid" value=<%=luid%>>
              <input type="hidden" name="lmp" value=<%=lmp%>>
<!--               <a href="javascript:chgdate()" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'"> 
              Select Date</a> | <a href="javascript:zoomentry()" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Interval 
              Zoom</a> |  --><a href="javascript:lmpmoveprev()" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Previous</a> 
              | <a href="javascript:lmpmovenext()" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Next</a> 
              | <a href="javascript:lmpnow()" style="text-decoration:none;" onMouseOver="this.style.color = 'lightblue'" onMouseOut="this.style.color = 'white'">Go 
              To Today</a></font></b></font></td>
          </tr>
        </table>
      </td>
    </tr>
  </form>
  <tr> 
    <td width="687" height="330"><iframe name="lmp" width="100%" height="100%" src=<%="lmpload.asp?m="&m&"&d="&d&"&b="&b&"&s="&s&"&e="&e&"&luid="&luid&"&lmp="&lmp%> scrolling="auto" marginwidth="0" marginheight="0" ></iframe> 
    </td>
  </tr>
</table>
<table width="784" border="1" cellspacing="0" cellpadding="0" align="center" height="312">
  <tr>
    <td><iframe name="dataset" width="100%" height="100%" src=<%="options.asp?m="&m&"&b="&b&"&luid="&luid%> scrolling="auto" marginwidth="0" marginheight="0" ></iframe></td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
