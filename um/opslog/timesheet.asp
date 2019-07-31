<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if isempty(Session("name")) then
%>
<script>
top.location="../index.asp"
</script>
<%
			'	Response.Redirect "http://www.genergyonline.com"
else
	if Session("ts") < 4 then 
		Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
		Response.Redirect "../main.asp"
	end if	
end if	
user="ghnet\"&Session("login")
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getConnect(0,0,"intranet")

sql="select username, startweek, endweek from user_cost where username='"&user&"'"
rst1.Open sql, cnn1, 0, 1, 1

if not rst1.eof then
	startweek=rst1("startweek")
	endweek=rst1("endweek")
end if
rst1.close
strsql = "SELECT *, matricola AS Expr1 FROM Times WHERE (matricola = '"& user &"'  and [date] between '" & Startweek - 18  & "' and '" & endweek  &"') order by date desc"
rst1.Open strsql, cnn1, 0, 1, 1
'Response.Write(strsql)
%>
<html>
<head>
<script>
function openpopup(){
//configure "Open Logout Window
    parent.document.location.href="../index.asp";
}
function loadpopup(){
    openpopup()
}
//document.main.location="http://www.yahoo.com"
function updateEntry(id){
	parent.frames.bottom.location="timedetail.asp?id="+id
}
function displaytotal(hrs, ot, expn, lastdate){

	var temp = "Totals as of " + lastdate + " : Hours = " + hrs + ", Overtime Hours =  " + ot + ", Expenses = " + expn
	alert(temp) 

} 
function delete1(key,u){
	if(confirm("Time has been deleted")){
	document.location="deletetime.asp?key="+key+"&u="+u
	}
}
</script>
<body bgcolor="FFFFFF">
<input type="button" name="Submit3" value="Totals" onClick="displaytotal(form2.hrstotal.value,form2.bhrstotal.value,form2.expensetotal.value, form2.lastdate.value)">
<br>
<table width="100%" border="1" height="8" align="center" cellpadding="0" cellspacing="0" bordercolor="#999999">
  <%
if not rst1.eof then
timesheettotal_hrs = 0
timesheettotal_ot = 0
expensetotal=0
lastdate=rst1("date")
	Do until rst1.EOF 
%>
  <tr bgcolor="#CCCCCC" valign="middle"> 
    <form name=form1 method="post" action="">
       <input type="hidden" name="key" value="<%=rst1("id")%>">
	   <input type="hidden" name="u" value="<%=rst1("expr1")%>">
      <td width=5% height="34"> 
        <input type="button" name="edit" value="edit" size="5" onClick="updateEntry(key.value)">
      </td>
	  <td width=2% height="34"> <font face="Arial, Helvetica, sans-serif" size="1"><a href="javascript:delete1('<%=rst1("id")%>','<%=rst1("expr1")%>')"><img src="delete.gif" border="0"></a>
        </font> </td>
      <td width=7% height="34"> <font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("date")%> 
        </font></td width=17%>
      <td width="8%" height="34"> <font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("jobno")%> 
        </font></td>
      <td width="57%" height="34"> <font face="Arial, Helvetica, sans-serif" size="1"><%=rst1("description")%> 
        </font></td>
      <td width="3%" bgcolor="#00CCFF" height="34"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#000000"><%=rst1("hours")%> 
		<%  if rst1("date") >= startweek then
			timesheettotal_hrs=timesheettotal_hrs + Formatnumber(rst1("hours"))
			end if%>
          </font></div>
      </td>
      <td width="3%" bgcolor="#3399CC" height="34"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><%=rst1("hours_bill")%> 
          </font></div>
      </td>
      <td width="2%" bgcolor="#0033FF" height="34"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><%=rst1("overt")%> 
		<% if rst1("date") >= startweek then
			timesheettotal_ot=timesheettotal_ot + Formatnumber(rst1("overt")) 
			end if%>
          </font></div>
      </td>
      <td width="3%" bgcolor="#0066CC" height="34"> 
        <div align="center"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><%=rst1("expense")%> 
          </font></div>
      </td>
      <td width="10%" bgcolor="#3300CC" height="34"> 
        <div align="right"><font face="Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><%=FormatCurrency(rst1("value"))%> 
		<% if rst1("date") >= startweek then 
		  expensetotal=expensetotal + Formatnumber(rst1("value"))
		  end if %>
          </font></div>
      </td>
    </form>
  </tr>
  <%  
    rst1.movenext
    loop
end if
%>
</table>
<form name="form2" method="post" action="">

<input type="hidden" name="hrstotal" value="<%=timesheettotal_hrs%>">
<input type="hidden" name="bhrstotal" value="<%=timesheettotal_ot%>">
<input type="hidden" name="expensetotal" value="<%=Formatcurrency(expensetotal)%>">
<input type="hidden" name="lastdate" value="<%=lastdate%>">
</form>
</body>
</html>
