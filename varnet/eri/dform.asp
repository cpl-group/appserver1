<% response.expires = 0 %> 
<html> 
<body bgcolor="#FFFFFF">
<% 
if request("submitter") = "submit" then 
' has the submit button has been pressed? 
' spit back the form elements entered 
response.write "here's what you entered:<br>" 
num = request.form("elements").count 
' how many elements 
' have already been entered? 
for i = 1 to num 
response.write request.form("elements")(i) 
response.write "<br>" 
next 
response.write request("element") 
' and finally the one 
' just entered, if present 
response.write "<br>" 

else


' create the form 
%>
<table border="1" align="left" cellspacing="0" cellpadding="0">
<form action=dform.asp method=post>
<br> 
<% 
num = request.form("elements").count 
for i = 1 to num 
' spit back the form elements entered 
%>
<td>
<% response.write "<input type=text name=elements value=""" 
response.write trim(request.form("elements")(i)) & """>" 
response.write "<br>" 
%>
</td>
<%
next 
if request.form("element") <> "" then 
' new input box...  %>
<td>
<%
response.write "<input type=text name=elements value=""" 
response.write trim(request.form("element")) & """>" 
response.write "<br>" 
%>
</td>
<%
end if 
%>
<td>
<%
response.write "enter item to add to list and press 'add'<br>" 
response.write "<input type=text size=50 name=element value="""""">" 
response.write "<br>" 

%>
<p> 
<input type="submit" name="adder" value="add"> 
<input type="submit" name="submitter" value="submit"> 
</td>
</table>
<br> 
</form>

<% end if %>

</body> 
</html>


