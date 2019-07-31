<html>
<head>
<%@Language="VBScript"%>
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
	
%>

<title>Operations Log</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script>
function fillup(bldgnum, items){
	document.location="capindex.asp?bldgnum=" + bldgnum +"&items="+ items
}
function searchcap(bldgnum, items, num) {
//alert(typebox)
//alert(searchitem)
	var temp
	temp="capsearch.asp?bldgnum="+bldgnum+"&items="+items+"&num="+num
	document.frames.capacity.location=temp
}
function report(spec, job) {
document.frames.oplog.location=spec
}

function timesheetjob(typebox, job){
//alert(typebox)
//alert(document.forms[0].findvar.value)
	var temp
	if(typebox =="[entry id]"){
	    if(job == ""){
		    alert("Please enter job number")
		}else{
            if(isNaN(job)){
				alert("Not a valid number")
            }else{
                temp="timesheetmain.asp?job="+job
//				temp="timesheetsearch.asp?job="+job
				document.frames.oplog.location=temp
			}
		}
	}else{
		temp="null.htm"
		document.frames.oplog.location=temp
	}
	
}

</script>
</head>
<%
bldgnum = Request("bldgnum")
items = Request("items")
'response.write(items)
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open "driver={SQL Server};server=10.0.7.20;uid=genergy1;pwd=g1appg1;database=Capacity_db;"

sql="select bldgnum, address from tlbldg"	
rst1.Open sql, cnn1, 0, 1, 1
if not rst1.eof then
%>
<body bgcolor="#FFFFFF" text="#000000">
<table width="100%" border="0">
  <tr> 
    <td bgcolor="#3399CC"> 
      <div align="center"><font color="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif">Capacity</font></b></font></div>
    </td>
  </tr>
</table>


<form name="form1">
<table width="100%" border="0" align="center">
  <tr> 
    <td align="left" height="36"> 
        <font face="Arial, Helvetica, sans-serif">
        Search for Building 
		</font>
		
        <select name="bldgnum" size="1" onChange="fillup(this.value, items.value)">
          <%
		  do until rst1.eof
		      if(bldgnum=Trim(rst1("bldgnum"))) then 
		  %>
          <option value="<%=Trim(bldgnum)%>" selected><font face="Arial, Helvetica, sans-serif"><%=rst1("address")%> 
          </font></option>
          <%
		      else
		  %>
          <option value="<%=Trim(rst1("bldgnum"))%>"><font face="Arial, Helvetica, sans-serif"><%=rst1("address")%></font> 
          <%
		      end if
		  rst1.movenext
		  loop
		  %>
        </select>
        <select name="items" onchange="fillup(bldgnum.value, this.value)">
		<%
		if not isempty(bldgnum) then
			if items="floor_name" then
		%>
        	  <option value="floor_name" selected>floor<font face="Arial, Helvetica, sans-serif"></font></option>
              <option value="riser_name">riser<font face="Arial, Helvetica, sans-serif"></font></option> 
		<%
			else
		%>
		  <option value="riser_name" selected>riser<font face="Arial, Helvetica, sans-serif"></font></option>
          <option value="floor_name">floor<font face="Arial, Helvetica, sans-serif"></font></option> 
		<%
			end if
		else
		%>
		  <option value="">======<font face="Arial, Helvetica, sans-serif"></font></option>
		  <option value="floor_name">floor<font face="Arial, Helvetica, sans-serif"></font></option>
          <option value="riser_name">riser<font face="Arial, Helvetica, sans-serif"></font></option>
		<%
		end if
		%>
        </select>
<%
	end if
	rst1.close
	if items <> "" then
		if items="floor_name" then
			sql="select distinct fl_name as num from tblassociation where bldgnum='"& bldgnum &"'"
		end if
		if items="riser_name" then
			sql="select distinct riser_name as num from tblassociation where bldgnum='"& bldgnum &"'"
		end if
	rst1.Open sql, cnn1, 0, 1, 1
	'response.write sql
	if not rst1.eof then
%>
<select name="num">
<%
    do until rst1.eof
%>
<option value="<%=rst1("num")%>"><%=rst1("num")%></option>
<%
	rst1.movenext
	loop

%>
</select>
        <input type="button" name="Submit2" value="Find" onClick='searchcap(bldgnum.value,items.value, num.value)'>
<%
	end if
end if
%>      
    </table>
</form>
<font face="Arial, Helvetica, sans-serif">
<input type="button" name="Submit3" value="New Building" onClick='javascript:capacity.location="capnewbldg.asp"'>
</font> 
<br><br>
<IFRAME name="capacity" width="100%" height="100%" src="null.htm" scrolling="auto" marginwidth="0" marginheight="0" ></IFRAME> 
</body>
</html>