<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<html>
<head>

<%
Dim cnn1
Set cnn1 = Server.CreateObject("ADODB.connection")
Set rst1 = Server.CreateObject("ADODB.recordset")

cnn1.Open getconnect(0,0,"engineering")
rst1.ActiveConnection = cnn1
rst1.Cursortype = adOpenStatic
id=request.querystring("id")
bldg=request.querystring ("bldg")
fid=request.querystring ("fid")
rid=request.querystring ("rid")
'response.write id

sqlstr= "select top 1 r.room as roomname, fl.floor as floorname, ft.*,f.* ,l.*, DATEADD(week, (est_lamp_life/est_hr_wk), datelastchanged) as estd ,datediff(week,getdate(), DATEADD(week, (est_lamp_life/est_hr_wk), datelastchanged)) as weeksr, bdatelastchanged+(ballast_life/est_hr_wk) as bestd ,datediff(week,getdate(), (bdatelastchanged+(ballast_life/est_hr_wk))) as bweeksr,r.id as rid, fid from fixture_types ft join fixtures  f on ft.id=f.typeid join lamping_sch l on f.id=l.fid INNER JOIN room r ON r.id=f.room INNER JOIN floor fl ON r.floor=fl.id where l.fid='"& id &"'  order by l.datelastchanged desc,bdatelastchanged desc"
rst1.Open sqlstr, cnn1, 0, 1, 1

if not rst1.eof then
  room=rst1("roomname")
  fl=rst1("floorname")
end if
%>
<script>
function history(fid){
  var temp = "lampinghistory.asp?fid=" +fid  
  //alert(temp)
	window.open(temp,"selectaccount", "scrollbars=yes,width=550, height=300, status=no" );
}

function checkfields(theform){
  retval = true;
  datealert = false;
  for (i=0;i<theform.length;i++){
    if (theform.elements[i].value.indexOf("'") > -1) {
      theform.elements[i].value = theform.elements[i].value.replace(/'/g,"''" );
    }
  }
  dlastdate = new Date(theform.dlast.value);
  dlastdatestr = dlastdate.toGMTString();
  dlastdatearray = dlastdatestr.split(' ');
  if (dlastdatearray[3].length != 4)  { retval = false; datealert = true; };
  if (dlastdate == 'NaN') { retval = false; datealert = true; };
  blastdate = new Date(theform.blast.value);
  blastdatestr = blastdate.toGMTString();
  blastdatearray = blastdatestr.split(' ');
  if (blastdatearray[3].length != 4)  { retval = false; datealert = true; };
  if (blastdate == 'NaN') { retval = false; datealert = true; };
  today = new Date();
  if (datealert) { alert("Please check that you have entered any dates using the format MM/DD/YYYY.")};
  return retval;
}

function checkNumber(thefield){
  re = /\D/;
    bad = re.test(document.forms['form2'].elements[thefield].value);
    if (bad) { 
      document.forms['form2'].elements[thefield].style.backgroundColor = "#ccccff";
      alert("Please only use numbers in this field.");
    } else {
      document.forms['form2'].elements[thefield].style.backgroundColor = "#ffffff"; 
    }
}

function confirmDelete(){
  retval = window.confirm("Are you sure you want to delete this item?");
  return retval;
}
try{top.applabel("Floor Management - View Fixture <%=fid%> in <%=room%> on <%=fl%> floor");}catch(exception){}
//-->
</script>

<title>Fixture Detail</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/styles.css">
</head>
<body bgcolor="#FFFFFF">
 <% if not rst1.EOF  then %>
 <form name="form1" method="post" action="updatesch.asp" onsubmit="return checkfields(this);">
<table width="100%" cellpadding="3" cellspacing="0" border="0" style="border:solid 1px #ffffff;">
  <tr> 
      <td bgcolor="#cccccc" width="47%" nowrap><font color="#000000">
        <b> <a href="floorsearch.asp?bldg=<%=bldg%>" class="breadcrumb">Floor</a>: 
        <%=fl%> | <a href="roomsearch.asp?bldg=<%=bldg%>&floor=<%=fl%>&fid=<%=fid%>" class="breadcrumb">Room</a>: 
        <%=room%> | <a href="fixsearch.asp?bldg=<%=bldg%>&floor=<%=fl%>&room=<%=room%>&fid=<%=fid%>&rid=<%=rid%>" class="breadcrumb">Fixture 
        Code</a>: <%=rst1("fix_catalog")%></b></font></td>
    <td width="53%" align="right" nowrap bgcolor="#cccccc"> 
    <input type="hidden" name="fxid" value="<%=request.querystring("id")%>">
	<% if fid <> "" and rid <> "" and bldg<>"" then %>
    <input type="hidden" name="fid" value="<%=fid%>">
    <input type="hidden" name="bldg" value="<%=bldg%>">
    <input type="hidden" name="rid" value="<%=rid%>">
        <input type="submit" name="Submit" value="Update" class="standard"> &nbsp; 
        <input type="submit" name="Submit" value="Delete" onClick="return confirmDelete();" class="standard"> 
        &nbsp; <input name="reset" type="reset" class="standard" onClick="try{history.back();}catch(exception){location='fixsearch.asp?bldg=<%=bldg%>&room=<%=room%>&floor=<%=fl%>&rid=<%=rid%>&fid=<%=fid%>'}" value="Cancel"> 
	<% end if%>
      </td>
  </tr>
  </table>
  
    
  <table cellpadding="3" cellspacing="1" width="100%" border="0">
    <tr> 
      <td colspan="2"><span class="standard"><b>Fixture (<%=request.querystring("id")%>)</b></span></td>
    </tr>
    <tr> 
      <td  width="35%" align="right" bgcolor="#eeeeee"> <span class="standard"> 
        <input type="hidden" name="lid" value="<%=rst1("id")%>">
        Manufacturer</span></td>
      <td width="65%" bgcolor="#eeeeee"><span class="standard"><%=rst1("manufacturer")%></span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><span class="standard">Fixture Catalog Number</span></td>
      <td bgcolor="#eeeeee"> <span class="standard"><%=rst1("fix_catalog")%></span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><span class="standard">Fixture Description</span></td>
      <td bgcolor="#eeeeee"> <span class="standard"><%=rst1("description")%></span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><span class="standard">FixtureType</span></td>
      <td bgcolor="#eeeeee"><span class="standard"> 
        <input type="text" name="tp" value="<%=rst1("type")%>">
        </span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><span class="standard">Fixture Quantity</span></td>
      <td bgcolor="#eeeeee"><span class="standard"> 
        <input type="text" name="fixqty" value="<%=rst1("fixtureqty")%>">
        </span></td>
    </tr>
    <tr> 
      <td height="10" colspan="2" bgcolor="#eeeeee"></td>
    </tr>
    <tr> 
      <td colspan="2" bgcolor="#cccccc"><span class="standard"><b>Ballast</b></span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><span class="standard">Ballast Type</span></td>
      <td bgcolor="#eeeeee"><span class="standard"><%=rst1("ballast_type")%></span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><font size="2" face="Arial, Helvetica, sans-serif" class="standard">Ballast 
        Quantity</span></td>
      <td bgcolor="#eeeeee"><span class="standard"> <%=rst1("ballastqty")%></span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><font size="2" face="Arial, Helvetica, sans-serif" class="standard">Date 
        Last Changed</span></td>
      <td bgcolor="#eeeeee"><span class="standard"> 
        <input type="text" name="blast" value="<%=rst1("bdatelastchanged")%>">
        <span style="font-size:7pt;">MM/DD/YYYY</span></span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><font size="2" face="Arial, Helvetica, sans-serif" class="standard">Electrician</font></span></td>
      <td bgcolor="#eeeeee"><span class="standard"> 
        <input type="text" name="belect" value="<%=rst1("belectrician")%>">
        </span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><span class="standard">Estimated Replacement Date</span></td>
      <td bgcolor="#eeeeee"><span class="standard"> <%=rst1("bestd")%></span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><font size="2" face="Arial, Helvetica, sans-serif" class="standard">Estimated 
        # Weeks Remaining</font></span></td>
      <td bgcolor="#eeeeee"><span class="standard"> 
        <%
     if not rst1("bweeksr") > 0 then
      %>
        <font color="#ff0000">0</font> (<%=abs(rst1("bweeksr"))%> weeks overdue) 
        <% else %>
        <b><%=rst1("bweeksr")%></b> 
        <% end if %>
        </span></td>
    </tr>
    <tr> 
      <td height="10" colspan="2" bgcolor="#eeeeee"></td>
    </tr>
    <tr> 
      <td bgcolor="#cccccc"><span class="standard"><b>Lamp </b></span></td>
      <td align="right" bgcolor="#cccccc"><span class="standard"><b>
        <input type="button" name="Submit2" value="Lamping History" onClick="history(fxid.value)" class="standard">
        </b></span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><span class="standard">Lamp Catalog Number</span></td>
      <td bgcolor="#eeeeee"><span class="standard"><%=rst1("lamp_catalog")%></span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><span class="standard">Lamp Quantity</span></td>
      <td bgcolor="#eeeeee"><span class="standard"> <%=rst1("lampqty")%> </span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><span class="standard">Lamp Power (W)</span></td>
      <td bgcolor="#eeeeee"><span class="standard"><%=rst1("lamp_watts")%></span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><span class="standard">Voltage (V)</span></td>
      <td bgcolor="#eeeeee"><span class="standard"><%=rst1("volts")%></span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><span class="standard"> Date Last Changed</span></td>
      <td bgcolor="#eeeeee"><span class="standard"> 
        <input type="text" name="dlast" value="<%=rst1("datelastchanged")%>">
        <span style="font-size:7pt;">MM/DD/YYYY</span></span></td>
    </tr>
    <tr> 
      <td  align="right" bgcolor="#eeeeee"><span class="standard"> Electrician</span></td>
      <td bgcolor="#eeeeee"><span class="standard"> 
        <input type="text" name="elect" value="<%=rst1("electrician")%> ">
        </span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><span class="standard">Estimated Lamp Life</span></td>
      <td bgcolor="#eeeeee"><span class="standard"><%=rst1("est_lamp_life")%></span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><span class="standard">Estimated Hours per Week</span></td>
      <td bgcolor="#eeeeee"><span class="standard"> 
        <input type="text" name="ehw" value="<%=rst1("est_hr_wk")%>">
        </span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><span class="standard">Estimated Replacement Date</span></td>
      <td bgcolor="#eeeeee"><span class="standard"><b><%=rst1("estd")%></b></span></td>
    </tr>
    <tr> 
      <td align="right" bgcolor="#eeeeee"><font size="2" face="Arial, Helvetica, sans-serif"><span class="standard">Estimated 
        # Weeks Remaining</span></td>
      <td bgcolor="#eeeeee"><span class="standard"> 
        <%
     if not rst1("weeksr") > 0 then
      %>
        <font color="#ff0000">0</font> (<%=abs(rst1("weeksr"))%> weeks overdue) 
        <% else %>
        <b><%=rst1("weeksr")%></b> 
        <% end if %>
        </span></td>
    </tr>
    <tr> 
      <td height="10" colspan="2" bgcolor="#eeeeee"></td>
    </tr>
    <tr> 
      <td colspan="2" bgcolor="#cccccc"><span class="standard"><b>General Remarks</b></span></td>
    </tr>
    <tr> 
      <td colspan="2" bgcolor="#eeeeee"><span class="standard"><%=rst1("remarks")%></span></td>
    </tr>
    <tr> 
      <td height="18" colspan="2" bgcolor="#eeeeee"></td>
    </tr>
    <tr> 
      <td colspan="2" bgcolor="#cccccc"><span class="standard"><b>Comments</b></span></td>
    </tr>
    <tr> 
      <td colspan="2" bgcolor="#eeeeee"><textarea name="comments"><%=rst1("comments")%></textarea></td>
    </tr>
    <tr> 
      <td height="10" colspan="2" bgcolor="#eeeeee"></td>
    </tr>
    <tr> 
      <td colspan="2" bgcolor="#cccccc">&nbsp;</td>
    </tr>
  </table>
	   	      
 
    
    
<%else%>
  <table width="100%" border="0" cellpadding="3" cellspacing="1">
  <tr> 
    <td height="30"><span class="standard">There is no information for this fixture</span></td>
  </tr>
  </table>
  </form>
<%rst1.close
end if
set cnn1=nothing%>
</HTML>


