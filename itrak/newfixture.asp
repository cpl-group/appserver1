<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include virtual="/genergy2/secure.inc" -->
<html>
<head>

<%
		if isempty(Session("name")) then
'			Response.Redirect "../index.asp"
		else
			if Session("opslog") < 2 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."

				Response.Redirect "../main.asp"
			end if	
		end if		
		user=Session("name")
		Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open getconnect(0,0,"engineering")
		
dim bldg, room, fl

bldg=request.querystring ("bldg")
room=request.querystring ("room")
fl=request.querystring ("fl")
fid=request.querystring ("fid")
rid=request.querystring ("rid")
	
%>

<title>New Fixture</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="/genergy2/styles.css">
<script language="JavaScript">
try{top.applabel("Floor Management - Add Fixture in <%=room%> on <%=fl%> floor");}catch(exception){}
//<!--
function checkRequired(){
  frm = document.forms["form2"];
  retval = true;
  if (frm.esthwk.value=="") { retval = false; }
  if (frm.fixqty.value=="") { retval = false; }
  if (frm.dlast.value=="") { retval = false; }
  newdate = new Date(frm.dlast.value);
  newdatestr = newdate.toGMTString();
  newdatearray = newdatestr.split(' ');
  if (newdatearray[3].length != 4)  { retval = false; };
//  alert(newdate);
  if (newdate == 'NaN') { retval = false };
  if (!retval) { alert("Please fill out all fields and make sure you have entered a valid date in the form of MM/DD/YYYY. Only comments are optional."); }
  return retval;
}
//-->
</script>
<script src="messages.js" type="text/javascript" language="Javascript1.2"></script>
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form name="form2" method="post" action="savefix.asp" onsubmit="return checkRequired();">

  <table width="100%" cellpadding="3" cellspacing="0" border="0" style="border:solid 1px #ffffff;">
    <tr valign="top"> 
      <td width="44%" align="left" nowrap bgcolor="#FFFFFF"><span class="standardheader"><font color="#000000">New 
          Fixture | <span class="standard"><b><a href="floorsearch.asp?bldg=<%=bldg%>" class="breadcrumb">Floor</a>: 
          <%=fl%> | <a href="roomsearch.asp?bldg=<%=bldg%>&floor=<%=fl%>&fid=<%=fid%>" class="breadcrumb">Room</a>: 
          <%=room%></b></span></font></span></td>
      <td width="45%" align="right" bgcolor="#FFFFFF"><font face="Arial, Helvetica, sans-serif"><span class="standard">
        <input type="submit" name="choice22"  value="Save"  class="standard">
        &nbsp;
        <input name="reset" type="reset" class="standard" onClick="location='fixsearch.asp?bldg=<%=bldg%>&rid=<%=rid%>&fid=<%=fid%>'" value="Cancel">
        </span></font></td>
    </tr>
  </table>

  <table width="100%" cellpadding="3" cellspacing="1" border="0">
    <tr bgcolor="#eeeeee"> 
      <td align="right"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"> 
        <input type="hidden" name="bldg" value="<%=request.querystring("bldg")%>">
        <input type="hidden" name="room" value="<%=request.querystring("room")%>">
        <input type="hidden" name="floor" value="<%=request.querystring("fl")%>">
        <input type="hidden" name="fid" value="<%=request.querystring("fid")%>">
        <input type="hidden" name="rid" value="<%=request.querystring("rid")%>">
        Fixture</span></font></td>
	<td><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard">
        <select name="type">
          <%
//			sqlstr = "select  fix_catalog+' '+lamp_catalog as type ,id from fixture_types "
			sqlstr = "select distinct description+' ('+fix_catalog+')' as type ,FT.id, description from fixture_types ft join facilityinfo f on  f.CLIENTID=ft.CLIENT WHERE FT.CLIENT=(SELECT CLIENTID FROM FACILITYINFO WHERE ID='"&BLDG&"') order by description "
   			rst1.Open sqlstr, cnn1, 0, 1, 1
			if not rst1.eof then
				do until rst1.eof	
		%>
          <option value="<%=rst1("id")%>"><%=rst1("type")%></option>
          <%
					rst1.movenext
					loop
					end if
					%>
        </select><br>
        <a onMouseOut="closeHelpBox()" onMouseOver="helpbox('new_fixture_in_rm',event.x,event.y)"><img src="images/question.gif" width="13" height="13" hspace="4" border="0"></a><span class="standard" style="font-size:7pt;">Lamp Type (Manufacturer/Catalog Number)</span>
        </span></font></td>
    </tr>
	 <tr bgcolor="#eeeeee"> 
      <td width="35%" align="right"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard">Fixture Type</span></font></td>
      <td width="65%"  bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"> 
        <input type="text" name="ft" size="5" maxlength="5" >
        </span></font></td>
	</tr>
    <tr bgcolor="#eeeeee"> 
      <td width="35%" align="right"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard">Estimated 
        Hours per Week</span></font></td>
      <td width="65%"  bgcolor="#eeeeee"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"> 
        <input type="text" name="esthwk" size="5" maxlength="5" >
        </span></font></td>
	</tr>
	<tr bgcolor="#eeeeee">
      <td align="right"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard">Fixture 
        Quantity</span></font></td>
      <td  bgcolor="#eeeeee" > <font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"> 
        <input type="text" name="fixqty" size="5" maxlength="5" >
        </span></font></td>
	</tr>
	<tr bgcolor="#eeeeee">
      <td align="right"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard">Date Last Changed</span></font></td>
      <td  bgcolor="#eeeeee" > <font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"> 
        <input type="text" name="dlast" value="<%=date()%>" size="10" > <span style="font-size:7pt;">MM/DD/YYYY</span>
        </span></font></td>
	</tr>
<!--
	[[tr bgcolor="#eeeeee"]]
      [[td align="right"]][[font face="Arial, Helvetica, sans-serif" size="2"]][[span class="standard"]]Lamp 
        Quantity[[/span]][[/font]][[/td]]
      [[td bgcolor="#eeeeee"]][[font face="Arial, Helvetica, sans-serif" size="2"]][[span class="standard"]] 
        [[input type="text" name="lampqty" size="5" maxlength="5" ]]
        [[/span]][[/font]][[/td]]
	[[/tr]]
	[[tr bgcolor="#eeeeee"]]
      [[td align="right"]][[font size="2" face="Arial, Helvetica, sans-serif"]][[span class="standard"]]Ballast 
        Quantity[[/span]][[/font]][[/td]]
     [[td  bgcolor="#eeeeee"]][[font face="Arial, Helvetica, sans-serif" size="2"]][[span class="standard"]] 
        [[input type="text" name="balqty" size="5" maxlength="5" ]]
        [[/span]][[/font]][[/td]]
	[[/tr]]
-->
	<tr bgcolor="#eeeeee">
	   <td  align="right"><font face="Arial, Helvetica, sans-serif" size="2"><span class="standard">Comments</span></font></td>
       <td  bgcolor="#eeeeee"> <font face="Arial, Helvetica, sans-serif" size="2"><span class="standard"> 
        <textarea name="comments"></textarea><span class="standard" style="font-size:7pt;">&nbsp;(Optional)</span>
        </span></font></td>
    </tr>
    <tr bgcolor="#cccccc">
		<td></td>
      <td> <font face="Arial, Helvetica, sans-serif"><span class="standard"> </span></font></td>
    </tr>
  </table>
</form>
<!--#INCLUDE FILE="helpbox.htm"-->
</body>
</html>
