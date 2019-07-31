<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
bldgnum=Request.form("bldgnum")
item=Request.form("item")
sqft=request.Form("sqft")
floor=left(request.form("floor"),30)
riser=left(request.form("riser"),30)
action=request.form("submit")
size=request.form("size")
metal=request.form("metal")
insulation=request.form("insulation")
sets=request.form("sets")
volts=request.form("volts")
sw_frame=request.form("sframe")
sw_fuse=request.form("sfuse")
note=request.form("note")
power_factor=request.form("powerfactor")
riser_length=request.form("riserlength")
onum=request.form("onum")
include=request.form("include")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open getConnect(0,0,"engineering")

if action="Save" then

	if item="riser" then
		sql = "Insert into tblriser (riser_name, bldgnum, size, metal, insulation, sets, volts, sw_frame, sw_fuse,note,power_factor,riser_length) "_
	& "values ("_
	& "'"&riser&"', "_
	& "'"&bldgnum&"', "_
	& "'"&size&"', "_
	& "'"&metal&"', "_
	& "'"&insulation&"', "_
	& "'"&sets&"', "_
	& "'"&volts&"', "_
	& "'"&sw_frame&"', "_
	& "'"&sw_fuse&"', "_
	& "'"&note&"', "_
	& "'"&power_factor&"', "_
	& "'"&riser_length&"')"
	else
		sql = "Insert into tblfloor (bldgnum, fl_name, sqft,include,orderno) "_
	& "values ("_
	& "'"&bldgnum&"', "_
	& "'"&floor&"', "_
	
	& "'"&sqft&"', "_
	& "'"&include&"', "_
	& "'"&onum&"')"
  	end if

	else
	if item="riser" then
		sql = "Update tblriser Set size='"&size&"', metal='"&metal&"', insulation='"& insulation&"', sets='"&sets&"', volts='"&volts&"',sw_frame='"&sw_frame&"', sw_fuse='"& sw_fuse&"', note='"& note&"', power_factor='"& power_factor&"', riser_length='"& riser_length&"' where bldgnum='"& bldgnum&"' and riser_name='"&riser&"'"
	else
	    sql = "Update tblfloor Set sqft='"&sqft&"',orderno='"&onum&"',include='"&include&"' where bldgnum='"&bldgnum&"' and fl_name='"& floor &"'"
	end if
end if
'response.write sql
'response.end

cnn1.execute sql

set cnn1=nothing


if item="riser" then
	tmpMoveFrame =  "parent.frames.riser.location = " & Chr(34) & _
				  "capriser.asp?bldgnum="&bldgnum& chr(34) & vbCrLf 
else
	tmpMoveFrame =  "parent.frames.floor.location = " & Chr(34) & _
				  "capfloor.asp?bldgnum="&bldgnum &chr(34) & vbCrLf 
end if
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf  
%>