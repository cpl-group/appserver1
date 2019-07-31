<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
<%
bldgnum=Request.form("bldgnum")
item=Request.form("item")
sqft=request.Form("sqft")
floor=request.form("floor")
riser=request.form("riser")
action=request.form("submit")
size=request.form("size")
metal=request.form("metal")
insulation=request.form("insulation")
sets=request.form("sets")
volts=request.form("volts")
sw_frame=request.form("sframe")
sw_fuse=request.form("sfuse")
Set cnn1 = Server.CreateObject("ADODB.Connection")
cnn1.Open application("cnnstr_capacity_db")
if action="Save" then
	if item="riser" then
		sql = "Insert into tblriser (riser_name, bldgnum, size, metal, insulation, sets, volts, sw_frame, sw_fuse) "_
	& "values ("_
	& "'"&riser&"', "_
	& "'"&bldgnum&"', "_
	& "'"&size&"', "_
	& "'"&metal&"', "_
	& "'"&insulation&"', "_
	& "'"&sets&"', "_
	& "'"&volts&"', "_
	& "'"&sw_frame&"', "_
	& "'"&sw_fuse&"')"
	else
		sql = "Insert into tblfloor (bldgnum, fl_name, sqft) "_
	& "values ("_
	& "'"&bldgnum&"', "_
	& "'"&floor&"', "_
	& "'"&sqft&"')"
  	end if
else
	if item="riser" then
		sql = "Update tblriser Set size='"&size&"', metal='"&metal&"', insulation='"& insulation&"', sets='"&sets&"', volts='"&volts&"',sw_frame='"&sw_frame&"', sw_fuse='"& sw_fuse&"' where bldgnum='"& bldgnum&"' and riser_name='"&riser&"'"
	else
	    sql = "Update tblfloor Set sqft='"&sqft&"' where bldgnum='"&bldgnum&"' and fl_name='"& floor &"'"
	end if
end if
'response.write sql
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