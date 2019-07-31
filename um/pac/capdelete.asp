<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
		if isempty(Session("name")) then
%>
<script>
//top.location="../index.asp"
//window.close()
</script>
<%
			'			Response.Redirect "http://www.genergyonline.com"
		else
			if Session("ts") < 4 then 
				Session("fMessage") = "Sorry, the module you attempted to access is unavailable to you."
				Response.Redirect "../main.asp"
			end if	
		end if	
user=Session("name")
uid="ghnet\"&Session("login")
Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.Recordset")
i=0
cnn1.Open getConnect(0,0,"engineering")
bldgnum=trim(request("bldgnum"))
buildingip = getbuildingip(bldgnum)
val=trim(request("val"))
item=trim(request("item"))
check=request("check")
Ary=split(request.form("list"),",")
if check =1 then
  	if item="floor" then
		floor=val
		sql="delete from tblfloor where bldgnum='"&bldgnum&"' and fl_name='"&val&"'"
		
	else
		riser=val
		
		sql2="select count(*) as c from ["&buildingip&"].dbBilling.dbo.meters m where m.riser='"& riser&"'" 
		rst1.Open sql2, cnn1, 0, 1, 1
		count=rst1("c")
		if count>0 then
			rst1.close
			msg="There are "&count&" associated to this riser"
			response.redirect "capdetail.asp?bldgnum="&bldgnum&"&riser="&riser&"&item="&riser&"&msg="&msg
		else
			sql="delete from tblriser where bldgnum='"&bldgnum&"' and riser_name='"&val&"'"
		end if
	end if
	cnn1.execute sql
else
	if ubound(Ary) = -1 then
		response.redirect "capnewitem.asp?bldgnum="&bldgnum&"&item=riser&riser="&riser
	end if
	if item="riser" then
		riser=val
		sql2="select count(*) as c from ["&buildingip&"].dbBilling.dbo.meters m where m.riser='"& riser&"'" 
		rst1.Open sql2, cnn1, 0, 1, 1
		count=rst1("c")
		if count>0 then
			rst1.close
			msg="There are "&count&" associated to this riser"
			response.redirect "capnewitem.asp?bldgnum="&bldgnum&"&item=riser&riser="&riser&"&msg="&msg
		end if
    	for i= lbound(Ary) to ubound(Ary)
        	floor=trim(Ary(i))
			sql="delete from tblassociation where bldgnum='"&bldgnum&"' and riser_name='"&val&"' and fl_name='"&floor&"'"
			cnn1.execute sql
			'response.write sql
		next
	else
		floor=val
	    for i= lbound(Ary) to ubound(Ary)
    	    riser=trim(Ary(i))
			sql = "delete from tblassociation where bldgnum='"&bldgnum&"' and riser_name='"&riser&"' and fl_name='"&val&"'"
			cnn1.execute sql
			'response.write sql
		next
	end if
end if
set cnn1=nothing
tmpMoveFrame =  "parent.floor.location = " & Chr(34) & _
				  "capfloor.asp?bldgnum="&bldgnum& chr(34) & vbCrLf 
tmpMoveFrame2 =  "parent.riser.location = " & Chr(34) & _
				  "capriser.asp?bldgnum="&bldgnum& chr(34) & vbCrLf 				  
Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write tmpMoveFrame2
Response.Write "</script>" & vbCrLf  
%>
