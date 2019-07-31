<%@Language="VBScript"%>
<!-- #include file="adovbs.inc" -->
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

bldgnum=trim(request.form("bldgnum"))
riser=trim(request.form("riser"))
floor=trim(request.form("floor"))
item=trim(request.form("item"))
action=trim(request.form("submit"))


if action="-> Del" then
	Ary=split(request.form("exist"),",")
	if item="riser" then
		if ubound(Ary) = -1 then
			response.redirect "capnewitem.asp?bldgnum="&bldgnum&"&item=riser&riser="&riser
		end if
		for i= lbound(Ary) to ubound(Ary)
        	floor=trim(Ary(i))
			sql = "delete from tblassociation where bldgnum='"&bldgnum&"' and riser_name ='"&riser& "' and fl_name='"&floor&"'"
			cnn1.execute sql
			'response.write sql
			'esponse.end
		next
		set cnn1=nothing
		response.redirect "capnewitem.asp?bldgnum="&bldgnum&"&item=riser&riser="&riser
	else
	    if ubound(Ary) = -1 then
			response.redirect "capnewitem.asp?bldgnum="&bldgnum&"&item=floor&floor="&floor
	    end if
		for i= lbound(Ary) to ubound(Ary)
        	riser=trim(Ary(i))
			sql = "delete from tblassociation where bldgnum='"&bldgnum&"' and riser_name ='"&riser& "' and fl_name='"&floor&"'"
			cnn1.execute sql
			'response.write sql
			'response.end
		next
		set cnn1=nothing
		response.redirect "capnewitem.asp?bldgnum="&bldgnum&"&item=floor&floor="&floor
	end if
	
else
	Ary=split(request.form("list"),",")
	if item="riser" then
		if ubound(Ary) = -1 then
			response.redirect "capnewitem.asp?bldgnum="&bldgnum&"&item=riser&riser="&riser
		end if
    	for i= lbound(Ary) to ubound(Ary)
			floor=trim(Ary(i))
        	sql = "Insert into tblassociation ( bldgnum,riser_name, fl_name) "_
			& "values ('"&bldgnum&"', '"&riser& "', '"&floor&"')"
			cnn1.execute sql
			
		next
		set cnn1=nothing
		response.redirect "capnewitem.asp?bldgnum="&bldgnum&"&item=riser&riser="&riser
'	tmpMoveFrame =  "opener.parent.floor.location = " & Chr(34) & _
'				  "capfloor.asp?bldgnum="&bldgnum&"&riser="&riser& chr(34) & vbCrLf 
	else
		if ubound(Ary) = -1  then
			response.redirect "capnewitem.asp?bldgnum="&bldgnum&"&item=floor&floor="&floor
		end if
    	for i= lbound(Ary) to ubound(Ary)
        	riser=trim(Ary(i))
			sql = "Insert into tblassociation ( bldgnum,riser_name, fl_name) "_
			& "values ('"&bldgnum&"', '"&riser& "', '"&floor&"')"
			cnn1.execute sql
			'response.write sql
		next
		set cnn1=nothing
		response.redirect "capnewitem.asp?bldgnum="&bldgnum&"&item=floor&floor="&floor
		tmpMoveFrame =  "opener.parent.riser.location = " & Chr(34) & _
					  "capriser.asp?bldgnum="&bldgnum&"&riser="&floor& chr(34) & vbCrLf 
	end if
end if

'Response.Write "<script>" & vbCrLf
'Response.Write tmpMoveFrame
'Response.Write "</script>" & vbCrLf  

%>
<script>
//window.close()
</script>