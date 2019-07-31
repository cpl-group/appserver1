<%@Language="VBScript"%>
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
'2/15/2008 N.Ambo amended sql for 'addbranch1' so that the correct buildingorder will be inserted for the new branch (it was picking up the same building order for each new branch)

mode 		= trim(request("mode"))
userid 		= Session("editemail")
lid  		= trim(request("lid"))
mv_dir 		= trim(request("mv_dir"))
catid		= trim(request("catid"))
regioncount = trim(request("regioncount"))
maxcatid = 0

Set cnn1 	= Server.CreateObject("ADODB.Connection")
Set rs 		= Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"dbCore")

	strsql = "select top 1 * from clientsetup where userid = '"&userid&"' order by catid desc"
	
	rs.Open strsql, cnn1, adOpenStatic
	
	if not rs.EOF then 
		maxcatid = rs("catid") + 1
	end if
	
	rs.close
 	strsql = "select top 1 * from clientsetup where userid = '"&userid&"' order by regioncount desc, id desc"
	
	rs.Open strsql, cnn1, adOpenStatic
	
	if not rs.EOF then 
		maxregioncount 	= rs("regioncount") + 1
		maxid 			= rs("id") + 1
	end if
	rs.close

if mode = "addcategory" then 

	strsql = "insert into clientsetup (userid, category, catid, region, regioncount,bldgid, bldgname) values ('"&userid&"', 'New Category', '" &maxcatid& "','New Region','"&maxregioncount&"','new"&maxid&"','New Building')"
	tmpMoveFrame =  "document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if

if mode = "mts" then

	strsql = "update clientsetup set bldgid = 'new"&maxid&"', bldgname = 'New Building', regioncount='"&maxregioncount&"' where id = '" &lid& "' and userid = '"&userid&"'" 
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 

end if

if mode = "bmts" then

	strsql = "update clientsetup set regioncount=  '"&maxregioncount&"', region = 'New Region' where id = '" &lid& "' and userid = '"&userid&"'" 
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 

end if
if mode = "rmts" then

	strsql = "update clientsetup set catid= '" &maxcatid& "', category= 'New Category' where regioncount = '" &lid& "' and userid = '"  & userid &"'"
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 

end if
if mode = "rmtp" then

	strsql = "update clientsetup set catid= '-1', category= '-1' where id = '" &lid& "' and userid = '"&userid&"'" 
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 

end if
if mode = "bmtp" then

	strsql = "update clientsetup set region= 'New Region' where id = '" &lid& "' and userid = '"&userid&"'" 
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 

end if
if mode = "lmtp" then

	strsql = "update clientsetup set bldgid = '-1' where id = '" &lid& "' and userid = '"&userid&"'"
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 

end if

if mode = "addentry" then 
	
	strsql = "insert into clientsetup (userid, category, catid, region,regioncount,bldgid, bldgname,bldgorder) select userid, category, catid, region,regioncount,bldgid, bldgname,bldgorder from clientsetup where id ='" &lid& "' and userid = '"&userid&"'"
		tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if
if mode = "initentry" then 
	strsql = "insert into clientsetup (userid) values ('"&userid&"')"
	tmpMoveFrame =  "document.location = " & Chr(34) & _
			  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
end if
if mode = "addbranch1" then 
	
	strsql = "insert into clientsetup (userid, category, catid, region,regioncount,bldgid, bldgname,bldgorder) select userid, category, catid, region, regioncount, 'new"&maxid&"' as bldgid, 'New Building' as bldgname, bldgorder + 1 as bldgorder from clientsetup where id ='" &lid& "' and userid = '"&userid&"'" 
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if
if mode = "addbranch1a" then 
	
	'strsql = "insert into clientsetup (userid, category, catid, region,regioncount,bldgid, bldgname) select userid, category, catid, region, '"&maxregioncount&"' as regioncount,'new"&maxid&"' as bldgid, bldgname from clientsetup where id ='" &lid& "' and userid = '"&userid&"'" 
	strsql = "insert into clientsetup (userid, category, catid, region,regioncount,bldgid, bldgname,bldgorder) select userid, category, catid, region, regioncount, 'new"&maxid&"' as bldgid, 'New Building' as bldgname, max(bldgorder) + 1 as bldgorder from clientsetup where userid = '"&userid&"' group by userid, category, catid,region,regioncount" 
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if
if mode = "addbranch2" then 

	strsql = "insert into clientsetup (userid, category, catid, region,regioncount,bldgid, bldgname) select userid, category, catid, 'New Region' as region,"& maxregioncount&" as regioncount,'new"&maxid&"' as bldgid, 'New Building' as bldgname from clientsetup where id ='" &lid& "' and userid = '"&userid&"'" 
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if
if mode = "addbranch3" then 

	strsql = "insert into clientsetup (userid, category, catid, region, regioncount,bldgid, bldgname) values ('"&userid&"', 'New Category', '" &maxcatid& "','New Region','"&maxregioncount&"','new"&maxid&"','New Building')"
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if

if mode = "delentry" then 
	
	strsql = "delete from clientsetup where id ='" &lid& "' and userid = '"&userid&"'" 
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if
if mode = "delbranch1" then 
	
	strsql = "delete from clientsetup where bldgid ='" &lid& "' and userid = '"&userid&"'" 
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if
if mode = "delbranch2" then 
	
	strsql = "delete from clientsetup where regioncount ='" &lid& "' and userid = '"&userid&"'" 
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if
if mode = "delbranch3" then 
	
	strsql = "delete from clientsetup where catid ='" &lid& "' and userid = '"&userid&"'" 
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if
if mode = "srvorder" then 
	
	strsql = "update clientsetup set customlabel = bldgname where regioncount = " & cint(lid) & " and userid = '"&userid&"'" 
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if
if mode = "bldgorder" then 
	
	strsql = "update clientsetup set customlabel = -1 where regioncount = " & cint(lid) & " and userid = '"&userid&"'"
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if
if mode = "moveregion" then 
	
	strsql = "exec sp_region " & cint(lid) & ","&catid&",'" & userid & "','"&mv_dir&"'"
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if
if mode = "movebuilding" then 
	
	strsql = "exec sp_branch " & cint(lid) & ","&regioncount&",'" & userid & "','"&mv_dir&"'"
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if
if mode = "movecategory" then 
	
	strsql = "exec sp_category " & cint(catid) & ",'" & userid & "','"&mv_dir&"'"
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if
cnn1.execute strsql
set cnn1=nothing

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>