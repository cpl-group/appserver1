<%@Language="VBScript"%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!-- #include VIRTUAL="/genergy2/secure.inc" -->
<%
lid 			= trim(request("id"))
serviceid  		= trim(request("serviceid"))
serviceurl 		= trim(request("serviceurl"))
mode 			= trim(request("mode"))
currentbldgid 	= trim(request("pBldgid"))
bldgid			= trim(request("bldgid"))
bldgname		= trim(request("bldgname"))
region			= trim(request("region"))
regioncount		= trim(request("regioncount"))
pregioncount 	= trim(request("pregioncount"))
category 		= trim(request("category"))
catid			= trim(request("catid"))
pcatid			= trim(request("pcatid"))
customlabel		= trim(request("customlabel"))
viewOrdSeq	    = trim(request("DisplayOrdSeq")) 
userid = Session("editemail")
maxcatid = 0

if pregioncount = "" then 
	pregioncount = regioncount
end if

Set cnn1 	= Server.CreateObject("ADODB.Connection")
Set rs 		= Server.CreateObject("ADODB.recordset")

cnn1.Open getConnect(0,0,"dbCore")

if mode = "updateleaf" then 

	strsql = "update clientsetup set serviceid= '"&serviceid&"', serviceurl= '"&serviceurl&"', customlabel= '"&customlabel&"', ViewOrderSeq= '"&viewOrdSeq&"' where id = '" &lid& "' and userid = '"&userid&"'" 
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if
if mode = "updatebranch1" then 

	strsql = "update clientsetup set bldgid= '"&bldgid&"', bldgname= '"&bldgname&"', regioncount = '"&regioncount&"', region= (select distinct region from clientsetup where regioncount = '"&regioncount&"' and userid ='"&userid&"') where bldgid = '" &currentbldgid& "' and userid ='"&userid&"'" 
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if
if mode = "updatebranch2" then 

	strsql = "update clientsetup set region= '"&region&"', regioncount = '"&regioncount&"' where regioncount = '" &pregioncount& "' and userid ='"&userid&"'"  
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if
if mode = "updatebranch2a" then 

	strsql = "update clientsetup set region= (select distinct region from clientsetup where regioncount = '"&regioncount&"' and userid ='"&userid&"'), regioncount = '"&regioncount&"' where regioncount = '" &pregioncount& "' and userid ='"&userid&"'"  
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
 
end if
if mode = "updatebranch3" then 

	strsql = "update clientsetup set category= '"&category&"' where catid= '" &catid& "' and userid ='"&userid&"'"  
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
end if
if mode = "updatebranch3a" then 

	strsql = "update clientsetup set category = (select distinct category from clientsetup where catid = '"&catid&"' and userid ='"&userid&"'), catid = '"&catid&"' where catid= '" &pcatid& "' and userid ='"&userid&"'"  
	tmpMoveFrame =  "parent.document.location = " & Chr(34) & _
				  "./g1nav/g1nav.asp?mode=edit&userid=" & userid & chr(34) & vbCrLf 
end if

cnn1.execute strsql
set cnn1=nothing
Response.Write "<script>" & vbCrLf

Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>