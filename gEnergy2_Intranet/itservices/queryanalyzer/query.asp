<!--#include file="common.asp"-->
<!--#include file="cls/clsRSA.asp"-->
<!--#include file="cls/clsBlowFish.asp"-->
<html>
<head>
<script>
	var bCancel = false;
	parent.nTime = -1;
	parent.Button.FixBorder();
	parent.ResetAlive();
	function t(a1,a2) {parent.SetTitle(a1.split("''"),a2.split("''"))}
	function l(a1) {
		if (bCancel) {return} else {bCancel = parent.AddLine(a1.split("''"))}
	}
</script>
<%
	const adOpenForwardOnly = 0
	const adOpenKeyset = 1
	const adOpenDynamic = 2
	const adOpenStatic = 3
	const adUseServer = 2
	const adUseClient = 3
	const adLockReadOnly = 1
	const adLockPessimistic = 2
	const adLockOptimistic = 3
	const adLockBatchOptimistic = 4

	dim adoConn
	dim adoRS
	dim i, x, y, strAction, strConn, strResponse, strAlign()
	dim strServer, strL, strP, strSQL, strDB, intID, strObject, strName, RSA

	strAction = Request.Querystring("a")
	intID = CInt(Request.Querystring("i"))
	intTRID = CInt(Request.Querystring("id"))
	strObject = Request.Querystring("o")
	strServer = Trim(Request.Querystring("s"))
	strL = Trim(Request.Querystring("l"))
	strP = Decode(Request.Querystring("p"))
	strSQL = Trim(Request.Querystring("sql"))
	strDB = Trim(Request.Querystring("db"))
	bEnc = Request.Querystring("enc")=1
	nEnc = Request.Querystring("e")
	nMod = Request.Querystring("n")

	if strServer = "" or strL = "" then
		strResponse = "<div style='padding:5 5;'>Missing required login information.</div>"
	else
		strDB = iif(strDB="","master",strDB)
		strConn =  "Driver={SQL Server};server="&strServer &";database="&strDB&";uid="&strL&";pwd="&strP
		select case strAction
			case "q": Query
			case "v": bValidate = true: Query
			case "l": Login
			case "b": Browse
		end select
	end if

	Function Decode(strPass)
		if strPass = "" then exit function
		set RSA = new clsRSA
		RSA.KeyMod = clng(Session("RSA_M"))
		RSA.KeyDec = clng(Session("RSA_D"))
		Decode = RSA.Decode(strPass)
		set RSA = Nothing
	End Function

	Sub Browse()
		On Error Resume Next
		Dim Path, Depth, strIcon, strIconOpen, bExpand, strExpand
		Path = split(strObject,".")
		Depth = ubound(Path)
		Select Case Depth
			Case 0:
				strIcon = "_folder"
				strIconOpen = "_folderopen"
				bExpand = true
				strSQL = "select distinct type from "&Path(0)&"..sysobjects where type in ('v','s','u','p','x','fn') order by type"
			Case 1:
				select case Path(1)
					case "U": strIcon = "_table"
					case "S": strIcon = "_table"
					case "V": strIcon = "_view"
					case "P": strIcon = "_sp"
					case "X": strIcon = "_sp"
					case "FN": strIcon = "_function"
				end select
				strIconOpen = strIcon
				bExpand = true
				strSQL = "select name from "&Path(0)&"..sysobjects where type = '"&Path(1)&"' order by name"
			Case 2:
				strIcon = "_column"
				strIconOpen = "_column"
				bExpand = false
				strSQL = "select a.name, c.name as type, a.length, a.isnullable from "&Path(0)&"..syscolumns a "&_
					"inner join "&Path(0)&"..sysobjects b on a.id = b.id "&_
					"inner join systypes c on a.xtype = c.xtype "&_
					"where b.name = '"&Path(2)&"' and b.type = '"&Path(1)&"'"
		End Select

		OpenCon(strConn)
		if adoConn.Errors.count = 1 then
			response.write "<script>window.status='Error: "&FixJS(Err.source)&" Rows';</script>"
			strResponse = strResponse & "<div style='padding:5 5;'>Error Number: "&adoConn.Errors.count&"<br>Source: "&Err.source&"<br>Description: "&Err.description&"</div><br>"
			response.write "<script>parent.divDetail.innerHTML='"&FixJS(strResponse)&"';</script>" & vbcrlf
		Else
			adoRS.open strSQL
			if adoRS.recordcount > 0 then
				strList = "<table cellpadding=0 cellspacing=0 border=0 style='padding:0 0;'>"
				do until adoRS.eof
					select case Depth
						case 0:
							strName = Trim(UCase(adoRS("type")))
							select case strName
								case "U": strPath = strObject &"."& strName: strName = "User Tables"
								case "S": strPath = strObject &"."& strName: strName = "System Tables"
								case "V": strPath = strObject &"."& strName: strName = "Views"
								case "P": strPath = strObject &"."& strName: strName = "Stored Procedures"
								case "X": strPath = strObject &"."& strName: strName = "Extended Procedures"
								case "FN": strPath = strObject &"."& strName: strName = "Functions"
							end select
						case 1:
							strName = adoRS("name")
							strPath = strObject &"."& strName
						case 2:
							strName = adoRS("name")
							strPath = strObject &"."& strName
							strName = strName & " ("&adoRS("type")&"("&adoRS("length")&"), "&iif(adoRS("isnullable"),"Null","Not Null")&")"
					end select
					adoRS.movenext
					ImageEnd = iif(adoRS.eof,"End","")
					if bExpand then
						strExpand = "_plus"
						OpenLink = "'FolderState("&intID&",true)' class=hand"
					else
						strExpand = "_blank"
						OpenLink = "'' "
					end if
					strList = strList & "<tr><td style='padding-left:2;'>"&_
						"<img id=img0_"&intID&" src=images/"&strExpand&ImageEnd&".gif width=16 height=16  onclick="&OpenLink&">"&_
						"<img id=img1_"&intID&" src=images/_minus"&ImageEnd&".gif style=display:none; width=16 height=16 class=hand onclick='FolderState("&intID&",false)'></td>"&_
						"<td style='padding-right:5;'  ondblclick="&OpenLink&">"&_
						"<img id=img2_"&intID&" src='images/"&strIcon&".gif' width=16 height=16 border=0>"&_
						"<img id=img3_"&intID&" src='images/"&strIconOpen&".gif' style=display:none; width=16 height=16 border=0></td>"&_
						"<td width=100% nowrap>"&strName&"</td>"&_
						"</tr><tr id=tr"&intID&" style='display:none;' link="""&strPath&""">"&_
						"<td style='width:22;background-position:2 0;background-image:url(images/_vert"&ImageEnd&".gif);'>&nbsp;</td>"&_
						"<td colspan=2 id=td"&intID&"></td></tr>"
					intID = intID + 1
				loop
				strList = strList & "</table>"
			else
				strList = "<table cellpadding=0 cellspacing=0 border=0 style='padding:0 0;height:16;'>"
				strList = strList & "<tr><td style='padding-left:5;'>"&_
						"<img src=images/_blankend.gif width=16 height=16></td>"&_
						"<td style='padding-right:5;'>None</td></tr>"
				strList = strList & "</table>"
			end if
			response.write "<script>parent.td"&intTRID&".innerHTML='"&FixJS(strList)&"';" & vbcrlf
			response.write "parent.ShowFolder("&intTRID&");" & vbcrlf
			response.write "parent.intID="&intID&";</script>" & vbcrlf
		end if
	End Sub

	Sub Login()
		On Error Resume Next
		Dim strList, strSelect, strDisplay, strDisplay2, OpenLink, strDisabled, strDisabled2
		OpenCon(strConn)
		if adoConn.Errors.count = 1 then
			strDisplay = "none"
			strDisplay2 = "inline"
			strDisabled = "false"
			response.write "<script>window.status='Error: "&FixJS(Err.source)&" Rows';</script>"
			strResponse = strResponse & "<div style='padding:5 5;'>Error Number: "&adoConn.Errors.count&"<br>Source: "&Err.source&"<br>Description: "&Err.description&"</div><br>"
		Else
			strSQL = "select name from sysdatabases where HAS_DBACCESS (name) = 1 order by name"
			adoRS.open strSQL
			if adoRS.recordcount > 0 then
				strList = "<table cellpadding=0 cellspacing=0 border=0 style='padding:0 0;'>"&_
					"<tr><td style='padding-left:2;' "&OpenLink&"><img src='images/_server.gif' width=16 height=16 border=0></td>"&_
					"<td colspan=2>"&strServer&"</td></tr>"

				strSelect = "<select id=selDB class=InputSmall>"
				do until adoRS.eof
					strName = adoRS("name")
					adoRS.movenext
					strSelect = strSelect & "<option value="""&strName&""""&iif(lcase(strName) = "master"," selected","")&">"&strName
					ImageEnd = iif(adoRS.eof,"End","")
					OpenLink = "'FolderState("&intID&",true)' class=hand"
					strList = strList & "<tr><td style='padding-left:5;'>"&_
						"<img id=img0_"&intID&" src=images/_plus"&ImageEnd&".gif width=16 height=16  onclick="&OpenLink&">"&_
						"<img id=img1_"&intID&" src=images/_minus"&ImageEnd&".gif style=display:none; width=16 height=16 class=hand onclick='FolderState("&intID&",false)'></td>"&_
						"<td style='padding-right:5;'  ondblclick="&OpenLink&">"&_
						"<img id=img2_"&intID&" src='images/_db.gif' width=16 height=16 border=0>"&_
						"<img id=img3_"&intID&" src='images/_db.gif' style=display:none; width=16 height=16 border=0></td>"&_
						"<td width=100% nowrap>"&strName&"</td>"&_
						"</tr><tr id=tr"&intID&" style='display:none;' link="""&strName&""">"&_
						"<td style='width:22;background-position:5 0;background-image:url(images/_vert"&ImageEnd&".gif);'>&nbsp;</td>"&_
						"<td colspan=2 id=td"&intID&"></td></tr>"
					intID = intID + 1
				loop
				strList = strList & "</table>"
				strSelect = strSelect & "</option>"
				strDisplay = "inline"
				strDisplay2 = "none"
				strDisabled = "true"
			end if
			response.write "<script>window.status='["&strServer&"] Connected';" & vbcrlf
			response.write "parent.document.title='Remote Query Analyzer - ["&strServer&"]';</script>" & vbcrlf
		end if
		response.write "<script>parent.ShowExec("&iif(strDisplay="inline",1,0)&");"
		response.write "parent.tdDB.innerHTML='"&FixJS(strSelect)&"';"
		response.write "parent.divList.innerHTML='"&FixJS(strList)&"';"
		response.write "parent.intID="&intID&";" & vbcrlf
		response.write "parent.divDetail.innerHTML='"&FixJS(strResponse)&"';" & vbcrlf
		response.write "parent.SaveCon();</script>" & vbcrlf
		CloseCon
	End Sub

	Sub Query()
		On Error Resume Next
		if strSQL = "" then
			strResponse = "<div style='padding:5 5;'>SQL statement empty.</div>"
		else
			OpenCon(strConn)
			if bValidate then adoRS.open "SET NOEXEC ON"
			adoRS.open strSQL
			if adoConn.Errors.count > 0 then
				response.write "<script>window.status='Error: "&FixJS(Err.source)&" Rows'</script>"
				strResponse = strResponse & "<div style='padding:5 5;'>Error Number: "&adoConn.Errors.count&"<br>Source: "&Err.source&"<br>Description: "&Err.description&"</div><br>"
			ElseIf adoRS.state then
				y = adoRS.fields.count -1
				redim strAlign(y)
				for i = 0 to y
					select case adoRS.fields(i).type
						case 7: strAlign(i) = ""
						case 8: strAlign(i) = ""
						case 11: strAlign(i) = "center"
						case 64: strAlign(i) = ""
						case 72: strAlign(i) = ""
						case 128: strAlign(i) = "Binary"
						case 129: strAlign(i) = ""
						case 130: strAlign(i) = ""
						case 133: strAlign(i) = ""
						case 134: strAlign(i) = ""
						case 135: strAlign(i) = ""
						case 200: strAlign(i) = ""
						case 201: strAlign(i) = ""
						case 202: strAlign(i) = ""
						case 203: strAlign(i) = ""
						case 204: strAlign(i) = "Binary"
						case 205: strAlign(i) = "Binary"
						case else: strAlign(i) = "right"
					end select
					strName = adoRS.fields(i).name

					strFields1 = strFields1 & iif(strName="","''&nbsp;","''"&FixJS(strName))
					strFields2 = strFields2 & iif(strAlign(i)="Binary","''","''"&strAlign(i))
				next
				response.write "<script>t("""&strFields1&""","""&strFields2&""")</script>" & vbcrlf

				if adoRS.recordcount > 0 then
					x=0
					response.write "<script>"
					if bEnc then
						set RSA = new clsRSA
						RSA.KeyMod = clng(nMod)
						RSA.KeyEnc = clng(nEnc)
					end if
					do until adoRS.eof
						if bEnc then
							If x mod 2 = 1 then: response.write "</script>" & vbcrlf: response.flush: response.write "<script>"
						else
							If x mod 5 = 4 then: response.write "</script>" & vbcrlf: response.flush: response.write "<script>"
						end if
						strFields1 = ""
						strFields2 = ""
						for i = 0 to adoRS.fields.count -1
							if strAlign(i) = "Binary" then
								strValue = "&lt;binary&gt;"
							else
								strValue = adoRS.fields(i)
								If isnull(strValue) then strValue = "NULL"
							end if

							IF bEnc then
								strFields1 = strFields1 & "''"&RSA.Encode(strValue)
							else
								strFields1 = strFields1 & "''"&Server.HTMLEncode(strValue)
							end if
						next
						response.write "l("""&strFields1&""");"
						adoRS.movenext
						x=x+1
					loop
					if bEnc then set RSA = Nothing
					response.write "</script>" & vbcrlf
				end if
			ElseIF bValidate then
				strResponse = "<div style='padding:5 5;'>The command(s) parsed successfully.</div>"
			else
				strResponse = "<div style='padding:5 5;'>The command(s) completed successfully.</div>"
			end if
			if bValidate then adoRS.open "SET NOEXEC OFF"
			CloseCon
		end if
		if Err > 0 then
			response.write "<script>window.status='Error: "&FixJS(Err.source)&"'</script>"
			strResponse = strResponse & "<div style='padding:5 5;'>Error Number: "&Err.number&"<br>Source: "&Err.source&"<br>Description: "&Err.description&"</div><br>"
		end if
		if strResponse <> "" then response.write "<script>parent.divDetail.innerHTML='"&FixJS(strResponse)&"'</script>" & vbcrlf
		response.write "<script>parent.ShowExec(1);</script>" & vbcrlf
	End Sub

	sub OpenCon(strConnection)
		Set adoConn = server.createobject("ADODB.Connection")
		adoConn.ConnectionString = strConnection
		adoConn.ConnectionTimeout = 5
		adoConn.CommandTimeout = 120
		adoConn.Open
		adoConn.BeginTrans
		Set adoRS = server.CreateObject("ADODB.Recordset")
		adoRS.ActiveConnection = adoConn
		adoRS.CursorType = adOpenForwardOnly
		adoRS.locktype = adLockBatchOptimistic
		adoRS.CursorLocation = adUseClient
	end sub

	sub ResetCon()
		If adoRS.state then adoRS.close
	end sub

	sub CloseCon()
		If adoConn.Errors.Count > 0 Or Err.Number <> 0 Then
			adoConn.RollbackTrans
		Else
			adoConn.CommitTrans
		End If
		If adoRS.state then adoRS.close
		Set adoRS = nothing
		Set adoConn = nothing
	end sub
%>
</head>
</html>