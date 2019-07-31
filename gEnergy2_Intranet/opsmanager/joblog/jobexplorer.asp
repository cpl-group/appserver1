<%@ Language=VBScript %>
<%Option Explicit%>
<%Response.Buffer = true%>
<%
	' For disabling password security set wexPassword = "" 
	Const wexPassword = ""
	' Root folder, it can be a physical or virtual folder: "c:\test", "/test"
	' Beware that you will not have web access(no downloading/viewing of files) if you use a physical folder like "c:\test"
	' If you want to have web access with a physical folder, you should create a virtual folder(IIS alias like "/folder") for that folder and use it instead.
	if request("initial") then 
	Session("newpath") = Request("newpath") 
	end if
	'Const wexRoot = "/data6/6457"
	Dim wexRoot
	wexroot = Session("newpath")
	
	' Preferred character set
	Const wexCharSet = "ISO-8859-1"
	' Show files and folders that have hidden attribute set?
	Const showHiddenItems = true
	' Calculate total size of the current folder? Disable if it takes long time with huge folders
	Const calculateTotalSize = true
	' Calculate total sizes of the folders in the listing? Disable if it takes long time with huge folders
	Const calculateFolderSize = true
	' List of file extensions which are showed with the "T" icon and can be edited by clicking the icon
	Const editableExtensions = "*htm*|*html*|*asp*|*asa*|*txt*|*inc*|*css*|*aspx*|*js*|*vbs*|*shtm*|*shtml*|*xml*|*xsl*|*log*"
	' List of file extensions which are showed with the "P" icon and can be viewed by clicking the icon
	Const viewableExtensions = "*gif*|*jpg*|*jpeg*|*png*|*bmp*|*jpe*"
	' Set script timeout value to higher values (in seconds) if the script fails when uploading large files
	Server.ScriptTimeout = 300
' ------------------------------------------------------------

	Const appName = "JobExplorer"
	Const appVersion = "1"

	Dim scriptName
	scriptName = Request.ServerVariables("SCRIPT_NAME")
	
	Dim FSO
	Set FSO = server.CreateObject ("Scripting.FileSystemObject")

	Dim wexId
	wexId = appName & appVersion & "-" & FSO.GetParentFolderName(scriptName) & "-"
	
	Dim wexMessage, wexRootPath

	Const iconFolderOpenBig = "<img align=absmiddle border=0 width=32 height=27 src=""./folder_open_big.gif"">"
	Const iconFolderUp = "<img align=absmiddle border=0 width=15 height=13  src=""./folder_up.gif"" alt=""One level up"">"
	Const iconFolder = "<img align=absmiddle border=0 width=15 height=13 src=""./folder.gif"" alt=""Folder - Click to learn details"">"
	Const iconFile = "<img align=absmiddle border=0 width=11 height=14 src=""./file.gif"" alt=""File - Click to learn details"">"
	Const iconFileEditable = "<img align=absmiddle border=0 width=11 height=14 src=""./file_editable.gif"" alt=""Text file - Click to edit and learn details"">"
	Const iconFileViewable = "<img align=absmiddle border=0 width=11 height=14 src=""./file_viewable.gif"" alt=""Picture file - Click to view and learn details"">"
	
	Const iconRefresh = "<img name=""iconRefresh"" align=absmiddle border=0 width=90 height=22 src=""./images/refresh-0.gif"" alt=""Refresh file listing"">"
	Const iconCreateFile = "<img name=""iconCreateFile"" align=absmiddle border=0 width=90 height=22 src=""./images/new_file-0.gif"" alt=""Create new file"">"
	Const iconCreateFolder = "<img name=""iconCreateFolder"" align=absmiddle border=0 width=90 height=22 src=""./images/new_folder-0.gif"" alt=""Create new folder"">"
	Const iconUpload = "<img name=""iconUpload"" align=absmiddle border=0 width=90 height=22 src=""./images/upload-0.gif"" alt=""Upload to this folder"">"
	Const iconLogout = "<img align=absmiddle border=0 width=21 height=20 src=""./logout.gif"" alt=""Logout WebExplorer"">"
	Const iconDelete = "<img align=absmiddle border=0 width=17 height=17 src=""./images/delete2.gif"" alt=""Delete"">"

' - WebExplorer functions ------------------------------------
	' Writes the html header
	Sub HtmlHeader (title, charset)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=<%=charSet%>">
<title><%=request("command")%></title>
<%HtmlStyle%>
<%HtmlJavaScript%>
<link rel="Stylesheet" href="http://testserver1.genergy.com/genergy2_intranet/styles.css" type="text/css">		
<style type="text/css">
<!--
BODY {
SCROLLBAR-FACE-COLOR: #dddddd;
SCROLLBAR-HIGHLIGHT-COLOR: #ffffff;
SCROLLBAR-SHADOW-COLOR: #eeeeee;
SCROLLBAR-3DLIGHT-COLOR: #999999;
SCROLLBAR-ARROW-COLOR: #000000;
SCROLLBAR-TRACK-COLOR: #336699;
SCROLLBAR-DARKSHADOW-COLOR: #333333;
}

td.red {color: red}
-->
</style>
</head>
<body onload="preloadImages();">
<%	
	End Sub
	
	' Writes the html footer
	Sub HtmlFooter ()
%>
</body>
</html>
<%
	End Sub
	
	' Writes the stylesheet
	Sub HtmlStyle
%>
<style>
BODY
{
    BACKGROUND-COLOR: #eeeeee
}
TD
{
    FONT-WEIGHT: normal;
    FONT-SIZE: 10pt;
    COLOR: black;
    FONT-FAMILY: Verdana,Tahoma,Arial,Helvetica
}
.formClass
{
    BACKGROUND-COLOR: #99ccff;
    FONT-WEIGHT: normal;
    FONT-SIZE: 10pt;
    COLOR: black;
    FONT-FAMILY: Verdana,Tahoma,Arial,Helvetica
}
.lightRow {
	BACKGROUND-COLOR: #ffffff
}
.headRow {
	BACKGROUND-COLOR: #eeeeee
}
.darkRow {
	BACKGROUND-COLOR: #ffffff
}
.titleRow {
	BACKGROUND-COLOR: #6699cc
}
.labelRow {
	BACKGROUND-COLOR: #dddddd
}
.loginRow {
	border: black solid 1px;
	BACKGROUND-COLOR: #eeeeee
}
.bevelHighlight {
	border-top: #ffffff solid 1px;
}
.bevelShadow {
	border-bottom: #cccccc solid 1px;
}
.bevel {
	border-top: #ffffff solid 1px;
	border-bottom: #cccccc solid 1px;
}
.boldText
{
    FONT-WEIGHT: bold
}
A
{
    FONT-WEIGHT: bold;
    COLOR: #99ccff;
    TEXT-DECORATION: none
}
A:hover
{
    COLOR: Black;
    TEXT-DECORATION: none
}
A:visited
{
    TEXT-DECORATION: none
}
A:active
{
    COLOR: #6699cc;
    TEXT-DECORATION: none
}
</style>
<%
	End Sub
	
	' Writes the javascript code
	Sub HtmlJavaScript
%>
<script language=javascript>
	function Command(cmd, param) {
		var str;
		var someWin;
		switch (cmd) {
			case "NewFile":
				str = prompt("Please enter a name for the new file", "New File");
				if(!str) return;
				document.forms.formBuffer.parameter.value = str;
				break;
			case "NewFolder":
				str = prompt("Please enter a name for the new folder", "New Folder");
				if(!str) return;
				document.forms.formBuffer.parameter.value = str;
				break;
			case "Edit":
				str = document.forms.formBuffer.folder.value + param;
				someWin = openWin(cmd + str, "", 600, 440, false, false);
				someWin.focus();
				createPage(someWin,cmd,param);
				someWin = null;
				return;
				break;
			case "View":
				if(document.forms.formBuffer.virtual.value=="") {alert("Can not view image without web access!"); return;}
				str = document.forms.formBuffer.folder.value + param;
				someWin = openWin(cmd + str, "", 600, 440, false, true);
				someWin.focus();
				createPage(someWin,cmd,param);
				someWin = null;
				return;
				break;
			case "openNETfolder":
				window.open(param,"JobFolder", "scrollbars=yes, width=500, height=300, resizeable, status" );
				break;
			case "FolderDetails":
				str = document.forms.formBuffer.folder.value + param;
				someWin = openWin(cmd + str, "", 350, 220, false, false);
				someWin.focus();
				createPage(someWin,cmd,param);
				someWin = null;
				return;
				break;
			case "NoDownload":
				alert("Can not download file without web access!");
				return;
				break;
			case "Upload":
				someWin = openWin(cmd, "", 400, 150, true, false);
				someWin.focus(); 
				createPage(someWin,cmd,param);
				someWin = null;
				return;
				break;
			case "DeleteFolder":
				if (!confirm('Are you sure you want to delete the folder "' + param + '" and all its contents ?')) return;
				document.forms.formBuffer.parameter.value = param;
				break;
			case "DeleteFile":
				if (!confirm('Are you sure you want to delete "' + param + '" ?')) return;
				document.forms.formBuffer.parameter.value = param;				
				break;
			default:
				document.forms.formBuffer.parameter.value = param;
		}
		document.forms.formBuffer.target = "";
		document.forms.formBuffer.command.value = cmd
		document.forms.formBuffer.submit();	
	}
	
	function Check() {
		if (document.forms.formBuffer.pwd.value == "") {
			alert("You haven't entered the password!"); 
			return false;
		} else return true;
	}

	function openWin(winName, urlLoc, w, h, showStatus, isViewer) {
		l = (screen.availWidth - w)/2;
		t = (screen.availHeight - h)/2;
		features  = "toolbar=no";      // yes|no 
		features += ",location=no";    // yes|no 
		features += ",directories=no"; // yes|no 
		features += ",status=" + (showStatus?"yes":"no");  // yes|no 
		features += ",menubar=no";     // yes|no 
		features += ",scrollbars=" + (isViewer?"yes":"no");   // auto|yes|no 
		features += ",resizable=" + (isViewer?"yes":"no");   // yes|no 
		features += ",dependent";      // close the parent, close the popup, omit if you want otherwise 
		features += ",height=" + h;
		features += ",width=" + w;
		features += ",left=" + l;
		features += ",top=" + t;
		winName = winName.replace(/[^a-z]/gi,"_");
		return window.open(urlLoc,winName,features);
	} 
	
	function createPage (theWin, cmd, param){
		document.forms.formBuffer.target = theWin.name;
		document.forms.formBuffer.command.value = cmd;
		document.forms.formBuffer.parameter.value = param;
		document.forms.formBuffer.submit();
	}

	function EditorCommand (cmd) {
		switch (cmd) {
			case "Info":
				alert(document.forms.formBuffer.info.value.replace(/\|/gi,"\n"));
				break;
			case "Reload":
				document.forms.formBuffer.reset();
				break;
			case "Save":
				document.forms.formBuffer.subcommand.value = "Save";
				document.forms.formBuffer.submit();
				break;
			case "SaveAs":
				var str, oldname;
				oldname = document.forms.formBuffer.parameter.value;
				str = prompt("Save as the file :", oldname);
				if (!str || str==oldname) return;
				document.forms.formBuffer.parameter.value = str;
				document.forms.formBuffer.subcommand.value = "SaveAs";
				document.forms.formBuffer.submit();
				break;
		}
	}

	function ViewerCommand (cmd) {
		switch (cmd) {
			case "Info":
				alert(document.forms.formBuffer.info.value.replace(/\|/gi,"\n"));
				break;
			case "Reload":
				document.forms.formBuffer.submit();
				break;
		}
	}

	function Upload() {
		document.forms.formBuffer.submit();
	}
	
	//button mouseover functions
	
	var loaded = 0; //images aren't preloaded yet
	function preloadImages(){
	  iconRefreshOn = new Image(); iconRefreshOn.src = "images/refresh-1.gif";
	  iconRefreshOff = new Image(); iconRefreshOff.src = "images/refresh-0.gif";
	  iconCreateFileOn = new Image(); iconCreateFileOn.src = "images/new_file-1.gif";
	  iconCreateFileOff = new Image(); iconCreateFileOff.src = "images/new_file-0.gif";
	  iconCreateFolderOn = new Image(); iconCreateFolderOn.src = "images/new_folder-1.gif";
	  iconCreateFolderOff = new Image(); iconCreateFolderOff.src = "images/new_folder-0.gif";
	  iconUploadOn = new Image(); iconUploadOn.src = "images/upload-1.gif";
	  iconUploadOff = new Image(); iconUploadOff.src = "images/upload-0.gif";
	  loaded = 1;
	}
	
	function msover(img){
	  if ((loaded) && (document.all)) {
	    document.all[img].src = eval (img + "On.src");
	  }
	}
	
	function msout(img){
	  if ((loaded) && (document.all)) {
	    document.all[img].src = eval (img + "Off.src");
	  }
	}
	
	//display quickhelp
	var helpIsOn = 0;
	function toggleHelp(){
	  if (helpIsOn) { 
	    document.all.quickhelptext.style.display='none';
	    helpIsOn = 0;
     } else { 
      document.all.quickhelptext.style.display='inline';
      helpIsOn = 1;
     }
	}
</script>
<%
	End Sub

	' Write file listing of the given folder
	Sub WriteListing (byref folder)
		Dim item, arr
		Dim rowType
		Dim listed
		
		on error resume next
		
%>
<form name=formGlobal action="noaction" onSubmit="return(false);">
<table cellspacing=0 cellpadding=3 border=0 width="100%">
	<tr class="titleRow">
		<td class="titleRow" align=left>&nbsp;<span class="standardheader"><%=appName%> v<%=appVersion%> - <%=Request.ServerVariables("SERVER_NAME")%></span></td>
		<td class="titleRow" align=right>&nbsp;</td>
	</tr>
</table>
<table cellspacing=0 cellpadding=2 border=0 width=100%>
	<tr class=headRow height=60>
		<td class="bevelShadow">
			<div style="font-size:12pt;">&nbsp;<%=iconFolderOpenBig%>&nbsp;<font class=boldText><%=folder.Name%></font></div><br><a href="javascript:Command('openNETfolder','<%="\\\\10.0.7.2\\genergy\\operations\\operations_log" & RealizeDocPath(VirtualPath(destFolder))%>')" ><img src="./images/arrow_rt.gif" width="10" height="12" align="absmiddle" border="0">&nbsp;Open this folder for document editing</a>
		</td>
		<td class="bevelShadow" nowrap>
			<font class=boldText><%=folder.subfolders.count%></font> subfolder(s)<br>
			<font class=boldText><%=folder.files.count%></font> file(s)
		</td>
		<td class="bevelShadow" nowrap>
			Total Size: <font class=boldText><%If err.Number<>0 or (not calculateTotalSize) Then Response.Write "N/A" Else Response.Write FormatSize(folder.size)%></font>
		</td>
		<td class="bevelShadow" align="right">
    <table border=0 cellpadding="1" cellspacing="0">
    <tr>
      <td><a href="javascript:Command('Refresh', '');" onmouseover="msover('iconRefresh');" onmouseout="msout('iconRefresh');"><%=iconRefresh%></a></td>
      <td><a href="javascript:Command('NewFile', '');" onmouseover="msover('iconCreateFile');" onmouseout="msout('iconCreateFile');"><%=iconCreateFile%></a></td>
      <td><a href="javascript:Command('NewFolder', '');" onmouseover="msover('iconCreateFolder');" onmouseout="msout('iconCreateFolder');"><%=iconCreateFolder%></a></td>
      <td><a href="javascript:Command('Upload', '');" onmouseover="msover('iconUpload');" onmouseout="msout('iconUpload');"><%=iconUpload%></a></td>
    </tr>
    </table>		
			
<%If wexPassword <> "" Then%>
			<a href="javascript:Command('Logout', '');"><%=iconLogout%></a>
<%End If%>
		</td>
		<td class="bevelShadow">&nbsp;</td>
	</tr>
	<tr>
	  <td colspan="3" class="bevel">&nbsp;<span id="wexMessage"><%=server.HTMLEncode(wexMessage)%></span></td>
		<td align="right" class="bevel"><img src="images/quick_help.gif" align="absmiddle" alt="?" width="19" height="19" border="0"><a href="javascript:toggleHelp();"><span class=boldText>Quick Help</span></a></td>
		<td class="bevel">&nbsp;</td>
	</tr>
</table>
<div id="quickhelptext" style="display:none;">
<table border=0 cellpadding="3" cellspacing="0">
<tr>
  <td><br>
  <ul>
  <li>Click a filename to view the file; click a folder name to open the folder
  <li>To edit files, navigate to the folder that contains the files you want to edit, then click &quot;<b><a href="javascript:Command('openNETfolder','<%="\\\\10.0.7.2\\genergy\\operations\\operations_log" & RealizeDocPath(VirtualPath(destFolder))%>')">Open this folder for document editing</a></b>&quot; underneath the large folder icon (above this help section). <br><span style="color:#990033;">Note that direct document editing is only available when connected to the Genergy in-office network or via VPN. You may see an error otherwise.</span>
  </ul>
  </td>
</tr>
</table>
</div>
<table cellspacing=1 cellpadding=2 border=0 width=100%>
	<tr class=labelRow>
		<td>&nbsp;<font class=boldText>Name</font></td>
		<td>&nbsp;<font class=boldText>Size</font></td>
		<td>&nbsp;<font class=boldText>Type</font></td>
		<td>&nbsp;<font class=boldText>Modified</font></td>
		<td width="86">&nbsp;<font class=boldText>Delete</font></td>
	</tr>
<%
	rowType = "darkRow"

	If len(destFolder) > len(wexRootPath) Then
%>
	<tr class=<%=rowType%>><td colpsna="2">&nbsp;<a href="javascript:Command('LevelUp','')"><%=iconFolderUp%></a>&nbsp;<a href="javascript:Command('LevelUp','')">..Up one directory</a></td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr>
<%
		rowType = "lightRow"
	End If
	
	listed = 0
	If (folder.subfolders.Count + folder.files.Count) = 0 Then
		' Do nothing when error occurs
	Else
		For each item in folder.subfolders
			If showHiddenItems or not item.Attributes and 2 Then
				listed = listed + 1
%>
	<tr class=<%=rowType%>><td>&nbsp;<%=GetIcon(item.Name, true)%>&nbsp;<a href="javascript:Command('OpenFolder','<%=item.Name%>')"><%=item.Name%></a></td><td>&nbsp;<%If calculateFolderSize Then Response.write FormatSize(item.Size)%></td><td>&nbsp;<%=item.Type%></td><td nowrap>&nbsp;<%=item.DateLastModified%></td><td>&nbsp;<a href="javascript:Command('DeleteFolder','<%=item.Name%>')"><%=iconDelete%></a></td></tr>
<%
				If rowType = "darkRow" Then rowType = "lightRow" Else rowType = "darkRow"
			End If
		Next

		For each item in folder.files
			If showHiddenItems or not item.Attributes and 2 Then
				listed = listed + 1
%>
	<tr class=<%=rowType%>><td>&nbsp;<%=GetIcon(item.Name, false)%>&nbsp;<a href="<%If VirtualPath(destFolder)<>"" Then Response.write VirtualPath(destFolder) & item.Name Else Response.Write "javascript:Command('NoDownload')"%>"><%=item.Name%></a></td><td>&nbsp;<%=FormatSize(item.Size)%></td><td>&nbsp;<%=item.Type%></td><td nowrap>&nbsp;<%=item.DateLastModified%></td><td>&nbsp;<a href="javascript:Command('DeleteFile','<%=item.Name%>')"><%=iconDelete%></a></td></tr>
<%
				If rowType = "darkRow" Then rowType = "lightRow" Else rowType = "darkRow"
			End If	
		Next
	End If
%>
	<tr></tr>
</table>
<table cellspacing=0 cellpadding=0 border=0 width=100%>
	<tr class=titleRow>
		<td></td>
	</tr>
</table>

</form>
<%
		If wexMessage="" Then 
			If (folder.subfolders.Count + folder.files.Count) <> listed Then
				wexMessage = "Listed " & listed & " of " & (folder.subfolders.Count + folder.files.Count) & " item(s) , " & (folder.subfolders.Count + folder.files.Count - listed) & " item(s) are hidden..."
			Else
				wexMessage = "Listed " & (folder.subfolders.Count + folder.files.Count) & " item(s)..."
			End If
			Response.Write "<script language=""javascript"">document.all.wexMessage.innerHTML='" & wexMessage & "'</script>"
		End If
	End Sub

	' WebExplorer Login screen
	Sub	Login
		If Request.Form("command") = "Login" Then
			If Request.Form("pwd") = wexPassword Then
				Session(wexId & "Login") = true
				Exit Sub
			Else
				wexMessage = "Wrong password!"
			End If
		End If
		
		HtmlHeader appName, wexCharSet
		If(wexMessage<>"") Then Response.Write "<script language=""javascript"">alert('" & wexMessage & "');</script>"
%>
<form name=formBuffer method=post action="<%=scriptName%>" onSubmit="javascript:return(Check());">
<table border=0 cellspacing=0 cellpadding=0 width=400 align=center>
	<tr><td><br><br><br></td></tr>
	<tr><td>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=left>
					&nbsp;<font class=boldText>Login</font>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr align=center class=lightRow>
				<td>
					<br>
					<font class=boldText>Welcome to <%=appName%> v<%=appVersion%></font>
					<br><br>
					<table cellspacing=0 cellpadding=5 border=0 class=loginRow>
						<tr>
							<td align=left>&nbsp;<font class=boldText>Password</font></td>
						</tr>
						<tr>
							<td align=center><input type="password" class=formClass name=pwd value="" size=21></td>
						</tr>
						<tr>
							<td align=right><input type=submit name=submitter value="Login" class=formClass></td>
						</tr>
					</table>
					<br><br><br>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=center>&nbsp;</td>
			</tr>
		</table>
	</td></tr>
</table>
<input type="hidden" name=command value="Login">


</form>
<script language="javascript">document.forms.formBuffer.pwd.focus();</script>
<%	
		HtmlFooter
		Response.End 
	End Sub
	
	' Relogin message for the pop windows
	Sub PopupRelogin
		Response.Write "<html><head>"
		Response.Write "<title>WebExplorer - Message</title>"
		Response.Write "<style>A{FONT-WEIGHT: bold; COLOR: #99ccff; TEXT-DECORATION: none}"
		Response.write "A:hover{COLOR: #ccffff; TEXT-DECORATION: none }"
		Response.Write "A:visited{TEXT-DECORATION: none}"
		Response.write "A:active {COLOR: #ccffff; TEXT-DECORATION: none}</style>"
		Response.Write "</head><body style=""BACKGROUND-COLOR: #003366"">"
		Response.Write "<div style=""COLOR: white; FONT-FAMILY: Verdana,Tahoma,Arial,Helvetica; FONT-SIZE: 10pt; FONT-WEIGHT: bold;"">"
		Response.Write appName & " session is destroyed, please "
		Response.write "<a href=""javascript:opener.Command('Refresh'); window.close();"">relogin</a>."
		Response.Write "</div></body></html>"
		Response.End 
	End Sub
	
	' Checks if there is a valid login
	Function Secure()
		If wexPassword = "" Then
			Secure = true
		Else
			If Session(wexId & "Login") Then Secure = true Else Secure = false
		End If
	End Function
		
	' Logouts WebExplorer session
	Sub Logout
		Session(wexId & "Login") = false
		Login
	End Sub
		
	' Returns the icon of the file
	Function GetIcon(fileName, isFolder)
		Dim ext
		If isFolder Then
			GetIcon = iconFolder
		Else
			ext = FSO.GetExtensionName(fileName)
			If InStr(1,editableExtensions, "*" & ext & "*", 1) <> 0 Then
				GetIcon = "<a href=""javascript:Command('Edit', '" & fileName & "');"">" & iconFileEditable & "</a>"
			ElseIf InStr(1,viewableExtensions, "*" & ext & "*", 1) <> 0 Then
				GetIcon = "<a href=""javascript:Command('View', '" & fileName & "');"">" & iconFileViewable & "</a>"
			Else
				GetIcon = iconFile
			End If
		End If
	End Function
	
	' Formats given size in bytes,KB,MB and GB
	Function FormatSize (givenSize)
		If (givenSize < 1024) Then
			FormatSize = givenSize & " B"
		ElseIf (givenSize < 1024*1024) Then
			FormatSize = FormatNumber(givenSize/1024,2) & " KB"
		ElseIf (givenSize < 1024*1024*1024) Then
			FormatSize = FormatNumber(givenSize/(1024*1024),2) & " MB"
		Else
			FormatSize = FormatNumber(givenSize/(1024*1024*1024),2) & " GB"
		End If
	End Function

	' Adds given type of the slash to the end of the path if required
	Function FixPath(path, slash)
		If Right(path, 1) <> slash Then
            FixPath = path & slash
        Else
			FixPath = path
        End If
	End Function

	' Converts the given path to physical path
	Function RealizePath(thePath)
		Dim path
		path = replace(thePath,"/","\")
		If left(path,1) = "\" Then
			on error resume next
			RealizePath = FixPath(server.MapPath(path),"\")
			If err.Number<>0 Then RealizePath = thePath
		Else
			If InStr(1,path, ":", 1) <> 0 Then
				RealizePath = FixPath(path,"\")
			Else
				RealizePath = thePath & "?"
			End If
		End If
	End Function	
	Function RealizeDocPath(thePath)
	
		Dim path
		path = replace(thePath,"/","\\")

		RealizeDocPath=path

	End Function	
	' Converts the given path to virtual path
	Function VirtualPath(thePath)
		Dim webRoot, path
		webRoot = FixPath(server.MapPath("/"),"\")
		path = FixPath(thePath,"\")
		VirtualPath = ""
		If left(wexRoot,1) = "/" Then
			VirtualPath = FixPath(wexRoot, "/")
			VirtualPath = VirtualPath & right(path, len(path) - len(wexRootPath))
			VirtualPath = replace(VirtualPath, "\", "/")
			VirtualPath = FixPath(VirtualPath,"/")
		ElseIf left(lcase(path), len(webRoot)) = lcase(webRoot) Then
			VirtualPath = "/" & right(path, len(path) - len(webRoot))
			VirtualPath = replace(VirtualPath, "\", "/")
			VirtualPath = FixPath(VirtualPath,"/")
		End If
	End Function
	
	' Makes sure that given file name does not contain path info
	Function SecureFileName(name)
		SecureFileName = replace(name,"/","?")
		SecureFileName = replace(SecureFileName,"\","?")
	End Function

	' Creates a folder or a file
	Function CreateItem()
		Dim itemType, itemName, itemPath
		itemType = Request.Form("command")
		itemName = SecureFileName(Request.Form("parameter"))
		itemPath = destFolder & itemName

		on error resume next
		
		Select Case itemType
			Case "NewFolder"
				If (FSO.FolderExists(itemPath) = false and FSO.FileExists(itemPath) = false ) Then 
					FSO.CreateFolder(itemPath)
					If (err.Number <> 0 ) Then 
						CreateItem = "Unable to create the folder """ & itemName & """, an error occured..." 
					Else
						CreateItem = "Created the folder """ & itemName & """..."
					End If
				Else
					CreateItem = "Unable to create the folder """ & itemName & """, there exists a file or a folder with the same name..."
				End If
			Case "NewFile"
				If (FSO.FolderExists(itemPath) = false and FSO.FileExists(itemPath) = false ) Then 
					FSO.CreateTextFile(itemPath)
					If (err.Number <> 0 ) Then 
						CreateItem = "Unable to create the file """ & itemName & """, an error occured..."
					Else
						CreateItem = "Created the file """ & itemName & """..."
					End If
				Else 
					CreateItem = "Unable to create the file """ & itemName & """, there exists a file or a folder with the same name..."
				End IF
		End Select
	End Function
	
	' Deletes a folder or a file
	Function DeleteItem
		Dim itemType, itemName, itemPath
		itemType = Request.Form("command")
		itemName = SecureFileName(Request.Form("parameter"))
		itemPath = destFolder & itemName

		on error resume next
		
		Select Case itemType
			Case "DeleteFolder"
				FSO.DeleteFolder itemPath, true
				If (err.Number <> 0 ) Then 
					DeleteItem = "Unable to delete the folder """ & itemName & """, an error occured..." 
				Else
					DeleteItem = "Deleted the folder """ & itemName & """..."
				End If
			Case "DeleteFile"
				FSO.DeleteFile itemPath, true
				If (err.Number <> 0 ) Then 
					DeleteItem = "Unable to delete the file """ & itemName & """, an error occured..." 
				Else
					DeleteItem = "Deleted the file """ & itemName & """..."
				End If
		End Select
	End Function
	
	' WebExplorer Editor
	Sub Editor
		Dim fileName, filePath, file
		
		If not Secure() Then PopupRelogin

		on error resume next

		Select Case Request.Form("subcommand")
			Case "Save", "SaveAs"
				fileName = SecureFileName(Request.Form("parameter"))
				filePath = destFolder & fileName
				Set file = FSO.OpenTextFile (filePath,2,true,0)
				If (err.Number<>0) Then 
					wexMessage = "Can not write to the file """ & fileName & """, permission denied!"
					err.Clear
				Else
					file.write Request.Form("content")
				End If
				Set file = Nothing
				Set file = FSO.OpenTextFile (filePath,1,false,0)
			Case Else
				fileName = SecureFileName(Request.Form("parameter"))
				filePath = destFolder & fileName
				
				If not FSO.FileExists(filePath) Then
					wexMessage = "The file """ & fileName & """ does not exist"
					Set file = FSO.CreateTextFile (filePath, false, false)
					If err.Number<>0 Then 
						wexMessage = wexMessage & ", also unable to create new file."
						err.Clear 
					Else
						wexMessage = wexMessage & ", created new file."
					End If
				Else
					Set file = FSO.OpenTextFile (filePath,1,false,0)
					If err.Number<>0 Then 
						wexMessage = "Can not read from the file """ & fileName & """, permission denied!"
						err.Clear 
					End If
				End If
		End Select
		
		HtmlHeader appName, wexCharSet
		If(wexMessage<>"") Then Response.Write "<script language=""javascript"">alert('" & wexMessage & "');</script>"
%>
<form name=formBuffer method=post action="<%=scriptName%>">
		<table border=0 cellspacing=0 cellpadding=3 width=100%>
			<tr class=titleRow>
				<td align=left>
					&nbsp;<font class=boldText>Editing</font> - <%=fileName%>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100% height=90%>
			<tr align=center class=lightRow>
				<td valign=middle>
<textarea name=content class=formClass rows=22 cols=46 style="width:580; height:370;" wrap="off">
<%=Server.HTMLEncode(file.ReadAll)%></textarea>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=center>					
					<a href="javascript:EditorCommand('Save');">Save</a> | <a href="javascript:EditorCommand('SaveAs');">Save As</a> | <a href="javascript:EditorCommand('Reload');">Reload</a> | <a href="javascript:EditorCommand('Info');">Info</a> | <a href="javascript:this.close();">Close</a>
				</td>
			</tr>
		</table>
<%
		Set file = Nothing
		Set file = FSO.GetFile (filePath)
%>
<input type="hidden" name=command value="Edit">
<input type="hidden" name=subcommand value="">
<input type="hidden" name=parameter value="<%=fileName%>">
<input type="hidden" name=folder value="<%=Request.Form("folder")%>">
<input type="hidden" name=info value="Size: <%=FormatSize(file.Size)%>|Type: <%=file.Type%>|Created: <%=file.DateCreated%>|Last Accessed: <%=file.DateLastAccessed%>|Last Modified: <%=file.DateLastModified%>">

</form>
<%
		Set file = Nothing
		HtmlFooter
		Response.End 
	End Sub

	' WebExplorer Viewer
	Sub Viewer
		Dim fileName, filePath, file, imageSrc

		If not Secure() Then PopupRelogin

		on error resume next
		fileName = Request.Form("parameter")
		filePath = destFolder & fileName
		imageSrc = replace(VirtualPath(destFolder) & fileName, " ", "%20")

		HtmlHeader appName, wexCharSet
%>
<form name=formBuffer method=post action="<%=scriptName%>">
		<table border=0 cellspacing=0 cellpadding=3 width=100%>
			<tr class=titleRow>
				<td align=left>
					&nbsp;<font class=boldText>Viewing</font> - <%=fileName%>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100% height=90%>
			<tr align=center class=lightRow>
				<td valign=middle>
					<img src="<%=imageSrc%>">
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=center>					
					<a href="javascript:ViewerCommand('Reload');">Reload</a> | <a href="javascript:ViewerCommand('Info');">Info</a> | <a href="javascript:this.close();">Close</a>
				</td>
			</tr>
		</table>
<%
		Set file = FSO.GetFile (filePath)
%>
<input type="hidden" name=command value="View">
<input type="hidden" name=subcommand value="Refresh">
<input type="hidden" name=parameter value="<%=fileName%>">
<input type="hidden" name=folder value="<%=Request.Form("folder")%>">
<input type="hidden" name=info value="Size: <%=FormatSize(file.Size)%>|Type: <%=file.Type%>|Created: <%=file.DateCreated%>|Last Accessed: <%=file.DateLastAccessed%>|Last Modified: <%=file.DateLastModified%>">

</form>
<%
		Set file = Nothing
		HtmlFooter
		Response.End 
	End Sub

	' File/Folder Details
	Sub Details
		Dim fileName, filePath, file
		
		If not Secure() Then PopupRelogin
		
		on error resume next
		fileName = Request.Form("parameter")
		filePath = destFolder & fileName

		HtmlHeader appName, wexCharSet
%>
<form name=formBuffer method=post action="<%=scriptName%>">
		<table border=0 cellspacing=0 cellpadding=3 width=100%>
			<tr class=titleRow>
				<td align=left>
					&nbsp;<font class=boldText>Details</font> - <%=fileName%>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100% height=80%>
			<tr align=center class=lightRow>
				<td valign=middle>
<%
		If Request.Form("command") = "FileDetails" Then
				Set file = FSO.GetFile (filePath)
		Else
				Set file = FSO.GetFolder (filePath)
		End If
%>
				<table border=0 cellspacing=5 cellpadding=0>
					<tr><td><font class=boldText>Size:</font></td><td><%=FormatSize(file.Size)%></td></tr>
					<tr><td><font class=boldText>Type:</font></td><td><%=file.Type%></td></tr>
					<tr><td><font class=boldText>Created:</font></td><td><%=file.DateCreated%></td></tr>
					<tr><td><font class=boldText>Last Accessed:</font></td><td><%=file.DateLastAccessed%></td></tr>
					<tr><td><font class=boldText>Last Modified:</font></td><td><%=file.DateLastModified%></td></tr>
				</table>
<%
		Set file = Nothing
%>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=center>					
					<a href="javascript:this.close();">Close</a>
				</td>
			</tr>
		</table>
<input type="hidden" name=command value="<%=Request.Form("command")%>">
<input type="hidden" name=parameter value="<%=fileName%>">
<input type="hidden" name=folder value="<%=Request.Form("folder")%>">

</form>
<%
		HtmlFooter
		Response.End 
	End Sub
	
	' Uploads a file
	Sub Upload
		If not Secure() Then PopupRelogin

		If Request.QueryString("command")="DoUpload" Then 
			destFolder = wexRootPath & Request.QueryString("folder")
			destFolder = FixPath(destFolder, "\")
			'response.write destfolder
			'response.end
			If len(destFolder) < len(wexRootPath) Then Response.End 
		End If
		
		HtmlHeader appName, wexCharSet
%>
		<table border=0 cellspacing=0 cellpadding=3 width=100%>
			<tr class=titleRow>
				<td align=left>
					&nbsp;<font class=boldText>Upload</font> - <%=FSO.GetBaseName(destFolder)%>
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100% height=80%>
			<tr align=center class=lightRow>
				<td valign=middle>
					<font class=boldText>
<%
		If Request.QueryString("command")="DoUpload" Then
	
			Dim Uploader
			Set Uploader = New classUploader
	
			Uploader.Upload()
	
			If not Uploader.uploaded Then
				Response.Write "No file sent"
			Else
				If Uploader.uploadedFile.FileName="" Then
					Response.Write "No file sent"
				Else
					If Uploader.uploadedFile.Save(destFolder) Then
						Response.Write Uploader.uploadedFile.FileName & " is uploaded<br>"
						Response.Write FormatSize(Uploader.uploadedFile.FileSize) & " (" & Uploader.uploadedFile.FileSize & " bytes) written<br>"
						Response.Write "<script language=""javascript"">opener.Command('Refresh');</script>"
					Else
						Response.Write Uploader.uploadedFile.FileName & " can not be written<br>"
					End If
				End If
			End If
%>
					</font>
					<form name=formBuffer method=post action="<%=scriptName%>">
						<input type=hidden name=command value="Upload">
						<input type=hidden name=folder value="<%=Request.QueryString("folder")%>">
						<input type=hidden name=newpath value="<%=Request.QueryString("newpath")%>">
					</form>
<%	
		Else
%>
					<form enctype="multipart/form-data" name=formBuffer method=post action="<%=scriptName%>?command=DoUpload&folder=<%=server.URLEncode(Request.Form("folder"))%>">
						<input type=file name=file class=formClass>
						
					</form>
<%
		End If
%>					
				</td>
			</tr>
		</table>
		<table border=0 cellspacing=0 cellpadding=0 width=100%>
			<tr class=titleRow>
				<td align=center>					
					<a href="javascript:Upload();">Upload</a> | <a href="javascript:this.close();">Close</a>
				</td>
			</tr>
		</table>
<%
		HtmlFooter
		Response.End 
	End Sub

	' Class containing upload data parsing functions
	Class classUploader
		Public uploaded
		Public uploadedFile
	
		Public Default Sub Upload()
			Dim biData, sInputName
			Dim nPosBegin, nPosEnd, nPos, vDataBounds, nDataBoundPos
			Dim nPosFile, nPosBound
	
			biData = Request.BinaryRead(Request.TotalBytes)
			nPosBegin = 1
			nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
			
			If (nPosEnd-nPosBegin) <= 0 Then Exit Sub
			 
			vDataBounds = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
			nDataBoundPos = InstrB(1, biData, vDataBounds)
			
			Do Until nDataBoundPos = InstrB(biData, vDataBounds & CByteString("--"))
				
				nPos = InstrB(nDataBoundPos, biData, CByteString("Content-Disposition"))
				nPos = InstrB(nPos, biData, CByteString("name="))
				nPosBegin = nPos + 6
				nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(34)))
				sInputName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
				nPosFile = InstrB(nDataBoundPos, biData, CByteString("filename="))
				nPosBound = InstrB(nPosEnd, biData, vDataBounds)
				
				If nPosFile <> 0 And  nPosFile < nPosBound Then
					Dim sFileName
					Set uploadedFile = New classUploadedFile
					
					nPosBegin = nPosFile + 10
					nPosEnd =  InstrB(nPosBegin, biData, CByteString(Chr(34)))
					sFileName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
					uploadedFile.FileName = Right(sFileName, Len(sFileName)-InStrRev(sFileName, "\"))
	
					nPos = InstrB(nPosEnd, biData, CByteString("Content-Type:"))
					nPosBegin = nPos + 14
					nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
					
					nPosBegin = nPosEnd+4
					nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
					uploadedFile.FileData = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
					
					If uploadedFile.FileSize > 0 Then uploaded = true Else uploaded = false
				End If

				nDataBoundPos = InstrB(nDataBoundPos + LenB(vDataBounds), biData, vDataBounds)
			Loop
		End Sub

		' String to byte string conversion
		Private Function CByteString(sString)
			Dim index
			For index = 1 to Len(sString)
				CByteString = CByteString & ChrB(AscB(Mid(sString,index,1)))
			Next
		End Function

		' Byte string to string conversion
		Private Function CWideString(bsString)
			Dim index
			CWideString =""
			For index = 1 to LenB(bsString)
				CWideString = CWideString & Chr(AscB(MidB(bsString,index,1))) 
			Next
		End Function
	End Class

	' Class containing file data writing functions
	Class classUploadedFile
		Public FileName, FileData
		
		Public Property Get FileSize()
			FileSize = LenB(FileData)
		End Property
	
		Public Function Save(path)
			Dim file
			Dim index
		
			If path = "" Or FileName = "" Then 
				Save = false
				Exit Function
			End If
			path = FixPath(path, "\")
		
			If Not FSO.FolderExists(path) Then
				Save = false
				Exit Function
			End If
		
			on error resume next
			Set file = FSO.CreateTextFile(path & FileName, True)
			
			For index = 1 to LenB(FileData)
			    file.Write Chr(AscB(MidB(FileData,index,1)))
			Next
	
			file.Close
			If err.Number<>0 Then
				Save = false
			Else
				Save = true
			End If
		End Function
	End Class
' ------------------------------------------------------------

' - WebExplorer Main -----------------------------------------
	Dim folder, destFolder

	wexRootPath = RealizePath(wexRoot)
	If Request.QueryString("command")="DoUpload" Then Upload()
	
	destFolder = wexRootPath & Request.Form("folder")
	destFolder = FixPath(destFolder, "\")
	If len(destFolder) < len(wexRootPath)  Then Response.End 
	
	' Actions in the popup windows
	Select Case Request.Form("command")
		Case "Edit"
			Editor()
		Case "View"
			Viewer()
		Case "FileDetails", "FolderDetails"
			Details()
		Case "Upload"
			Upload()
	End Select
	
	' Actions in the main window
	If not Secure() Then Login
	
	Select Case Request.Form("command")
		Case "NewFile", "NewFolder"
			wexMessage = CreateItem()
		Case "Logout"
			Logout()
		Case "DeleteFile", "DeleteFolder"
			wexMessage = DeleteItem()
		Case "OpenFolder"
			If Request.Form("folder") = "" Then
				destFolder = wexRootPath & Request.Form("parameter")
			Else	
				destFolder = wexRootPath & FixPath(Request.Form("folder"),"\") & Request.Form("parameter")
			End If
			destFolder = FixPath(destFolder, "\")
			
			If len(destFolder) < len(wexRootPath) Then Response.End 
		Case "LevelUp"
			destFolder = FSO.GetParentFolderName(destFolder)
			destFolder = FixPath(destFolder, "\")
			If len(destFolder) < len(wexRootPath) Then Response.End 
	End Select

	on error resume next
	Set folder = FSO.GetFolder(destFolder)
	if err.Number<>0 Then wexMessage = "Error opening folder """ & destFolder & """. Your browser session may have expired. Please quit your browser and log back in after a minute or two."

	HtmlHeader appName, wexCharSet
	WriteListing (folder)
%>
<form method=post action="<%=scriptName%>" name=formBuffer>
	<input type=hidden name=command value="">
	<input type=hidden name=parameter value="">
	<input type=hidden name=virtual value="<%=VirtualPath(destFolder)%>">
	<input type=hidden name=folder value="<%=right(destFolder, len(destFolder)-len(wexRootPath))%>">
	
</form>
<%
	Set folder = Nothing
	HtmlCopyright
	HtmlFooter
%>