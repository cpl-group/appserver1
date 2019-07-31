<%
	'Set this to the home directory in relationship to the script.
	'You cannot have the directory set to up one directory (ie "../") for security reasons unless you disable "Security" below.
	'Note that if your file home directory is up one level the script may not function properly.
	HomeDir = "/um"
	'Do not put a "/" on the end of this!
	'Do not edit below this line unless you know what you're doing.
	Security = "On"
	'If you do not want users browsing directories up one level from the specified one, set this to "On".
	'Otherwise, set it to "Off"
	CurrentDir = Request.QueryString("dir")
	if Left(CurrentDir, 3) = "../" and Security = "On" then
	Response.Write "You are not permitted to browse this directory."
	else
	Set FSO = Server.CreateObject("Scripting.FileSystemObject")
	if CurrentDir = "" then
	CurrentDir = HomeDir
	Set BrowseContents = FSO.GetFolder(Server.MapPath(HomeDir))
	else
	Set BrowseContents = FSO.GetFolder(Server.MapPath(CurrentDir))
	end if
%>
<HTML>
<HEAD>
<TITLE>ArgoBrowse</TITLE>
<LINK REL=STYLESHEET TYPE=TEXT/CSS HREF="styles.css">
</HEAD>
<BODY><DIV ALIGN=CENTER>
<TABLE WIDTH=80%> <TR><TD>
<TABLE WIDTH=100% BGCOLOR=#3399CC BORDERCOLOR=#999999 BORDER=1 CELLSPACING=0>
  <TR><TD>
  <TABLE WIDTH=100%>
    <TR>
      <TD COLSPAN=2><H3>Contents of <%=CurrentDir%>:</TD>
    </TR>
    <TR> 
      <TD WIDTH="75%">Available Documents:</TD>
      <TD WIDTH="25%">
        <DIV ALIGN=RIGHT>Item size:</DIV>
      </TD>
    </TR>
    <%
	if not CurrentDir = HomeDir then
	UpOneDirBegPos = InStrRev(CurrentDir, "/")
	UpOneDirTotal = len(CurrentDir)
	UpOneDirPath = Left(CurrentDir, UpOneDirBegPos - 1)
	Response.Write "<TABLE WIDTH=""100%"" BORDER=""0""><TR>"
	Response.Write "<TD WIDTH=""*""><IMG SRC=""images/foldericon.gif""></TD><TD WIDTH=""70%""><A HREF=""?dir=" & UpOneDirPath & """>(up one level)</A><BR>"
	Response.Write "</TD><TD WIDTH=""25%""></TD></TR></TABLE>"
	end if
	
	for each BrowsedFolder in BrowseContents.SubFolders
	if BrowsedFolder.Name <> "_vti_cnf" then
	Response.Write "<TABLE WIDTH=""100%"" BORDER=""0""><TR>"
	Response.Write "<TD WIDTH=""*""><IMG SRC=""images/foldericon.gif""></TD><TD WIDTH=""70%""><A HREF=""?dir=" & CurrentDir & "/" & BrowsedFolder.Name & """>" & BrowsedFolder.Name & "</A><BR>"
	Response.Write "</TD><TD WIDTH=""25%""><DIV ALIGN=""RIGHT"">"
	Response.Write BrowsedFolder.Size \ 1024 & "KB"
	Response.Write "</DIV></TD></TR></TABLE>"
	ContainsSomething = "True"
	end if
	next
	
	for each BrowsedFile in BrowseContents.Files
	FileExtBegPos = InStrRev(BrowsedFile.Name, ".")
	FileExtTotal = len(BrowsedFile.Name)
	FileExt = Right(BrowsedFile.Name, FileExtTotal - FileExtBegPos)
	
	select case FileExt
	case "asp"
	FileIcon = "aspicon.gif"
	case "htm", "html", "shtm", "shtml"
	FileIcon = "htmlicon.gif"
	case "jpg", "jpeg", "gif", "png", "bmp"
	FileIcon = "imgicon.gif"
	case "zip", "rar", "tar", "gz"
	FileIcon = "zipicon.gif"
	case "exe"
	FileIcon = "exeicon.gif"
	case "mp3", "wav", "rm", "wmv"
	FileIcon = "mp3icon.gif"
	case else
	FileIcon = "icon.gif"
	end select
	
	Response.Write "<TABLE WIDTH=""100%"" BORDER=""0""><TR>"
	Response.Write "<TD WIDTH=""*""><IMG SRC=""images/" & FileIcon & """></TD><TD WIDTH=""70%""> <A HREF=""" & CurrentDir & "/" & BrowsedFile.Name & """>" & BrowsedFile.Name & "</A><BR>"
	Response.Write "</TD><TD WIDTH=""25%""><DIV ALIGN=""RIGHT"">"
	Response.Write BrowsedFile.Size \ 1024 & "KB"
	Response.Write "</DIV></TD></TR></TABLE>"
	ContainsSomething = "True"
	next
	
	if ContainsSomething <> "True" then
	Response.Write "<TABLE WIDTH=""100%""><TR><TD>There is nothing in this directory.</TD></TR></TABLE>"
	end if
%></TD></TR>
  </TABLE>
  <font color="#FFFFFF"></TD></font>
  <font color="#FFFFFF"></TR></font>
  <TR> 
    <TD> 
      <TABLE WIDTH=100% CELLSPACING="0" CELLPADDING="2" BORDER="1" BORDERCOLOR="#666666" BGCOLOR="BLACK">
        <TR> 
          <TD bgcolor="#3399CC"> <font face="Arial, Helvetica, sans-serif" color="#FFFFFF"> 
            <a href="http://10.0.7.20/um/opslog/files/default.asp?dir=/um"><b> 
            TOP</b></a></font> </TD>
        </TR>
      </TABLE>
    </TD>
  </TR>
</TABLE>
<%
	Set FSO = Nothing
	Set BrowseContents = Nothing
	end if
%>
</BODY>
</HTML>
