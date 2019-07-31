<%
Set Upload = Server.CreateObject("Persits.Upload.1")

' Upload files
Upload.OverwriteFiles = False ' Generate unique names
Upload.SetMaxSize 1048576 ' Truncate files above 1MB
Upload.Save "\upload" 

' Process all files received
For Each File in Upload.Files

  ' Save in the database as blob
  'File.ToDatabase "DSN=data;UID=sa;PWD=zzz;", _ 
   '"insert into mytable(blob) values(?)"

  ' Move to a different location
  'File.Copy "d:\archive\" & File.ExtractFileName
  'File.Delete
Next
' Display description field
Response.Write Upload.Form("Description") & "<BR>"

' Display all selected categories
For Each Item in Upload.Form
  If Item.Name = "Category" Then
    Response.Write Item.Value & "<BR>"
  End If
Next
%>