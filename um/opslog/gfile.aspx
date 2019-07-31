
<%@ Page Language="VB" %>

<script runat="server">
    Protected Sub Button1_Click(ByVal sender As Object, _
      ByVal e As System.EventArgs)
        Dim strpath
        strpath = Request("path")
        If FileUpload1.HasFile Then
            Try
                FileUpload1.SaveAs("D:\WebSites\isabella\appserver1\"&strpath & _
                   FileUpload1.FileName)
                Label1.Text = "File name: " & _
                   FileUpload1.PostedFile.FileName & "<br>" & _
                   "File Size: " & _
                   FileUpload1.PostedFile.ContentLength & " kb<br>" & _
                   "Content type: " & _
                   FileUpload1.PostedFile.ContentType
            Catch ex As Exception
                Label1.Text = "ERROR: " & ex.Message.ToString()
            End Try
        Else
            Label1.Text = "You have not specified a file."
        End If
                
    End Sub
    
   
</script>


<script type="text/javascript">
    function refreshAndClose() {
       
        window.opener.location.reload(true);
        window.close();
    }
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Upload Files</title>
</head>
<body >
    <form id="form1" runat="server">
    <div>
        <asp:FileUpload ID="FileUpload1" runat="server" /><br />
        <br />
        <asp:Button ID="Button1" runat="server" OnClick="Button1_Click" width="150px"
         Text="Upload File" />&nbsp;<br />
        <br />
        <asp:Button ID="Button2" runat="server" OnClientClick="javascript:window.refreshAndClose()" Width="150px"
         Text="Close Window" />&nbsp;<br />
        <br />
        <asp:Label ID="Label1" runat="server"></asp:Label></div>
    </form>
</body>
</html>