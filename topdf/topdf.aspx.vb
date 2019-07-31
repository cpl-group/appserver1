' file: PdfGenerator.aspx.vb
Imports System
Imports System.IO

Namespace Website
  Public Partial Class PdfGenerator
      Inherits System.Web.UI.Page
      Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs)
        Dim Response As System.Web.HttpResponse = System.Web.HttpContext.Current.Response
        Try
            ' create an API client instance
            Dim client As New pdfcrowd.Client("username", "apikey")

            ' convert a web page and write the generated PDF to a memory stream
            Dim Stream As New MemoryStream
            client.convertURI("hhttp://appserver1.genergy.com/genergy2/invoices/GenergyInvoice.H2O.asp?demo=&l=15545&y=79221&building=RNEEMBKA&pid=163&byear=2017&bperiod=9&logo=invoice_logo_1.jpg&genergy2=true&utilityid=3&detailed=false&meterbreakdown=&SJPproperties=&summaryusage=&summarydemand=&textheader=&billid=", Stream)

            ' set HTTP response headers
            Response.Clear() 
            Response.AddHeader("Content-Type", "application/pdf")
            Response.AddHeader("Cache-Control", "max-age=0")
            Response.AddHeader("Accept-Ranges", "none")
            Response.AddHeader("Content-Disposition", "attachment; filename=google_com.pdf") 

            ' send the generated PDF
            Stream.WriteTo(Response.OutputStream)
            Stream.Close()
            Response.Flush() 
            Response.End()
        Catch why As pdfcrowd.Error
            Response.Write(why.ToString())
        End Try
      End Sub
  End Class
End Namespace