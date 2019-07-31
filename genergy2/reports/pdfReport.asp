<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<%
dim qs
qs = request("qs")

dim apw, pdf, r, result, pdffilename
Set APW = Server.CreateObject("APWebGrabber.Object")
Set PDF = Server.CreateObject("APServer.Object")
pdf.OutputDirectory = "c:\appser~1\eri_th\pdfMaker"
APW.PrintBackGroundColors = 3
APW.URL = "http://appserver1.genergy.com" & qs & "&pdf=yes"
'response.write "http://appserver1.genergy.com" & qs & "&pdf=yes"
'response.end
APW.EngineToUse = 0
APW.MainFont "Helvetica",8,False,False,False,False
APW.FontNormal "Helvetica",8,False,False,False,False
'APW.HeaderHeight = 1.4
APW.Prt2DiskSettings = pdf.ToString()
R= APW.DoPrint("127.0.0.1",64320)
' Now wait for a result
Result = APW.Wait("127.0.0.1",64320,120,"")
' To get the name of the PDF, you have to use the activePDF Server object
If Result = "019" Then ' This was a good request
  ' Get the
  PDF.FromString APW.Prt2DiskSettings
  PDFFileName = "http://appserver1.genergy.com/eri_th/pdfMaker/" & PDF.NewUniqueID & ".pdf"
  Response.ContentType = "application/pdf"
  response.redirect PDFFileName
  %>
  <script>
  document.all['pdflink'].href = '<%=pdffilename%>'
  </script>
  <%
  Response.Write pdffilename
Else
	Response.Write("Error! " & Result)
End If
call APW.Cleanup("127.0.0.1",64320)
%>