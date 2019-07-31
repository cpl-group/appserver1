<% 
'disable caching so that image is always retrieved from server
Response.Buffer = true
Response.CacheControl = "Private"
Response.Expires = -1000

'On Error Resume Next

Dim objChart 'chart object
Dim rt 'stores return value of function calls

'create an intsance of the chart object
Set objChart = server.CreateObject("Dundas.ChartServer")

'set directory properties
'objChart.DirTemplate = "c:\dschart\dschart_local\templates\area\"
'objChart.DirTexture = "c:\dschart\dschart_local\textures\"

'load a template
rt = objChart.LoadTemplate("area.cuc")

'add a series of data 
rt = objChart.AddSeries("ACME Inc.")
'now add data to this first series
rt = objChart.AddData(5)
rt = objChart.AddData(8)
rt = objChart.AddData(5)
rt = objChart.AddData(5)
rt = objChart.AddData(9)

'add another series of data for demonstration purposes
rt = objChart.AddSeries("XYZ Corp.")
'now add data to this second series
rt = objChart.AddData(4)
rt = objChart.AddData(9)
rt = objChart.AddData(4)
rt = objChart.AddData(11)
rt = objChart.AddData(12)

'set up the axis
rt = objChart.SetAutoAxisRanges

'set the tick mark labels for the X-axis (optional)
'note that tick mark labels are not carried over from the template
rt = objChart.AddLabel("January")
rt = objChart.AddLabel("February")
rt = objChart.AddLabel("March")
rt = objChart.AddLabel("April")
rt = objChart.AddLabel("May")

'we will set the Y-axis label (optional)
rt = objChart.SetAxisName("Sales",1)

'now attempt to generate the jpeg, specifying width, height, and compression
rt = objChart.SendJpeg(600,450,0)

'release resources
Set objChart = Nothing
%>

