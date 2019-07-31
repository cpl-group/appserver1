<%option explicit

dim rst, cnn
set cnn = server.createobject("ADODB.Connection")
set rst = server.createobject("ADODB.RecordSet")


cnn.open "driver={SQL Server};server=168.103.58.249;uid=sa;pwd=!general!;database=genergy1;"


%>
hey