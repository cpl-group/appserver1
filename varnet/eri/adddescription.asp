<%@Language="VBScript"%>
<%
	
choice=Request("submit")
count=Request("count")

Set cnn1 = Server.CreateObject("ADODB.Connection")
Set rst1 = Server.CreateObject("ADODB.recordset")
cnn1.Open "driver={SQL Server};server=10.0.7.110;uid=genergy1;pwd=g1appg1;database=eri_data;"

    strsql = "Insert into tblsurveylib (type,description,amps,volt,ph,pf,watt,monthfactor,adjfactor) "_
	& "values ("_
	& "'" & Request.Form("type") & "', "_
	& "'" & Request.form("description") & "', "_
	& Request.Form("amps") & ", "_
	& Request.Form("volts") & ", "_
 	& Request.form("ph") & ", "_
        & Request.Form("pf") & ", "_
        & Request.Form("watt") & ", "_
		& Request.Form("adjfactor") & ", "_
	& Request.Form("monthfactor") & ")"

cnn1.execute strsql
set cnn1=nothing
%>
<script>
			opener.document.forms[<%=count%>].description.value="<%=Request.form("description")%>"
			opener.document.forms[<%=count%>].amps.value="<%=Request.Form("amps")%>"
			opener.document.forms[<%=count%>].volt.value="<%=Request.Form("volts")%>"
			opener.document.forms[<%=count%>].ph.value="<%=Request.Form("ph")%>"
			opener.document.forms[<%=count%>].pf.value="<%=Request.Form("pf")%>"
			opener.document.forms[<%=count%>].mf.value="<%=Request.Form("monthfactor")%>"
			opener.document.forms[<%=count%>].watt.value="<%=Request.Form("watt")%>"
			opener.document.forms[<%=count%>].totkw.value=opener.document.forms[<%=count%>].qty.value*<%=Request.Form("watt")%>
			opener.document.forms[<%=count%>].adj.value="<%=Request.Form("adjfactor")%>"
			opener.document.forms[<%=count%>].adjkw.value=opener.document.forms[<%=count%>].qty.value*<%=(Request.Form("adjfactor")*Request.Form("watt"))%>
			window.close()
</script>
