<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
	dim pid,bldgNum,bldgname,sDataSource,strDataSource,UtilityId,rst2,str,strQuery,strmsg
	strmsg=""
	pid = Request.QueryString("pid")
	bldgNum = request("bldgNum")
	bldgname = request("bldgname")
	
	if request("bldgNum")="" then
		bldgNum = trim(Request("BuildingNumber"))
	End if
	if request("bldgname")="" then
		bldgname = trim(Request("BuildingName"))
	End if
	
	
	'strDataSource="pulse_" & trim(Request("BuildingNumber")) & "_" & trim(secureRequest("DataSource"))
	strDataSource="pulse_" & trim(Request("BuildingNumber")) & trim(secureRequest("DataSource"))
	'UtilityId=Request("ddlUtility")
	if secureRequest("UtilityId")="" then
		UtilityId=Trim(Request("Utility"))
	else
		UtilityId=Trim(secureRequest("UtilityId"))
	End if
	
	if UtilityId="" Then
		UtilityId=2
	End If
	
	
	if Request("action")="Save" then
		strQuery=""
		Select Case UtilityId
		Case 2	'Electricty(KWH)
			strQuery="CREATE TABLE " & strDataSource &_
						"( " &_
						"[id] [int] IDENTITY(1,1) NOT NULL, " &_
						"[bldgnum] [varchar](10) NOT NULL DEFAULT ('010'), " &_
						"[meterid] [int] NOT NULL DEFAULT (0), " &_
						"[date] [smalldatetime] NOT NULL DEFAULT (1 / 1 / 1900), " &_
						"[pulse] [numeric](18, 3) NOT NULL DEFAULT (0), " &_
						"[delta] [int] NOT NULL DEFAULT (0), " &_
						"[kwh] [decimal](18, 2) NOT NULL DEFAULT (0), " &_
						"[billyear] [int] NOT NULL DEFAULT (0), " &_
						"[billperiod] [tinyint] NOT NULL DEFAULT (0), " &_
						"[validated] [int] NOT NULL DEFAULT (0)," &_
						"[est] [int] NOT NULL DEFAULT (0)," &_
						"[pd] [decimal](18, 2) NULL, " &_
						"[est_value] [decimal](18, 2) NULL, " &_
						"[process] [bit] NOT NULL DEFAULT (0), " &_
						"[PrevDelta] [int] NULL, " &_
						"[PctDifference]  AS (case when (([PrevDelta] is null or [PrevDelta] = 0)) " &_
						"then null else (100.0 * (convert(float(4),[delta]) / convert(float(4),[PrevDelta]))) end), " &_
						"CONSTRAINT pk_" & strDataSource & " PRIMARY KEY NONCLUSTERED (meterid, [date]) ) "
		Case 6, 21 'Chilled Water, or Condenser Water
			strQuery="CREATE TABLE " & strDataSource &_
					"(" &_
					"[id] [int] IDENTITY(1,1) NOT NULL," &_
					"[meterid] [int] NOT NULL," &_
					"[date] [datetime] NOT NULL," &_
					"[billyear] [int] NULL  DEFAULT (0)," &_
					"[billperiod] [int] NULL  DEFAULT (0)," &_
					"[TONS] [decimal](18, 4) NULL," &_
					"[CHWR] [decimal](9, 4) NULL," &_
					"[DT] [decimal](18, 4) NULL," &_
					"[TTons] [decimal](18, 4) NULL," &_
					"[bldgnum] [varchar](50) NULL," &_
					"[chws] [decimal](18, 3) NULL," &_
					"[GPM] [decimal](9, 0) NULL," &_
					"[DTTONS] [decimal](18, 2) NULL," &_
					"[ouctons] [bit] NOT NULL DEFAULT (0)," &_
					"[HUM] [int] NULL," &_
					"[TONHRS] [decimal](18, 2) NULL," &_
					"[TEMP] [int] NULL," &_
					"[calctons] [decimal](18, 2) NULL," &_
					"[pd] [decimal](18, 2) NULL DEFAULT (0)," &_
					"[est] [int] NOT NULL DEFAULT (0)," &_
					"[est_value] [decimal](18, 2) NULL," &_
					"[validated] [int] NOT NULL DEFAULT (0)," &_
					"[process] [bit] NOT NULL DEFAULT (0)," &_
					"[pulse] [decimal](18, 4) NULL," &_  
					"CONSTRAINT [PK_" & strDataSource & "] PRIMARY KEY NONCLUSTERED (meterid, [date])) "
		Case 1 'Steam  --N.Ambo added 5/21/2009
			strQuery="CREATE TABLE " & strDataSource &_
					"(" &_
					"[id] [int] IDENTITY(1,1) NOT NULL," &_
					"[meterid] [int] NOT NULL," &_
					"[date] [datetime] NOT NULL," &_
					"[billyear] [int] NULL  DEFAULT (0)," &_
					"[billperiod] [int] NULL  DEFAULT (0)," &_
					"[MLBS] [decimal](18, 2) NULL," &_
					"[Pressure] [decimal](9, 4) NULL," &_
					"[MLBhrs] [decimal](18, 4) NULL," &_
					"[bldgnum] [varchar](50) NULL," &_
					"[pd] [decimal](18, 2) NULL DEFAULT (0)," &_
					"[est] [int] NOT NULL DEFAULT (0)," &_
					"[est_value] [decimal](18, 2) NULL," &_
					"[validated] [int] NOT NULL DEFAULT (0)," &_
					"[process] [bit] NOT NULL DEFAULT (0)," &_
					"[pulse] [decimal](18, 4) NULL," &_  
					"[kwh] [decimal](18, 4) NULL," &_
					"CONSTRAINT [PK_" & strDataSource & "] PRIMARY KEY NONCLUSTERED (meterid, [date])) "
			Case 10,3 'Hot or Cold Water --N.Ambo added 5/21/2009
			strQuery="CREATE TABLE " & strDataSource &_
					"(" &_
					"[id] [int] IDENTITY(1,1) NOT NULL," &_
					"[meterid] [int] NOT NULL," &_
					"[date] [datetime] NOT NULL," &_
					"[billyear] [int] NULL  DEFAULT (0)," &_
					"[billperiod] [int] NULL  DEFAULT (0)," &_
					"[cf] [decimal](18, 4) NULL," &_
					"[bldgnum] [varchar](50) NULL," &_
					"[pd] [decimal](18, 2) NULL DEFAULT (0)," &_
					"[est] [int] NOT NULL DEFAULT (0)," &_
					"[est_value] [decimal](18, 2) NULL," &_
					"[validated] [int] NOT NULL DEFAULT (0)," &_
					"[process] [bit] NOT NULL DEFAULT (0)," &_
					"[pulse] [decimal](18, 4) NULL," &_  
					"[kwh] [decimal](18, 4) NULL," &_
					"CONSTRAINT [PK_" & strDataSource & "] PRIMARY KEY NONCLUSTERED (meterid, [date])) "
		Case Else 'N.Ambo amedned 5/21/2009 to add extra fields
			strQuery="CREATE TABLE " & strDataSource &_
					"(" &_
					"[id] [int] IDENTITY(1,1) NOT NULL," &_
					"[meterid] integer," &_
					"[date] datetime," &_ 
					"billyear integer," &_
					"billperiod integer," &_
					"[bldgnum] [varchar](50) NULL," &_
					"[pd] [decimal](18, 2) NULL DEFAULT (0)," &_
					"est integer NOT NULL default 0," &_
					"est_value integer," &_
					"validated bit NOT NULL default 0," &_
					"[process] [bit] NOT NULL DEFAULT (0)," &_
					"[pulse] [decimal](18, 4) NULL," &_  
					"[kwh] [decimal](18, 4) NULL," &_
					"CONSTRAINT pk_" & strDataSource & " PRIMARY KEY NONCLUSTERED (meterid, [date])) "
		End Select
		'Response.Write(strQuery)
		if strQuery<>"" then
			dim rst,rstcheck,strCheck
			set rst	= server.createobject("ADODB.Recordset")
			set rstcheck = server.createobject("ADODB.Recordset")
			
			strCheck="Select name, type From Sysobjects Where name ='" & strDataSource & "'"	'Checking if the table exists
			rstcheck.Open strCheck, getConnect(0,0,"intervaldata"), 0, 1, 1
			if rstcheck.EOF then	'if not then
			
				'Creating Table 
				rst.open strQuery, getConnect(0,0,"intervaldata")
				set rst=nothing
				
				'Saving the Info in the BuildingDataSource
				dim rstds,strds
				strds=	"INSERT INTO [BuildingDataSource]([DataSource],[bldgnum],[UtilityId]) " &_
						"VALUES('"& strDataSource &"','"& bldgNum &"',"& UtilityId &")"
						
				Set rstds=Server.CreateObject("ADODB.Recordset")
				rstds.Open strds,getConnect(0,0,"billing"), 0, 1, 1
				set rstds=nothing 
				
				strmsg="Table Created!"
			else	
				strmsg="Table with the Same Name Exists!"
				Response.Write(strmsg)
					
			End if
			rstcheck.Close 
		End if	
	End if

%>
<head>
	<link rel="Stylesheet" href="/genergy2/SETUP/setup.css" type="text/css">
	<title>Add New DataSource</title>
	<script>
		function closeWinda(){ 
		window.close()
		}
		
		function checkform(frm){
			var err = "";
			if(frm.DataSource.value=='' & frm.ddlUtility.value != 2 ) err+="No DataSource name entered\n";
			if(frm.ddlUtility.value=='') err+="Select Utility name\n"
			
			if(err=="") 
				return window.confirm('Create DataSource with the Name  pulse_' + frm.BuildingNumber.value +  frm.DataSource.value + '   !');
			else alert(err);
				return false;
		}
	</script>
</head>

<body bgcolor="#eeeeee" topmargin=0 leftmargin=0 marginwidth="0" marginheight="0">
	<form action="AddNewDataSource.asp" onsubmit="return(checkform(this))">
		<table width="100%" border="0" cellpadding="3" cellspacing="0" align="center">
			<tr bgcolor="#3399cc">
				<td><font color='white'>Add New DataSource for  </font>&nbsp;<b>Building # <%=bldgNum%>(<%=bldgname%>)</b></td>
			</tr>
		</table>
		 &nbsp;
		<div id="Datasourceinfo" style="BORDER-RIGHT: #cccccc 1px solid; PADDING-RIGHT: 3px; BORDER-TOP: #cccccc 1px solid; DISPLAY: inline; PADDING-LEFT: 3px; PADDING-BOTTOM: 3px; BORDER-LEFT: #cccccc 1px solid; WIDTH: 98%; PADDING-TOP: 3px; BORDER-BOTTOM: #cccccc 1px solid"> 
		  <table width="100%" border="0" cellpadding="3" cellspacing="0">
		    <tr> 
		      <td>Data Source Name</td>
		      <td><font size=-2>pulse_ <%=bldgNum%></font></td>
		      <td><INPUT style="WIDTH: 202px; HEIGHT: 22px" size=26 name="DataSource" value="<%=sDataSource%>"></td>
		    </tr>
		    <tr> 
		      <td>Utility Name</td>
		      <td>&nbsp;</td>
		      <td>
				<select name="ddlUtility" style="WIDTH: 202px; HEIGHT: 22px" onChange="document.location='AddNewDataSource.asp?pid=<%=pid%>&bldgNum=<%=bldgNum%>&bldgname=<%=bldgname%>&UtilityId='+this.value">
					<%Set rst2 = Server.CreateObject("ADODB.recordset")
					str="select * from tblutility order by utility"
					rst2.Open str, getConnect(0,0,"dbCore"), 0, 1, 1
					do until rst2.eof%>
					<option value="<%=rst2("utilityid")%>" <%if trim(UtilityId)=trim(rst2("utilityid")) then response.write " SELECTED"%>><%=rst2("utilitydisplay")%>
					</option>
					<%
					rst2.movenext
					loop
					rst2.close%>
				</select>
		      </td>
		    </tr>
		    <tr>
				<td>&nbsp;</td>
				<td align="left" colspan="3">
					<input name="action" style="BORDER-RIGHT: #ffffff 1px outset; BORDER-TOP: #ffffff 1px outset; BORDER-LEFT: #ffffff 1px outset; CURSOR: hand; COLOR: #336699; BORDER-BOTTOM: #ffffff 1px outset; BACKGROUND-COLOR: #eeeeee" type="submit" value="Save" >&nbsp;
					<input name="close"  style="BORDER-RIGHT: #ffffff 1px outset; BORDER-TOP: #ffffff 1px outset; BORDER-LEFT: #ffffff 1px outset; CURSOR: hand; COLOR: #336699; BORDER-BOTTOM: #ffffff 1px outset; BACKGROUND-COLOR: #eeeeee" type="button" value="Cancel" onclick="javascript:closeWinda();">
				</td>
			</tr>
		  </table>
		</div>
		 &nbsp;
		<TABLE width="100%" border="0" cellpadding="3" cellspacing="0" align="center">  
		  <TR>
		    <TD align=absmiddle bgColor=#3399cc><font color='white'>Existing Data Source</font></TD>
		  </TR>
		  
		    <% 
		    str="" 
		    Set rst2 = Server.CreateObject("ADODB.recordset")
				str="select * from BuildingDataSource Where UtilityId="& UtilityId &" And bldgnum ='"& bldgNum &"'"
				rst2.Open str, getConnect(0,0,"billing"), 0, 1, 1	
				do until rst2.eof%>
				<tr bgcolor="#cccccc">
					<td><span class="standard"><%=rst2("DataSource")%></span></td>
				</tr>	
				<%
				rst2.movenext
				loop
			rst2.close
			%>
		</TABLE>
		 <p>
		  <input type="hidden" name="Utility" value="<%=UtilityId%>">
		  <input type="hidden" name="BuildingNumber" value="<%=bldgNum%>">
		  <input type="hidden" name="BuildingName" value="<%=bldgname%>">
		  <input type="hidden" name="PortfolioId" value="<%=pid%>">
		</p>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<p>&nbsp;</p>
		<p>&nbsp; </p>
		<p>&nbsp;</p>
		<p>&nbsp; </p>
	</form>
	<TABLE cellSpacing=0 cellPadding=0 width="100%" align=center  border=0 style="FONT-SIZE: smaller; BOTTOM: 0px; TEXT-ALIGN: center">
		<TR>
			<TD align=absmiddle bgColor=#3399cc><FONT color=white><%=strmsg%></FONT></TD>
		</TR>
	</TABLE>
</body>