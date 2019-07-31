<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
response.write "Saving..."
dim action, value, user, job, id, temp, description, clkSource, oldjob, tmpMoveFrame
action=lcase(trim(Request("modify")))
value=Request("value")
'value=Formatcurrency(value, 2) '5/23/2008 N.Ambo removed because this line seems to prevent larger numbers fomr being entered
user="ghnet\"&trim(request("name"))

if (Request("job")="") then
	Response.Redirect  "timedetail.asp?name=" & trim(request("name"))
end if
job=Request("job")
oldjob = Request("oldjob")


id=Request("id")
temp=Request("date")
description=Replace(Request("description"),"'","''")
clkSource = trim(request("source"))

dim cnn1, rst, strsql,sqltest,rsttest
Set cnn1 = Server.CreateObject("ADODB.Connection")
set rst=Server.CreateObject("ADODB.Recordset")
set rsttest=Server.CreateObject("ADODB.Recordset")
cnn1.Open getConnect(0,0,"intranet")


if(action = "save") then
    if (Request("hrs")=0 and request("ot")=0) then
		Response.Redirect  "timedetail.asp?source="&clksource&"&day="&temp&"&job="&job&"&name="& trim(request("name"))
	end if
	
	'See if this is first time entered against this job
	rst.open	"select top 1 id from times where jobno='"&job&"'",cnn1
	if rst.eof then ' job is new
	  cnn1.execute "update master_job set last_invoice='"&(cdate(Request("date"))-1)&"' where id="&job
	end if
	rst.close
	set rst= Nothing
	strsql = "Insert into Times (date, jobno, description, hours, overt, expense, value, matricola) "_
	& "values ("_
	& "'" & Request("date") & "', "_
	& "'" & cint(Request("job")) & "', "_
	& "'" & description & "', "_
	& "'" & cdbl(Request("hrs")) & "', "_
	& "'" & cdbl(Request("ot")) & "', "_
	& "'" & Request("exp") & "', "& value &" ,"_
	& "'" & user & "')"
else		
  id=Request("id")
  if (Request("hrs")=0 and request("ot")=0)  then
	Response.Redirect  "timedetail.asp?source="&clksource&"&day="&temp&"&id="&id
  end if	
  if job<>oldjob then oldjob = ", entry_time='" & now() & "'" else oldjob = ""
  strsql = "Update Times Set date='" & Request("date") & "', jobno='" & job & "', description='" & description & "', hours='" & Request("hrs") & "', overt='" & Request("ot") & "', expense='" & Request("exp") & "', value=" & value & " "&oldjob&" where  id='"& id &"'"

end if
'response.write strsql
'response.end
IF not allowgroups("Timesheet Supervisors") THEN 
dim testdate


testdate = DateDiff("d", Request("date"),Date)
if testdate > 9 then
Response.Write "<script>" & vbCrLf
Response.Write "alert('Cannot change old time submissions.')"
Response.Write "</script>" & vbCrLf 
response.end
end if
end if


IF not allowgroups("Timesheet Supervisors") THEN  


'THIS IS USED TO STOP PEOPLE FROM GOING BACK TO RE-ENTER AND MODIFY OLD TIME SUBMISSION, SO WE LOG IT AND ONLY GIVE LUNA N CHRIS ACCESS TO DO SO
sqltest = "SELECT * FROM ARCHIVE_TIME_SUBMISSION WHERE '" & Request("date") & "' BETWEEN STARTWEEK AND ENDWEEK AND APPROVED =1 AND CAPPROVED =1 AND SUBMITTED=1 and username ='"&user&"'"
rsttest.open sqltest, cnn1
if rsttest.eof then
cnn1.execute strsql
else
rsttest.close
Response.Write "<script>" & vbCrLf
Response.Write "alert('Cannot enter time for this date ,time range already approved.')"
Response.Write "</script>" & vbCrLf 
end if
ELSE
cnn1.execute strsql ' FOR CERTAIN GROUP RIGHTS
END IF

Select case trim(clksource) 
	case "review"
		tmpMoveFrame =  "opener.document.location = opener.document.location;window.close()"
	case "personaltasks"
		tmpMoveFrame =  "window.close()"
	case else
		tmpMoveFrame =  "parent.frames.tstop.location = parent.frames.tstop.location;parent.frames.tsbottom.location ='timedetail.asp?name=" & request("name") & "'"
end select

Response.Write "<script>" & vbCrLf
Response.Write tmpMoveFrame
Response.Write "</script>" & vbCrLf 

%>