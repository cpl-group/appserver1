<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
if 	not(allowGroups("Genergy Users,clientOperations")) then
%><!--#INCLUDE VIRTUAL="/genergy2/securityerror.inc"--><%
end if

dim bldg, action, tid, pid, lid, transfer, oldtid
lid = request("lid")
tid = request("tid")
pid = request("pid")
bldg = request("bldg")
action = request("action")
transfer = request("transfer")
oldtid = request("oldtid")
'dim DBmainmodIP
'DBmainmodIP = "["&getPidIP(pid)&"].mainmodule.dbo."

dim cnn1, rst1, strsql, strsql2, rst2
set cnn1 = server.createobject("ADODB.connection")
set rst1 = server.createobject("ADODB.recordset")
set rst2 = server.CreateObject("ADODB.recordset")
cnn1.open getLocalConnect(bldg)

dim AdminFee, TenantRate, AddonFee, ModifyRate, Coincident, CoinWithPeak, FullOnPeak, utility, procname, calcintpeak, acctidref, use_acctid, billnote, shadow
'kCheng 6/2/2009 - added two more field for unit of measure display
dim cMeasure, dMeasure
cMeasure = request("consumptionMeasure")
dMeasure = request("demandMeasure")

AdminFee = request("AdminFee")
TenantRate = request("TenantRate")
AddonFee = request("AddonFee")
ModifyRate = request("ModifyRate")
Coincident = request("Coincident")
CoinWithPeak = request("CoinWithPeak")
FullOnPeak = request("FullOnPeak")
utility = request("utility")
procname = request("procname")
calcintpeak = request("calcintpeak")
acctidref = request("acctidref")
use_acctid = request("use_acctid")
billnote = left(request("billnote"),250)
shadow = trim(request("shadow"))

if trim(FullOnPeak)="" then FullOnPeak = 0
if trim(Coincident)="" then Coincident = 0
if trim(CoinWithPeak)="" then CoinWithPeak = 0
if trim(AddonFee)="" then AddonFee = 0
if trim(calcintpeak)="" then calcintpeak=0
if trim(use_acctid)="" then use_acctid=0
if trim(shadow)=""  then shadow=0

if trim(action)="Save" then
	strsql = "INSERT INTO tblleasesutilityprices (AdminFee, RateTenant, AddonFee, RateModify, Coincident, Coincident_peak, FullOnPeak,billingid,utility, procname, calcintpeak, acctid, use_acctid, bill_note, shadow) values ('"&AdminFee&"', '"&TenantRate&"', "&AddonFee&", '"&ModifyRate&"', '"&Coincident&"', '"&CoinWithPeak&"', '"&FullOnPeak&"', '"&tid&"', '"&utility&"', '"&procname&"', '"&calcintpeak&"', '"&acctidref&"', "&use_acctid&", '"&billnote&"', '"&shadow&"')"
elseif trim(action)="Delete" then
	strsql = "DELETE FROM tblleasesutilityprices WHERE leaseutilityid="&lid
else
	strsql = "UPDATE tblleasesutilityprices set AdminFee="&AdminFee&", RateTenant='"&TenantRate&"', AddonFee="&AddonFee&", RateModify='"&ModifyRate&"', Coincident='"&Coincident&"', Coincident_peak='"&CoinWithPeak&"', FullOnPeak='"&FullOnPeak&"', billingid='"&tid&"', utility='"&utility&"', procname='"&procname&"', calcintpeak='"&calcintpeak&"', acctid='"&acctidref&"', use_acctid='"&use_acctid&"', bill_note='"&billnote&"', shadow='"&shadow&"' WHERE leaseutilityid="&lid
end if
'response.Write strsql
'response.End
logger strsql
cnn1.Execute strsql
if trim(action)="Save" then
	rst1.open "SELECT top 1 leaseutilityid FROM tblleasesutilityprices ORDER BY leaseutilityid desc", cnn1
	if not rst1.eof then lid = rst1(0)
	rst1.close
end if

'KCheng 6/2/2009 - added for Unit of display
if (utility = 6 OR utility = 21) then

    if trim(action)="Save" then
         strsql2 = "INSERT INTO tblleasespecificmeasure (LeaseutilityId, ConsumptionMeasure, DemandMeasure) VALUES("&lid&", '"&cMeasure&"', '"&dMeasure&"')"
    elseif trim(action)="Delete" then
         strsql2 = "DELETE FROM tblleasespecificmeasure where LeaseutilityId="&lid
    else
            strsql2 = "Select * FROM tblleasespecificmeasure where LeaseutilityId="&lid
        rst2.open strsql2, cnn1
        if rst2.EOF then
            strsql2 = "INSERT INTO tblleasespecificmeasure (LeaseutilityId, ConsumptionMeasure, DemandMeasure) VALUES("&lid&", '"&cMeasure&"', '"&dMeasure&"')"
        else
            strsql2 = "UPDATE tblleasespecificmeasure set ConsumptionMeasure='"&cMeasure&"', DemandMeasure='"&dMeasure&"' where LeaseutilityId="&lid
        end if    
    end if
if (strsql2 <>null OR strsql2 <> "") then
    cnn1.execute strsql2
end if

end if ' end  new codes 

dim returnpage
if trim(transfer) <> "" and trim(oldtid) <> "" then
  returnpage = "newleaseutilityedit.asp?pid="&pid&"&bldg="&bldg&"&tid="&tid&"&oldtid="&oldtid
else
	if trim(action)="Save" then
		returnpage = "meteredit.asp?pid="&pid&"&bldg="&bldg&"&tid="&tid&"&lid="&lid
	else
		returnpage = "tenantedit.asp?pid="&pid&"&bldg="&bldg&"&tid="&tid&"&lid="&lid
	end if
end if

Response.Redirect returnpage
%>