// JScript source code for historic_acctFile.asp

function acctselect(building,utility){
	var temp = "acctlist.asp?building=" + building+ "&utility="+utility
	window.open(temp,"selectaccount", "scrollbars=yes,width=1000, height=300, status=no" );
//	window.selectaccount.focus();
}
function ypid(id1,building,utility){
	var temp = "ypid.asp?acctid=" + id1 + "&building=" + building+ "&utility=" + utility
	window.open(temp,"ypid", "scrollbars=yes,width=500, height=300, status=no" );
}
function setup(building,utility){
	clearselections();
	var temp = "editacct.asp?building="+building+ "&utility=" + utility
	document.all.entryframe.style.visibility = "visible"
	document.frames.entry.location=temp;
	}

function editacct(acctid)
{	if(acctid.length>0)
	{	var temp = "editacct.asp?acctid=" +acctid
		document.frames.entry.location=temp;
	}else
	{	alert('Select an Account');
	}
}

function clearselections()
{	document.frames.entry.location='about:blank';
	document.form1.acctid.value=''
	document.all['accountdisplay'].innerText='No Account Selected'
	document.all['enterbillbutton'].style.visibility='hidden'
}

function loadportfolio()
{	
    var frm = document.forms['form1'];
	var newhref = "entry.asp?pid="+frm.pid.value;
	document.location.href=newhref;
}

function loadbuilding(pid, building)
{	
    document.location.href = 'historic_acctFile.asp?pid=' + pid + '&bldg='+building;
}
function JumpTo(url, pid , bldg)
{
	var frm = document.forms['form1'];
	var url = url + "?pid="+pid+"&bldg="+bldg+"&building="+bldg+"&utilityid=2";
	window.document.location=url;
}
function checkForm()
{
    var returnVal = true;
    var goodDate = true;
    var message = "Following error : \n\n";
    if (document.acctTransForm.pid.value == "")
    {
        returnVal=false;
        message += "Portfolio information is missing \n";
    }
    if (document.acctTransForm.bldg.value == "")
    {
        returnVal = false;
        message += "Building number is missing \n";
    }
    if(document.acctTransForm.tenant.value =="")
    {
        returnVal=false;
        message += "Tenant information is missing \n";
    }
    if (document.acctTransForm.util.value =="")
    {   
        returnVal=false;
        message += "Utility information is missing \n";
    }
    if (document.acctTransForm.billPeriod.value == "")
    {
        if ((document.acctTransForm.dateFrom.value == "dd/mm/yy") || (document.acctTransForm.dateTo.value == ""))
        {
            returnVal = false;
            message += "Date From is missing \n";
        }
        else 
        {
            var validDate = isDate(document.acctTransForm.dateFrom.value)
            if (validDate == false)
            {
                returnVal = false;
                goodDate = false;
                message += "Invalid Date From \n";
             }
         }
         
        if ((document.acctTransForm.dateTo.value == "dd/mm/yy") || (document.acctTransForm.dateTo.value == ""))
        {
            returnVal = false;
            message += "Date To is missing \n";
        }
        else 
        {
            var validDate = isDate(document.acctTransForm.dateTo.value)
            if (validDate == false)
            {
                returnVal = false;
                goodDate = false;
                message += "Invalid Date To \n";
             }
         }
         if ((document.acctTransForm.dateTo.value != "") && (document.acctTransForm.dateFrom.value != "") && goodDate)
         {
            var date1 = document.acctTransForm.dateFrom.value;
            var date2 = document.acctTransForm.dateTo.value;
            if (date2 < date1)
            {
                returnVal = false;
                message += "Date From has to be after Date To /n";
            }
         }
     }
     
     if ((document.acctTransForm.dateTo.value =="") && (document.acctTransForm.dateFrom.value ==""))
     {
         if (document.acctTransForm.billPeriod.value == "")
         {
            returnVal=false;
            message += "Bill Period is missing \n";
         }
     }
     
    if (returnVal == false)
    {
        alert(message);
    }
    
    return returnVal;
}

function isDate(sDate) 
{
    var scratch = new Date(sDate);
    if (scratch.toString() == "NaN" || scratch.toString() == "Invalid Date") {
        return false;
    } 
    else 
    {
        return true;
    }
}

function popUp(newPage, title){
  popper = window.open(newPage, title,"width=400,height=300,scrollbars=1,status=0,resizeable=1")
}

