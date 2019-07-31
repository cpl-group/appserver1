// JScript source code for buildingTransfer.asp

function openCustomWin(clink, cname, cspec)
{	cWin = window.open(clink, cname, cspec)
  cWin.focus();
}
function pidChange(pid)
{
    document.location.href = 'buildingTransfer.asp?pid=' + pid
}
function checkForm() 
{
    var returnVal = true;
    var message = "Error! Please check the following and try again. \n\n";
    var goodDay = true;
    if (document.buildingXferFrm.pid.value == "") 
    {
        returnVal=false;
        message += "Old Protfolio Name is missing \n";
    }
    
    if (document.buildingXferFrm.newPid.value == "")
    {
        returnVal= false;
        message += "New Protfolio Name is missing \n";
    }
    if ((document.buildingXferFrm.xferDate.value == "dd/mm/yy") || (document.buildingXferFrm.xferDate.value == ""))
    {
        returnVal=false;
        message += "Transfer Date is missing \n";
    }
    else    
    {
        var validDate = isDate(document.buildingXferFrm.xferDate.value)
        if (validDate == false)
        {
            returnVal = false;
            goodDate = false;
            message += "Invalid Transfer Date \n";
         }
    }
    
     var atLeastOneChecked = false;
     var buildStr = "";
     var frm = document.buildingXferFrm;
     for (var i = 0; i < frm.elements.length; i++ ) {
        if (frm.elements[i].type == 'checkbox') {
            if (frm.elements[i].checked == true) {
                atLeastOneChecked = true;
                if (buildStr == "") 
                    buildStr = frm.elements[i].name;
                else
                    buildStr += "+" + frm.elements[i].name;
            }
        }
     }
    if (atLeastOneChecked == false)
    {
        returnVal = false;
        message += "Plesea select at least one building to transfer \n";
     }   
     
    if(buildStr != "")
        document.buildingXferFrm.buildingStr.value = buildStr;
    
    if (returnVal == false)
        alert(message);
    
    return returnVal;    
}

function isDate(sDate) {
    var scratch = new Date(sDate);
    if (scratch.toString() == "NaN" || scratch.toString() == "Invalid Date") {
        return false;
    } 
    else 
    {
        return true;
    }
}
