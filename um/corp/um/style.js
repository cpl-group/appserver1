//UDMv2.1.1
//**DO NOT EDIT THIS ******************************************
if (!exclude) { //*********************************************
// document object variable
var d = document;
// adjust font size discrepancy
if (ns4) { fSIZE+=1; sfSIZE+=1; }
// filter for bad definitions
if(bSIZE<0)bSIZE=0;if(fSIZE<5) fSIZE=5;if(tINDENT<0)tINDENT=0;if(vPADDING<0)vPADDING=0;
if(sbSIZE<0)sbSIZE=0;if(sfSIZE<5) sfSIZE=5;if(stINDENT<0)stINDENT=0;if(svPADDING<0)svPADDING=0;
// link style definitions
var stySTR='';
stySTR+='<style>';
stySTR+='.menubarTD A  \{ color: ' + aLINK + '\; font-weight:' + fWEIGHT + '\; \}';
stySTR+='.menubarTD A:Link  \{ color: ' + aLINK + ' \}';
stySTR+='.menubarTD A:Visited  \{ color: ' + aLINK + ' \}';
if (op5&&(aHOVER!="")) { stySTR+='.menubarTD A:Hover  \{ color: ' + aHOVER + ' \}'; }
stySTR+='.menubarTD A:Active,.menubarTD A:Link,.menubarTD A:Visited,.menubarTD A:Hover \{ font-weight:' + fWEIGHT + '\; font-size:' + fSIZE + 'px\; font-family:' + fFONT + '\; text-decoration:' + aDEC + '; position:relative\; \}'; 
stySTR+='.SUBmenubarTD A  \{  color: ' + saLINK + '\; font-weight:' + sfWEIGHT + '\; \}';
stySTR+='.SUBmenubarTD A:Link  \{ color: ' + saLINK + ' \}';
stySTR+='.SUBmenubarTD A:Visited  \{ color: ' + saLINK + ' \}';
if (op5&&(saHOVER!="")) { stySTR+='.SUBmenubarTD A:Hover  \{ color: ' + saHOVER + ' \}'; }
stySTR+='.SUBmenubarTD A:Active,.SUBmenubarTD A:Link,.SUBmenubarTD A:Visited,.SUBmenubarTD A:Hover \{ font-weight:' + sfWEIGHT + '\; font-size:' + sfSIZE + 'px\; font-family:' + sfFONT + '\; text-decoration:' + saDEC + '\; \}';
//*************************************************************
//*************************************************************


//** USE THIS SPACE FOR NEW STYLE DEFINITIONS *****************






var cl = '#000000'; var fs = 14; 
if (ns4) { cl = '#000096'; fs = 15; } 
stySTR+='.roman \{ font-size:' + fs + 'px\; color:' + cl + '\; background-color:white\; font-family:times new roman\; \}'; 









//**DO NOT EDIT THIS ******************************************
stySTR+='</style>';
d.write(stySTR);
}//************************************************************
//*************************************************************

