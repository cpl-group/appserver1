//UDMv2.1.1
//**DO NOT EDIT THIS *****
if (!exclude) { //********


var menuALIGN = "left";		// alignment
var absLEFT = 	0;		// absolute left or right position (if not center)
var absTOP = 0; 		// absolute top position

var staticMENU = false;		// static positioning mode (not Opera 5)

var stretchMENU = true;		// show empty cells
var showBORDERS = false;		// show empty cell borders

var baseHREF =	"";	        // base path to .js and image files
var zORDER = 	100;		// base z-order of nav structure (not ns4)

var mCOLOR = 	"#3399CC";	        // main nav cell color
var rCOLOR = 	"lightblue";	// main nav cell rollover color
var keepLIT =	true;		// keep rollover color when browsing menu
var bSIZE = 	1;		// main nav border size
var bCOLOR = 	"#3399CC"	// main nav border color
var aLINK = 	"#FFFFFF";	// main nav link color
var aHOVER = 	"";		// main nav link hover-color (not ns4)
var aDEC = 	"none";		// main nav link decoration
var fFONT = 	"arial";	// main nav font face		
var fSIZE = 	13;		// main nav font size (pixels)	
var fWEIGHT = 	"bold"		// main nav font weight
var tINDENT = 	7;		// main nav text indent (if text is left or right aligned)
var vPADDING = 	5;		// main nav vertical cell padding
var vtOFFSET = 	0;		// main nav vertical text offset (+/- pixels from middle)

var vOFFSET = 	-5;		// shift the submenus vertically
var hOFFSET = 	4;		// shift the submenus horizontally

var smCOLOR = 	"lightblue";	// submenu cell color
var srCOLOR = 	"lightgreen";	// submenu cell rollover color
var sbSIZE = 	1;		// submenu border size
var sbCOLOR = 	"black"	// submenu border color
var saLINK = 	"black";	// submenu link color
var saHOVER = 	"";		// submenu link hover-color (not ns4)
var saDEC = 	"none";		// submenu link decoration
var sfFONT = 	"comic sans ms";// submenu font face		
var sfSIZE = 	13;		// submenu font size (pixels)	
var sfWEIGHT = 	"normal"	// submenu font weight
var stINDENT = 	5;		// submenu text indent (if text is left or right aligned)
var svPADDING = 2;		// submenu vertical cell padding
var svtOFFSET = 0;		// submenu vertical text offset (+/- pixels from middle)

var shSIZE =	2;		// submenu drop shadow size
var shCOLOR =	"#FFFFFF";	// submenu drop shadow color
var shOPACITY = 45;		// submenu drop shadow opacity (not ns4 or Opera 5)
var count=0
var path=""


//** LINKS ***********************************************************


// add main link item ("url","Link name",width,"text-alignment","target")

addMainItem("javascript:openpopup()","Logout",60,"center",""); 


	// define submenu properties (width,"align to edge","text-alignment")

	defineSubmenuProperties(140,"left","left");
	
	// add submenu link items ("url","Link name","target")
	// addSubmenuItem("http://www.dynamicdrive.com/new.htm","What\'s New","");
addMainItem("main.asp","Home",50,"center","app"); 

	defineSubmenuProperties(150,"left","left");	
addMainItem("um/um_gate.asp","Utility Manager",130,"center","app"); 
	defineSubmenuProperties(150,"right","left");
	addSubmenuItem("um/bldglist.asp","Bill Printer/Viewer","app");
	addSubmenuItem("billsummary/billsummary_select.asp","Bill Summary","app");
	addSubmenuItem("um/meterlist.asp","Meter List View","app");
	addSubmenuItem("um/meternotes.asp","Meter Problems","app");
	addSubmenuItem("um/lmsetup.asp","Meter LM Setup","app");
	//addSubmenuItem("um/validate.asp","Review/Edit+","app");
	addSubmenuItem("um/portfoliolist.asp","Building TC Setup","app");
	addSubmenuItem("um/tenantbilllist.asp","Bill Processor","app");
	addSubmenuItem("/um/validation/validation_select.asp","Review/Edit+","app");
	addSubmenuItem("/client_entry/entry.asp","Bill Entry","app");
	
addMainItem("","ERI Manager",120,"center",""); 
  
	defineSubmenuProperties(120,"right","left");
	addSubmenuItem("eri/tenantsetup.asp","Tenant Setup","app");
	addSubmenuItem("na.htm","MAC Adjustment","app");
	addSubmenuItem("na.htm","IBS Export","app");
	addSubmenuItem("eri/survey.asp","Survey","app");
	addSubmenuItem("na.htm","Reports","app");
	addSubmenuItem("na.htm","System Setup","app");
	addSubmenuItem("eri/libraryedit.asp","Library Edit","app");
	

addMainItem("","Operations Manager",140,"center","app");

	defineSubmenuProperties(150,"right","left");
	addSubmenuItem("/genergy2_intranet/opsmanager/joblog/frameset.html","Job Log","app");
	//addSubmenuItem("opslog/rfpindex.asp","RFP Log","app");
	addSubmenuItem("opslog/time.asp","TimeSheets","app");
	addSubmenuItem("opslog/poindex.asp","Purchase Orders","app");
	
addMainItem("","Power Capacity",140,"center","app");

	defineSubmenuProperties(150,"right","left");
	addSubmenuItem("pac/capindex.asp","Setup","app");
	

addMainItem("","Schedules",80,"center","app");

	defineSubmenuProperties(150,"right","left");
	addSubmenuItem("schedules/ItTasks.html","IT Project Schedule","app");
	
addMainItem("","Admin",60,"center",""); 

	defineSubmenuProperties(150,"right","left");
    
    addSubmenuItem("security/admin.asp","Employee Setup","app");
	addSubmenuItem("security/usrinfo.asp","Client Setup","app");
	addSubmenuItem("opslog/admininvoices.asp", "Invoices", "app");
	addSubmenuItem("opslog/adminpo.asp", "Purchase Orders", "app");
	addSubmenuItem("opslog/admintimesheet.asp", "Timesheets", "app");
	addSubmenuItem("http://dev.genergy.com/test/email/email.asp", "QuickMail!", "app");
		
	
addMainItem("","Corporate",60,"center",""); 

	defineSubmenuProperties(150,"right","left");
    addSubmenuItem("war/index.asp","Report Viewer","app");
    addSubmenuItem("corp/bomasearch.asp","BOMA Search","app");
	addSubmenuItem("crm/mktindex.asp","CRM Log","app");
	addSubmenuItem("http://209.10.53.121/ViewerFrame?Mode=Motion","View Office","app");
	//addSubmenuItem("http://209.10.53.117","View Network Room","app");
	//addSubmenuItem("http://64.152.63.236","View PQ Meter","app");
//********************************************************************

//**DO NOT EDIT THIS *****
}//***********************
//************************

