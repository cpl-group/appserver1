//UDMv2.1.1
//**DO NOT EDIT THIS *****
if (!exclude) { //********


var menuALIGN = "left";		// alignment
var absLEFT = 	0;		// absolute left or right position (if not center)
var absTOP = 	0; 		// absolute top position

var staticMENU = false;		// static positioning mode (not Opera 5)

var stretchMENU = true;		// show empty cells
var showBORDERS = true;		// show empty cell borders

var baseHREF =	"";	        // base path to .js and image files
var zORDER = 	100;		// base z-order of nav structure (not ns4)

var mCOLOR = 	"lightblue";	        // main nav cell color
var rCOLOR = 	"lightgreen";	// main nav cell rollover color
var keepLIT =	true;		// keep rollover color when browsing menu
var bSIZE = 	1;		// main nav border size
var bCOLOR = 	"black"	// main nav border color
var aLINK = 	"brown";	// main nav link color
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
var shCOLOR =	"#cccccc";	// submenu drop shadow color
var shOPACITY = 45;		// submenu drop shadow opacity (not ns4 or Opera 5)



//** LINKS ***********************************************************


// add main link item ("url","Link name",width,"text-alignment","target")

addMainItem("../logout.asp","Logout",80,"center",""); 


	// define submenu properties (width,"align to edge","text-alignment")

	defineSubmenuProperties(140,"left","left");
	
	// add submenu link items ("url","Link name","target")
	// addSubmenuItem("http://www.dynamicdrive.com/new.htm","What\'s New","");
	
addMainItem("","Utility Manager",140,"center",""); 

	defineSubmenuProperties(150,"right","left");
	
	addSubmenuItem("umsetup/umsetup.asp","Setup","_self");
	addSubmenuItem("http://www.genergy.com/um/eub.asp","Utility Bills","_blank");
	addSubmenuItem("http://www.genergy.com/um/manualreading.asp","Manual Readings","_blank");
	addSubmenuItem("http://www.genergy.com/um/r_e_readings.asp","Review / Edit","_blank");
	addSubmenuItem("http://www.genergy.com/um/validation.asp","Data Validation","_blank");


addMainItem("","ERI Manager",140,"center",""); 

	defineSubmenuProperties(150,"right","left");

	addSubmenuItem("../eri/tenantsetup.asp","Tenant Setup","");
	addSubmenuItem("../eri/macadj.asp","MAC Adjustment","");
	addSubmenuItem("../eri/ibsexport.asp","IBS Export","");
	addSubmenuItem("../eri/survey.asp","Survey","");
	addSubmenuItem("../eri/reports.asp","Reports","");
	addSubmenuItem("../eri/systemsetup.asp","System Setup","");
	addSubmenuItem("../eri/libraryedit.asp","Library Edit","");
	

addMainItem("","Operations Log",140,"center",""); 

	defineSubmenuProperties(135,"right","right");

	//addSubmenuItem("http://www.space.com/","Space.com","");
	

addMainItem("","TimeSheets",140,"center",""); 

	defineSubmenuProperties(150,"left","left");

//	addSubmenuItem("http://www.mrShowBiz.com","MrShowbiz","");



//********************************************************************

//**DO NOT EDIT THIS *****
}//***********************
//************************

