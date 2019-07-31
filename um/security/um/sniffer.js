//UDMv2.1.1


// browser detection based on Ultimate Client Sniffer, from 
// http://developer.netscape.com/docs/examples/javascript/browser_type.html
//********************************************************************************
    var agt=navigator.userAgent.toLowerCase();
    var is_major = parseInt(navigator.appVersion);
    var is_minor = parseFloat(navigator.appVersion);
    var is_nav  = ((agt.indexOf('mozilla')!=-1) && (agt.indexOf('spoofer')==-1)
                && (agt.indexOf('compatible') == -1) && (agt.indexOf('opera')==-1)
                && (agt.indexOf('webtv')==-1));
    var is_nav2 = (is_nav && (is_major == 2));
    var is_nav3 = (is_nav && (is_major == 3));
    var is_nav4 = (is_nav && (is_major == 4));
    var is_nav4up = (is_nav && (is_major >= 4));
    var is_navonly      = (is_nav && ((agt.indexOf(";nav") != -1) ||
                          (agt.indexOf("; nav") != -1)) );
    var is_nav5 = (is_nav && (is_major == 5));
    var is_nav5up = (is_nav && (is_major >= 5));
    var is_moz7 = (is_nav && (agt.indexOf('netscape6')==-1));
    var is_ie   = (agt.indexOf("msie") != -1);
    var is_ie3  = (is_ie && (is_major < 4));
    var is_ie4  = (is_ie && (agt.indexOf("msie 4")!=-1));
    var is_ie5  = (is_ie && (agt.indexOf("msie 5")!=-1));
    var is_ie5up  = (is_ie  && !is_ie3 && !is_ie4);
    var is_opera = (agt.indexOf("opera") != -1);
    var is_opera4 = (is_opera && (agt.indexOf("opera 4")!=-1));
    var is_opera5up = ((is_opera && ((agt.indexOf("opera 5.11")!=-1) || (agt.indexOf("opera 5.02")!=-1) || (agt.indexOf("opera 5.01")!=-1))) || (is_opera && (is_major > 4)));
    var is_icab    = (agt.indexOf("icab")!=-1);
    var is_webtv = (agt.indexOf("webtv") != -1);
    
    var is_linux = (agt.indexOf("inux")!=-1);
    var is_mac    = (agt.indexOf("mac")!=-1);
    var is_win   = ( (agt.indexOf("win")!=-1) || (agt.indexOf("16bit")!=-1) );
//********************************************************************************


// if client settings are not specific, treat as old browser (no script) on windows

var bType = "old";
var osType = "win";


// find browser

if (is_ie3 == true) { bType = "ie3"; }
if (is_ie4 == true && is_opera == false && is_webtv == false) { bType = "ie4"; }
if (is_ie5up == true && is_opera == false && is_webtv == false) { bType = "ie5"; }
if (is_nav3 == true) { bType = "ns3"; }
if (is_nav4up == true && is_nav5up == false) { bType = "ns4"; }
if (is_nav5up == true && is_moz7 == false) { bType = "ns6"; }
if (is_nav5up == true && is_moz7 == true) { bType = "mz7"; }
if (is_opera == true) {  bType = "op3"; }
if (is_opera4 == true) {  bType = "op4"; }
if (is_opera5up == true) {  bType = "op5"; }
if (is_webtv == true) {  bType = "tv"; }
if (is_icab == true) { bType = "ic"; }


// find os

if (is_linux==true) { osType = "lnx"; }
if (is_mac==true) { osType = "mac"; }
if (is_win==true) { osType = "win"; }


//create a set of handy variables

var ie3=false;var ie4=false;var ie5=false; var ie=false;
var ns3=false;var ns4=false;var ns6=false; var mz7=false;
var op3=false;var op4=false;var op5=false; var op=false; 
var ic=false;
var tv=false;var old=false;var exclude=false;

var lnx=false; var mac=false; var win=false; 

if (bType == "ie3") ie3 = true;if (bType == "ie4") ie4 = true; if (bType == "ie5") ie5 = true;
if (bType == "ns3") ns3 = true;if (bType == "ns4") ns4 = true; if (bType == "ns6") ns6 = true;
if (bType == "mz7") mz7 = true; 
if (bType == "op3") op3 = true;if (bType == "op4") op4 = true; if (bType == "op5") op5 = true;
if (bType == "ic") { ic = true; }
if (bType == "tv") tv = true; if (bType == "old") old = true;

if (osType=="win") { win = true; }
if (osType=="mac") { mac = true; }
if (osType=="lnx") { lnx = true; }


// create some browser groups

if (ie4 || ie5) { ie = true; }
if (op3 || op4 || op5) { op = true; }
if ((ie3==true) || (mac&&op) || (ns3==true) || op3 || op4 || ic || tv || old) { exclude = true; }

// array building functions for custom.js

var m=0;
var sm=0;
var cm=0;
var mainItem = new Array;

function addMainItem(ma,mb,mc,md,me) { 
sm=0;
mainItem[m] = new Array(ma,mb,mc,md,me);
if (mainItem[m][4]=="") { mainItem[m][4]="_self"; }

m++;
}

var sp=0;
var subProps = new Array;

function defineSubmenuProperties(spa,spb,spc) {
subProps[(m-1)] = new Array(spa,spb,spc);
}

var subItem = new Array;

function addSubmenuItem(sma,smb,smc) {
if (sm==0) { subItem[(m-1)] = new Array; }
subItem[(m-1)][sm] = new Array(sma,smb,smc);
if (subItem[(m-1)][sm][0]=="") { subItem[(m-1)][sm][0]="#"; }
if (subItem[(m-1)][sm][1]=="") { subItem[(m-1)][sm][1]="&nbsp;"; }
if (subItem[(m-1)][sm][2]=="") { subItem[(m-1)][sm][2]="_self"; }
ary=subItem[(m-1)][sm][0].split("/")
//for(i=0;i<ary.length-1;i++){
//    j=i+1
	//alert(ary[ary.length-3])
//	alert(ary[ary.length-2])
//    if(ary[i] == ary[j]){
//    	alert(subItem[(m-1)][sm][0])
//    }
//}
//alert(subItem[(m-1)][sm][0])
//alert(subItem[(m-1)][sm][1])
//alert(subItem[(m-1)][sm][2])
sm++;
}
