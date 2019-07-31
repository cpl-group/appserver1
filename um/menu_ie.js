//UDMv2.1.1


// filter for undefined arrays 

for (f=0;f<mainItem.length;f++) {
	if (!subProps[f]) { subProps[f] = new Array(mainItem[f][2],mainItem[f][3],mainItem[f][3]); }
	if (!subItem[f]) { subItem[f] = new Array; }
	}


if (absTOP<0) { absTOP = 0; }
if (absLEFT<0) { absLEFT = 0; }
if (svPADDING<=0) { svPADDING=1; }

// appreviated base image path
var bh = baseHREF;


// find the inner height of the browser window

var bHeight = document.body.clientHeight;


// find background colors or images

var back_defs = new Array(mCOLOR,bCOLOR,rCOLOR,smCOLOR,sbCOLOR,srCOLOR,shCOLOR);
var useIMG = new Array(false,false,false,false,false,false)
var backers = new Array;
var mbackers = new Array;

for (b=0;b<back_defs.length;b++) {
	backers[b] = 'bgcolor=' + back_defs[b];
	if ((back_defs[b].indexOf('.gif') != -1) || (back_defs[b].indexOf('.jpg') != -1) || back_defs[b]=='') { useIMG[b] = true; }
	if (useIMG[b]) { 
		backers[b] = 'background="' + bh + back_defs[b] + '"'; 
		}
	if (back_defs[b]=='') { backers[b] = ''; } 
	}




// find the inner width of the browser window

var nWid = document.body.clientWidth;
var bWid = nWid;



// find the nav width and horizontal space

var endSpace = 0;
var navSpace = 0;
for (j=0;j<mainItem.length;j++) { 
	if (mainItem[j][2]=="") { mainItem[j][2]=10; }
	if (mainItem[j][1]=="") { mainItem[j][1]='&nbsp;'; 
		}
	navSpace+=mainItem[j][2]+bSIZE; 
	}
endSpace = bWid-navSpace-bSIZE;

if (bWid<navSpace) { menuALIGN="left"; absLEFT=0; stretchMENU=false; showBORDERS=false; }


// find the nav height

var navHeight = fSIZE+5+vPADDING;
var actualHeight = navHeight+(bSIZE*2); 


// find the subnav item height

var subnavHeight = sfSIZE+5+svPADDING;
var actualsubHeight = subnavHeight+(sbSIZE*2); 



// set values for different alignments

var stA=0; var absR; var relLEFT; 

var ra=false;
if (menuALIGN=="right") { 
	ra=true;
	absR = absLEFT; 
	absLEFT = endSpace-absR;
	relLEFT = absLEFT;
	if (relLEFT<0) { relLEFT=0; }
	absLEFT=0;
	}

var ca=false;
if (menuALIGN=="center") { 
	ca=true;
	absLEFT = endSpace / 2;
	relLEFT = absLEFT;
	if (relLEFT<0) { relLEFT=0; }
	absLEFT=0;
	}

var la=false;
if (menuALIGN=="left") {
	la = true;
	absR = absLEFT;
	relLEFT = absLEFT; 
	if (stretchMENU&&absR>0) { stA = absR+(2*bSIZE); }
	if (relLEFT<0) { relLEFT=0; }
	if (stretchMENU) { absLEFT=0; }
	}

var subLEFT = relLEFT;
if (ra) { subLEFT-=absR; }
if (la) { subLEFT+=absR; }
	
// how many main nav items
var ntl = 0;
for (intl=0;intl<mainItem.length;intl++) { if (mainItem[intl][1]!="") { ntl++; }}

var cSt = 'cursor:hand';


// do nothing
function doNothing() { }


// clear submenus

var previousId = 0;

function clearMenus(num) {
d.all["grid"].style.visibility = 'hidden';
if (vOFFSET>0) { d.all['gridblocker' + previousId].style.visibility = 'hidden'; }
if (shCOLOR!="") { d.all['shadow' + previousId].style.visibility = 'hidden'; }
d.all['subnav' + previousId].style.visibility = 'hidden'; 
d.all['sublinks' + previousId].style.visibility = 'hidden'; 
if (num!=previousId) { d.all['roll' + previousId].style.visibility = 'hidden'; }
previousId=0;
}



// open selected submenu

var rObj; var num;
var gridOkay = false;
function openMenu(num) {
//alert(num);
d.all["grid"].style.visibility = 'visible';
if (subItem[num]!="") { 
	if (vOFFSET>0) { d.all['gridblocker' + num].style.visibility = 'visible'; }
	d.all['roll' + num].style.visibility = 'visible'; 
	d.all['subnav' + num].style.visibility = 'visible'; 
	d.all['sublinks' + num].style.visibility = 'visible'; 
	if (shCOLOR!=""&&subItem[num]!='') { d.all['shadow' + num].style.visibility = 'visible'; }
	}
previousId = num;
}



//alert("assemble main nav");

var tSTR='';

// event capturing layer
tSTR+='<span id="grid" style="visibility:hidden\; position:absolute\; top:0\; left:0\; width:' + bWid + '\; height:' + bHeight + '\; z-index:' + zORDER + '" onmouseover="clearMenus()">&nbsp;</span>';


// nav stretching and event capturing layer
var stbSize = bSIZE;
tSTR+='<table cellpadding=0 cellspacing=' + stbSize + ' border=0 width=' + bWid + '  height=' + actualHeight + ' id="stretchnav" ';
if (showBORDERS) {
	if (useIMG[1]) { tSTR+='style="background-image:url(' + bh + bCOLOR + ')\; z-index:' + (zORDER+1) + '\; position:absolute\; top:' + absTOP + '\; left:0\;">'; }
	else { tSTR+='style="background:' + bCOLOR + '\; z-index:' + (zORDER+1) + '\; position:absolute\; top:' + absTOP + '\; left:0\;">'; }
	} else {
	tSTR+='style="z-index:' + (zORDER+1) + '\; position:absolute\; top:' + absTOP + '\; left:0\;">'; 
	stbSize = 0;
	}

tSTR+='<tr><td onmouseover="clearMenus()"></td></tr></table>';

if (stretchMENU&&showBORDERS) { 
	tSTR+='<table  id="stretchnav-sm-sb" cellpadding=0 cellspacing=' + stbSize + ' border=0 width=' + bWid + '  height=' + actualHeight + ' style="z-index:' + (zORDER+2) + '\; position:absolute\; top:' + absTOP + '\; left:0\;"><tr><td ' + backers[0] + ' onmouseover="clearMenus()"></td></tr></table>';
	}

if (stretchMENU&&!showBORDERS) { 
	tSTR+='<table  id="stretchnav-sm"cellpadding=0 cellspacing=' + stbSize + ' border=0 ' + backers[0] + ' width=' + bWid + '  height=' + actualHeight + ' style="z-index:' + (zORDER+2) + '\; position:absolute\; top:' + absTOP + '\; left:0\;"><tr><td onmouseover="clearMenus()"></td></tr></table>';
	}


// open frame

tSTR+='<table cellpadding=0 cellspacing=' + bSIZE + ' border=0 width=' + navSpace + ' height=' + actualHeight + ' id="mainnav"  ' + backers[1] + ' style="z-index:' + (zORDER+4) + '\; position:absolute\; top:' + absTOP + 'px\; left:' + relLEFT + 'px\;"><tr>'; 


// links
var cSt = new Array;
var linkHover = ' onmouseover="this.style.color=\''+aHOVER+'\'" onmouseout="this.style.color=\''+aLINK+'\'"';
if (aHOVER=='') { linkHover=''; }
for (i=0;i<mainItem.length;i++) { 
	var space = ' left:0px\;';
	if (mainItem[i][3]=="left") { space = ' left:' + tINDENT + 'px\;'; }
	if (mainItem[i][3]=="right") { space = ' left:-' + tINDENT + 'px\;'; }
	if(mainItem[i][1]!="") { 
		if (mainItem[i][0]=="") {  mainItem[i][0] = "javascript:doNothing()"; cSt[i] = 'cursor:default'; } else { cSt[i] = 'cursor:hand'; }
		tSTR+='<td onmouseover="d.all[\'roll' + i + '\'].style.visibility=\'visible\'\; clearMenus(' + i + ')\; openMenu(' + i + ')" class=menubarTD style="height:' + navHeight + 'px\;" ' + backers[0] + '><table cellpadding=0 cellspacing=0 border=0 width=' + mainItem[i][2] + ' style="z-index:' + (zORDER+4) + '\; height:' + navHeight + '\;"><tr><td align="' + mainItem[i][3] + '"><a href="' + mainItem[i][0] + '" target="'+ mainItem[i][4] + '" style="' + cSt + '\; position:relative\; top:' + (vtOFFSET) + '\; ' + space + '" ' + linkHover + '>' + mainItem[i][1] + '</a></td></tr></table></td>'; 
		}
		//alert(mainItem[i][4]) 
	}

// close frame
tSTR+='</tr></table>'; 



// rollover cells 
var rollLeft = relLEFT+bSIZE;

for (i=0;i<mainItem.length;i++) { 
	var space = ' left:0px\;';
	if (mainItem[i][3]=="left") { space = ' left:' + tINDENT + 'px\;'; }
	if (mainItem[i][3]=="right") { space = ' left:-' + tINDENT + 'px\;'; }
	if (mainItem[i][1]!="") {
		tSTR+='<table id="roll' + i + '" cellpadding=0 cellspacing=0 border=0 style="visibility:hidden\; z-index:' + (zORDER+5) + '\; height:' + navHeight + '\; position:absolute\; top:' + (absTOP+bSIZE) + 'px\; left:' + rollLeft + 'px\;" onmouseout="if(!keepLIT){this.style.visibility=\'hidden\'}"><tr><td class=menubarTD ' + backers[2] + '><table cellpadding=0 cellspacing=0 border=0 width=' + mainItem[i][2] + ' style="height:' + navHeight + '\;"><tr><td align="' + mainItem[i][3] + '"><a href="' + mainItem[i][0] + '" target="'+ mainItem[i][4]  + '" style="' + cSt[i] + '\; position:relative\; top:' + (vtOFFSET) + '\; ' + space + '" ' + linkHover + '>' + mainItem[i][1] + '</a></td></tr></table></td></tr></table>';
		}		
	rollLeft+=(mainItem[i][2]+bSIZE);
	}



//alert("assemble submenus");

var mSTR='';

var SUBabsLEFT=0;

for (count=0;count<mainItem.length;count++) {

	// text alignment and indentation
	var stAlign = subProps[count][2];
	var space = ' left:0px\;';
	if (stAlign=="left") { space = ' left:' + stINDENT + 'px\;'; }
	if (stAlign=="right") { space = ' left:-' + stINDENT + 'px\;'; }
	
	// find next submenu position
	if (count==0) { 
		SUBabsLEFT=relLEFT+bSIZE;
		}
	else { SUBabsLEFT+=mainItem[(count-1)][2]+bSIZE; }
	
	// specify edge alignemnt
	var actualLEFT = SUBabsLEFT+hOFFSET;
	//if ((SUBabsLEFT+subProps[count][0]+shSIZE)>bWid) { SUBabsLEFT-=shSIZE; }
	if (subProps[count][1]=="right") { 
		actualLEFT = SUBabsLEFT-(subProps[count][0]-mainItem[count][2])-hOFFSET;
		}
		
	// grid blocking image for vertical offset
	if (vOFFSET>0) { mSTR+='<span id="gridblocker' + count + '" style="width:' + subProps[count][0] + 'px\; height:' + vOFFSET + 'px\; visibility:hidden\; z-index:' + (zORDER+1) + '\;position:absolute\; top:' + (absTOP+actualHeight) + 'px\; left:' + actualLEFT + 'px\;"></span>'; }
		
	// container cells
	mSTR+='<table id="subnav' + count + '" cellpadding=0 cellspacing=' + sbSIZE + ' width="' + subProps[count][0] + '" border=0 ' + backers[4] + ' style="visibility:hidden\; z-index:' + (zORDER+7) + '\; position:absolute\; top:' + (absTOP+actualHeight+vOFFSET) + 'px\; left:' + actualLEFT + 'px\;">'; 
	var SUBabsTOP = absTOP+actualHeight+vOFFSET;
	for (i=0;i<subItem[count].length;i++) { 
		if (subItem[count][i][0]!=''&&subItem[count][i][1]!='') {

			mSTR+='<tr><td class=SUBmenubarTD ' + backers[3] + ' align="' + stAlign + '" width="' + (subProps[count][0]-(sbSIZE*2)) + '" style="height:' + subnavHeight + 'px\;"></td></tr>';
			}
		}
	mSTR+='</table>';

	// links
	var sublinkHover = ' onmouseover="this.style.color=\''+saHOVER+'\'" onmouseout="this.style.color=\''+saLINK+'\'"';
	if (saHOVER=='') { sublinkHover=''; }
	mSTR+='<table cellpadding=0 cellspacing=' + sbSIZE + ' width="' + subProps[count][0] + '" border=0 id="sublinks' + count + '" style="visibility:hidden\; z-index:' + (zORDER+9) + '\; position:absolute\; top:' + (absTOP+actualHeight+vOFFSET) + 'px\; left:' + actualLEFT + 'px\;">'; 
	var SUBabsTOP = absTOP+actualHeight+vOFFSET;
	for (i=0;i<subItem[count].length;i++) { 
		if (subItem[count][i][0]!=''&&subItem[count][i][1]!='') {
			mSTR+='<tr><td class=SUBmenubarTD height=' + subnavHeight + ' onmouseover="d.all[\'subroll' + count + i + '\'].style.visibility=\'visible\'" onmouseout="d.all[\'subroll' + count + i + '\'].style.visibility=\'hidden\'"><table cellpadding=0 cellspacing=0 border=0 id="sublink' + count + i + '" width=' + (subProps[count][0]-(sbSIZE*2)) + ' style="height:' + subnavHeight + '"><tr><td align="' + stAlign + '"><a href="' + subItem[count][i][0] + '" target="'+ subItem[count][i][2]  + '" style="position:relative\;top:' + (svtOFFSET) + '\; ' + space + '" ' + sublinkHover + '>' +  subItem[count][i][1] + '</a></td></tr></table></td></tr>';
			}
		}//
	mSTR+='</table>';
	
	// subrollover cells
	var subrollTop = absTOP+actualHeight+vOFFSET+sbSIZE;
	for (i=0;i<subItem[count].length;i++) { 
	if (subItem[count][i][0]!=''&&subItem[count][i][1]!='') {
	//alert(subItem[count][i][0])
		tSTR+='<table id="subroll' + count + i + '" cellpadding=0 cellspacing=0 border=0 width=' + (subProps[count][0]-(sbSIZE*2)) + ' ' + backers[5] + ' style="visibility:hidden\; z-index:' + (zORDER+8) + '\; height:' + subnavHeight + '\; position:absolute\; top:' + subrollTop + 'px\; left:' + (actualLEFT+sbSIZE) + 'px\;"><tr><td align="' + stAlign + '" class=SUBmenubarTD onmouseout="this.style.visibility=\'hidden\'"><a href="' + subItem[count][i][0] + '" target="'+ subItem[count][i][2]  + '" style="position:relative\; top:' + (svtOFFSET) + '\; ' + space + '" ' + sublinkHover + '>' +  subItem[count][i][1] + '</a></td></tr></table>';
		}		
	subrollTop+=subnavHeight+sbSIZE;
	}//



//alert("drop shadow"); 
if (shCOLOR!="") { mSTR+='<table cellpadding=0 cellspacing=0 border=0 ' + backers[6] + ' id="shadow' + count + '" style="visibility:hidden\; filter:alpha(opacity=' + shOPACITY + ')\; z-index:' + (zORDER+6) + '\; position:absolute\; top:' + (absTOP+actualHeight+vOFFSET+shSIZE) + 'px\; left:' + (actualLEFT+shSIZE) + 'px\; width:' + subProps[count][0] + '\; height:1\;"><tr><td>hello</td></tr></table>'; }

}


// resize / reload trap

window.onresize=new Function("window.location.reload()");




// static positioning properties from Dynamic Drive
// http://www.dynamicdrive.com/dynamicindex1/staticmenu2.htm

var staticObj;
function makendSpacetatic() {
d.all["grid"].style.pixelTop=d.body.scrollTop; 
d.all["stretchnav"].style.pixelTop=d.body.scrollTop+absTOP; 
if (stretchMENU&&showBORDERS) { d.all["stretchnav-sm-sb"].style.pixelTop=d.body.scrollTop+absTOP; }
if (stretchMENU&&!showBORDERS) { d.all["stretchnav-sm"].style.pixelTop=d.body.scrollTop+absTOP; }
d.all["mainnav"].style.pixelTop=d.body.scrollTop+absTOP; 
for (s=0;s<mainItem.length;s++) {
	d.all['subnav' + s].style.pixelTop=d.body.scrollTop+(absTOP+actualHeight+vOFFSET); 
	d.all['sublinks' + s].style.pixelTop=d.body.scrollTop+(absTOP+actualHeight+vOFFSET); 
	d.all['roll' + s].style.pixelTop=d.body.scrollTop+(absTOP+bSIZE); 
	subrollTop = absTOP+actualHeight+vOFFSET+sbSIZE;
	for (sr=0;sr<subItem[s].length;sr++) {
		d.all['subroll' + s + sr].style.pixelTop=d.body.scrollTop+subrollTop;
		subrollTop+=subnavHeight+sbSIZE;
		}
	if (subItem[num]!=""&&shCOLOR!="") { d.all['shadow' + s].style.pixelTop=d.body.scrollTop+(absTOP+actualHeight+vOFFSET+shSIZE); }
	if (vOFFSET>0) { d.all['gridblocker' + s].style.pixelTop=d.body.scrollTop+(absTOP+actualHeight); }
	}
setTimeout("makendSpacetatic()",0); 
}


// draw main nav 

d.write(tSTR);


// draw submenus

d.write(mSTR);



// find shadow heights

function findHeights() {
    var h=0
	if (shCOLOR!="") {
		for(h=0;h<mainItem.length;h++) {
			d.all['shadow' + h].style.height = d.all['subnav' + h].clientHeight;
			}
		}
	// static mode trigger
	if (staticMENU==true||staticMENU=="safe") { makendSpacetatic(); }
	}

window.onload=findHeights;



