<!--#include file="clsRSA.asp"-->
<html>
<head>

<title>Remote Query Analyzer</title>
<link rel="shortcut icon" href="images/_logo.ico" />
<link rel=stylesheet href="default.css" type="text/css">

<script language="JavaScript" src="cls/clsRSA.js"></script>
<script language="JavaScript" src="cls/clsBlowfish.js"></script>
<script>

	var bConnected = false;
	var bHResize = false;
	var bVResize = false;
	var nTime = -1;
	var intID = 0;

	var aAlign = new Array();
	var bCancel = false;
	var strCancel = "";

	var strTime = '';
	var strRows = ''

	var SQLServers = new Array();
	var SQLLogins = new Array();
	var SQLPasswords = new Array();

	var RSA2 = new clsRSA;
	var bEnc = false;

	var nKA = setTimeout('KeepAlive()',900000);

	<%
		if Session("RSA_M") = "" then
			set RSA = new clsRSA
			RSA.KeyGen
			Session("RSA_M") = RSA.KeyMod
			Session("RSA_D") = RSA.KeyDec
			Response.write "SetCookie('KeyEnc'," & RSA.KeyEnc & ");"&vbcrlf
			Response.write "SetCookie('KeyMod'," & RSA.KeyMod & ");"
			set RSA = Nothing
		end if
	%>

	function Encrypt(strText) {
		var RSA = new clsRSA;
		RSA.E = parseInt(GetCookie('KeyEnc',0));
		RSA.N = parseInt(GetCookie('KeyMod',0));
		return RSA.Enc(strText);
	}

	function KeepAlive() {
		//15min timer to keep session alive
		frmKeepAlive.location.replace('query.asp');
		nKA = setTimeout('KeepAlive()',900000);
	}

	function ResetAlive() {
		clearTimeout(nKA);
		nKA = setTimeout('KeepAlive()',900000);
	}

	function HResize() {
		if (event.button != 1) {bHResize=false; return;}
		if (bHResize) {
			Clear();
			var x = event.clientX;
			if (x < 50) {x = 50}
			if (x > document.body.clientWidth - 50) {x = document.body.clientWidth - 50}
			nList = tdList.offsetWidth;
			nOrig = divDetail.offsetWidth;
			NewX = document.body.clientWidth - x - 3;
			divDetail.style.width = NewX;
			tdQuery.style.width = NewX;
			divList.style.width = x - 3;
			NewX = nOrig + (nList - tdList.offsetWidth);
			divDetail.style.width = NewX;
			tdQuery.style.width = NewX;
			divList.style.width = tdList.offsetWidth;
		}
	}

	function VResize() {
		if (event.button != 1) {bVResize=false; return;}
		if (bVResize) {
			var y = 0;
			Clear();
			y = event.clientY;
			if (y < 50 + tblAuth.offsetHeight) {y = 50 + tblAuth.offsetHeight}
			if (y > document.body.clientHeight - 50) {y = document.body.clientHeight - 50}
			trQuery.style.height = y - tblAuth.offsetHeight - 3;
		}
	}

	function Clear() {
		document.selection.empty();
	}

	function WindowResize() {
		var obj = tblMain;
		var width = document.body.clientWidth;
		var height = document.body.clientHeight;
		var MarginTop = document.body.topMargin;
		var MarginLeft = document.body.leftMargin;
		if (typeof obj == 'object') {
			newWidth = width - obj.offsetLeft - MarginLeft;
			newHeight = height - obj.offsetTop - MarginTop;
			if (newWidth > 0) {obj.style.width = newWidth;}
			if (newHeight > 0) {obj.style.height = newHeight;}
		}
		var obj = divDetail;
		if (typeof obj == 'object') {
			newWidth = width - 6 - divList.offsetWidth;
			//newHeight = height - obj.offsetTop - MarginTop;
			if (newWidth > 0) {obj.style.width = newWidth;}
			//if (newHeight > 0) {obj.style.height = newHeight;}
		}

	}

	function clsButton() {

		//Make Functions Visible
		this.Init = Init;
		this.Action = Action;
		this.FixBorder = FixBorder;

		//Declare Functions
		function Init() {
			if (document.readyState != 'complete') {setTimeout('Button.Init()',100);return;}
			obj = document.all.tags('TD');
			for (x=0;x<obj.length;x++){
				if (obj[x].className.toLowerCase() == 'button') {
					obj[x].onmouseover = Button.Action;
					obj[x].onmouseout = Button.Action;
					obj[x].onmousedown = Button.Action;
					obj[x].onmouseup = Button.Action;
					obj[x].onmousemove = Button.Action;
					obj[x].innerHTML = '<div style=position:relative;>' + obj[x].innerHTML + '</div>'
				}
			}
		}

		function Action() {
			var obj = event.srcElement;
			if (event.button > 1){return}
			while (obj.tagName != 'TD') {obj = obj.parentElement;}
			switch (event.type) {
				case 'mouseover':
					if (obj.pressed) {break}
					obj.style.border='1 outset #ffffff';
					obj.style.borderLeft='1 solid #ffffff';
					obj.style.borderTop='1 solid #ffffff';
					break;
				case 'mouseout':
					if (obj.pressed) {break}
					obj.style.border='1 solid menu';
					SetPos(obj,0);
					break;
				case 'mousedown':
					FixBorder();
					obj.style.border='1 inset #ffffff';
					SetPos(obj,1);
					obj.pressed = true;
					break;
				case 'mouseup':
					if (obj.pressed) {break}
					obj.style.border='1 outset #ffffff';
					obj.style.borderLeft='1 solid #ffffff';
					obj.style.borderTop='1 solid #ffffff';
					SetPos(obj,0);
					break;
				case 'mousemove':
					break;
			}
		}

		function FixBorder() {
			var obj = -1;
			var x = 0;
			obj = tblAuth.all.tags('TD');
			for (x=0;x<obj.length;x++){
				if (obj[x].className.toLowerCase() == 'button') {
					if (obj[x].pressed) {
						obj[x].style.border='1 solid menu';
						SetPos(obj[x],0);
						obj[x].pressed = false;
					}
				}
			}
		}

		function SetPos(obj, pos) {
			obj.children[0].style.top=pos;
			obj.children[0].style.left=pos;
		}

		//Initialize Button Events
		Init();
	}

	//Initialize Class
	if (typeof Button == 'undefined'){Button = new clsButton()}

	function GetCookie(strName, strDefault) {
		// cookies are separated by semicolons
		var aCookie = document.cookie.split("; ");
		for (var i=0; i < aCookie.length; i++) {
			// a name/value pair (a crumb) is separated by an equal sign
			var aCrumb = aCookie[i].split("=");
			if (strName == aCrumb[0]) {return unescape(aCrumb[1])}
		}
		// a cookie with the requested name does not exist
		return strDefault;
	}

	function SetCookie(strName, strValue, strPath) {
		date = new Date();
		strDate = date.toGMTString().split(' ');
		strNewDate = '';
		strDate[3] = eval(date.getYear() + 1).toString();
		for (i=0; i < strDate.length; i++) {
			strNewDate += strDate[i] + ' ';
		}
		strPath = (typeof strPath == 'undefined')? '':'Path = ' + strPath + ";";
		document.cookie = strName + "=" + escape(strValue) + "; expires=" + strNewDate + ";"+strPath
	}

	function SetTmpCookie(strName, strValue, strPath) {
		strPath = (typeof strPath == 'undefined')? '':'Path = ' + strPath + ";";
		document.cookie = strName + "=" + escape(strValue) + ";" + strPath;
	}

	function DelCookie(strName) {
		document.cookie = strName + "= ''; expires=Fri, 31 Dec 1999 23:59:59 GMT;";
	}

	function FixTitle() {
		var x = 0;
		var strText = '';
		if (typeof trTitle == "object") {
			if (typeof trTitle.length == 'undefined') {
				trTitle.style.top=divDetail.scrollTop;
			} else {
				y = divDetail.scrollTop;
				for (x=0;x<trTitle.length;x++) {
					y1 = y;
					tbl = trTitle[x].parentNode.parentNode
					y3 = tbl.offsetTop
					if (y < y3) {y1=y3}
					y2 = y3 + tbl.offsetHeight - (trTitle[x].offsetHeight*2) + 2;
					if (y > y2) {y1 = y2}
					trTitle[x].style.top=y1-y3;
				}
			}
		}
	}

	function ExecSQL(bValidate) {
		bCancel = false;
		if (typeof selDB != 'object') {divDetail.innerHTML='<div style=\'padding:5 5;\'>Not connected to any server.</div>';return false;}
		var sql = '';
		if (document.selection.type=='Control'){
			sql = txtQuery.innerText;
		} else {
			obj.selection = document.selection.createRange();
			if (document.selection.type == 'Text'){
				sql = obj.selection.text;
			} else {
				sql = txtQuery.innerText;
			}
		}
		bEnc = chkEnc.checked;
		if (bEnc) {RSA2.KeyGen()}
		frmRequest.location.replace('query.asp?enc='+(bEnc?1:0)+'&e='+RSA2.E+'&n='+RSA2.N+'&i='+intID+'&s='+escape(txtServer.value)+'&l='+escape(txtLogin.value)+'&p='+escape(Encrypt(txtPass.value))+'&db='+escape(selDB.value)+'&a='+(bValidate?'v':'q')+'&sql='+escape(sql));

		nTime = 0;
		ShowExec(bValidate?3:2);
		UpdateTimer();
	}

	function StopSQL() {
		//var oHTTP = new ActiveXObject("Microsoft.XMLHTTP");
		//oHTTP.open("GET","cancel.asp?x="+escape(strCancel),false);
		//oHTTP.send();
		//window.status = oHTTP.responseText;
		bCancel = true;
		ShowExec(3);
	}

	function ShowExec(state) {
		bConnected = (state!=0);

		SetInput(txtServer,bConnected);
		SetInput(txtLogin,bConnected);
		SetInput(txtPass,bConnected);
		SetInput(selServer,bConnected);

		imgCon.style.display = (!bConnected)?'inline':'none';
		imgDisc.style.display = (bConnected)?'inline':'none';
		tdEnc.style.display = (bConnected)?'inline':'none';

		tdChk.style.display = (state==1)?'inline':'none';
		tdExec.style.display = (state==1)?'inline':'none';

		tdStop.style.display = (state==2)?'inline':'none';

		tdBlank1.style.display = (state==2 || state==3)?'inline':'none';
		tdBlank2.style.display = (state==3)?'inline':'none';
	}

	function Connect() {
		if (bConnected) {
			divList.innerHTML = '';
			tdDB.innerHTML='';
			Button.FixBorder();
			ShowExec(0);
			document.title = 'Remote Query Analyzer';
		} else {
			frmRequest.location.replace('query.asp?i='+intID+'&s='+escape(txtServer.value)+'&l='+escape(txtLogin.value)+'&p='+escape(Encrypt(txtPass.value))+'&a=l');
		}
	}

	function SetInput(obj, bDisabled) {
		obj.disabled = bDisabled;
		obj.style.backgroundColor = (bDisabled?'menu':'');
	}

	function CheckKey() {
		var obj = event.srcElement;
		var x = 0;
		if (event.keyCode == 69 && event.ctrlKey) {ExecSQL(false);return false;}
		if (event.keyCode == 69 && event.altKey) {ExecSQL(true);return false;}
		if (event.ctrlKey && event.shiftKey && event.altKey) {
			var obj = document.all('frmRequest')
			obj.style.display=(obj.style.display=='none')?'inline':'none';
		}
		if (event.keyCode != 9) {return true}
		if (document.selection.type=='Control'){return true}
		obj.selection = document.selection.createRange();
		if (document.selection.type == 'Text'){
			//Multi-Line Indenting
			if (obj.selection.text.search(/\n/) > 0) {

				//Find Start of Line
				tmp = obj.selection;
				obj.select();
				obj.selection = document.selection.createRange();
				x = obj.selection.offsetLeft;

				//Move selection to start of line
				obj.selection = tmp;
				while(x < obj.selection.offsetLeft){obj.selection.moveStart('character',-1)}

				//Unselect trailing linefeeds
				x = obj.selection.text.length - 1
				while(x < obj.selection.text.length){obj.selection.moveEnd('character',-1)}
				obj.selection.moveEnd('character',1)

				//Process Tabbing
				txt = obj.selection.text;
				if (event.shiftKey) {
					txt = txt.replace(/\n\t/g,'\n').replace(/^\t/,'');
				} else {
					txt = '\t'+txt.replace(/\n/g,'\n\t');
				}
				obj.selection.text = txt
				obj.selection.select();
				return false;
			}
		}
		//Insert single Tab
		obj.selection.text=String.fromCharCode(9);
		obj.selection.select();
		return false;
	}

	function UpdateTimer() {
		if (nTime >= 0) {
			setTimeout('UpdateTimer()',1000);
			var sec = nTime % 60;
			var min = (nTime - sec)/60;
			if (min < 10) {min = '0'+min}
			if (sec < 10) {sec = '0'+sec}
			strTime = '\['+min+':'+sec+'\]';
			nTime++;
			UpdateStatus();
		}
	}

	function UpdateStatus() {
		window.status = strTime + ' ' + strRows;
	}

	function ShowFolder(ID) {
		if (eval('typeof tr'+ID) == 'undefined') {return}
		var obj = eval('tr'+ID);
		obj.style.display = 'inline';
	}

	function FolderState(ID, State) {
		if (eval('typeof tr'+ID) == 'undefined') {return}
		var obj = eval('tr'+ID);
		var img = eval('img0_'+ID);
		img.style.display = (State)?'none':'inline';
		img = eval('img1_'+ID);
		img.style.display = (State)?'inline':'none';
		img = eval('img2_'+ID);
		img.style.display = (State)?'none':'inline';
		img = eval('img3_'+ID);
		img.style.display = (State)?'inline':'none';
		if (State) {
			frmRequest.location.replace('query.asp?i='+intID+'&id='+ID+'&s='+escape(txtServer.value)+'&l='+escape(txtLogin.value)+'&p='+escape(Encrypt(txtPass.value))+'&a=b'+'&o='+escape(obj.link));
		} else {
			obj.style.display = 'none';
		}
	}

	function SaveCon() {
		var x = 0;
		var strServer = txtServer.value;

		if (!bConnected) {return}

		for (x=0;x<SQLServers.length;x++) {
			if (SQLServers[x]==strServer){break}
		}

		SQLServers[x] = strServer;
		SQLLogins[x] = txtLogin.value;
		SQLPasswords[x] = txtPass.value;

		SetCookie('SQLServers',SQLServers);
		SetCookie('SQLLogins',SQLLogins);
		SetTmpCookie('SQLPasswords',SQLPasswords);
		fillOptions();
	}

	function GetCon() {
		SQLServers = GetCookie('SQLServers','').split(',');
		SQLLogins = GetCookie('SQLLogins','').split(',');
		SQLPasswords = GetCookie('SQLPasswords',',,,,,,,,,,,,,,,,,,,').split(',');
		fillOptions();
	}

	function fillOptions() {
		removeAllOptions(selServer);
		for (x=0;x<SQLServers.length;x++){
			if (SQLServers[x] != '') {
				addOption(selServer,SQLServers[x],SQLLogins[x],SQLPasswords[x],(SQLServers[x]==txtServer.value));
			}
		}
		if (selServer.length > 0) {GetInfo(selServer)}
	}

	function addOption(obj,text,value,pass,selected) {
		var x = obj.length;
		obj.options[x]=new Option(text,value);
		obj.options[x].pass = pass;
		obj.options[x].selected = selected;
	}

	function removeAllOptions(obj) {
		var len = obj.length;
		if(len <= 0) {return}
		for(i=0;i<len;i++) {obj.options[0]=null}
	}

	function GetInfo(obj) {
		var x = obj.selectedIndex;
		txtServer.value = obj.options[x].text;
		txtLogin.value = obj.options[x].value;
		txtPass.value = obj.options[x].pass;
	}

	function SetTitle(a1, a2){
		var x = 0;
		if (typeof aAlign != 'undefined') {x = aAlign[0] + 1}
		aAlign[0] = x;
		var strTBL = "<table id=tblDetail"+x+" border=0 cellpadding=0 cellspacing=0  style='padding:0 5;'>";
		strTBL += "<tr id=trTitle style='display:inline;position:relative;'>";
		strTBL += "<td class=ColumnTitle>&nbsp;</td>";
		for (var x=1; x < a1.length; x++) {
			aAlign[x] = a2[x];
			strTBL += "<td class=ColumnTitle align="+a2[x]+">"+a1[x]+"</td>";
		}
		strTBL += "</tr></table><br>";
		divDetail.innerHTML = strTBL;
	}

	function AddLine(a1){
		if (bCancel) {return true}
		var tbody = eval("tblDetail"+aAlign[0]+".all.tags('TBODY')[0]");
		var x = tbody.children.length;
		var tr = document.createElement("TR");
		var td = document.createElement("TD");
 		strRows = x + ' row(s) returned';
 		UpdateStatus();
		td.innerHTML = x;
		td.align = 'right';
		td.className = 'ColumnTitle';
		tr.appendChild(td);
		for (var x=1; x < a1.length; x++) {
			td = document.createElement("TD");
			td.noWrap = true;
			if (bEnc) {
				td.innerHTML = RSA2.Dec(a1[x]);
			} else {
				td.innerHTML = a1[x];
			}
			td.align = aAlign[x];
			tr.appendChild(td);
		}
 		tbody.appendChild(tr);
	}

	function Info() {
		var result = ''
		result += '<div style="padding:5 5;">';
		result += '<b>Remote Query Analyzer v1.01</b> (Not Yet Released)<br>';
		result += '<br>Author: <a href="mailto:benwhite@columbus.rr.com?subject=Remote File Explorer" >benwhite@columbus.rr.com</a><br>';
		result += '<br>ChangeLog: <a href="readme.txt" target=_new>ReadMe.txt</a><br>';
		result += '<br>Download: You can find the latest version of this code on <a href="http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=8193&lngWId=4" target=psc>www.PlanetSourceCode.com</a><br>';
		result += '<br>Donate $$$: If you would like to donate any money to the author, for who knows what reason --> <a href="https://www.paypal.com/xclick/business=benwhite%40columbus.rr.com&item_name=Remote+File+Explorer&no_note=1&tax=0&currency_code=USD" target=pp>PayPal</a><br>';
		result += '<br>Terms of Agreement:';
		result += '<br>1) You may use this code freely and with no charge.';
		result += '<br>2) You MAY NOT redistribute this code without written permission from me, the author.';
		result += '<br>3) You MAY NOT commercially use this code without written permission from me, the author.';
		result += '<br>4) This information must remain unchanged.';
		result += '<br><br>Failure to follow these guidlines is a violation of copyright laws.';
		result += '</div>';
		divDetail.innerHTML = result;
	}
</script>

</head>

<body leftmargin=0 topmargin=0 marginheight=0 marginwidth=0 scroll=no onmousemove='HResize();VResize();' onresize='WindowResize();'>

<div id=tblAuth class=ColumnTitle  style='padding:1 0;' height=100% width=100% >
	<table border=0 cellpadding=0 cellspacing=0  style='padding:0 5;'><tr>
			<td>Server</td>
			<td style='padding-left:0;'>
				<select id=selServer  class=InputSmall style='display:none;width:115;position:absolute;Clip:rect(auto auto auto 97px);' onchange='GetInfo(this);'>Select Server</select>
				<input id=txtServer class=InputSmall type=text style='position:relative;width:100;height:20;margin-right:15;padding-top:3px;' onkeypress='if (event.keyCode == 13){Connect();return false;}'>
			</td>
			<td>Login</td>
			<td style='padding-left:0;'><input id=txtLogin class=InputSmall type=text style='width:100;' onkeypress='if (event.keyCode == 13){Connect();return false;}'></td>
			<td>Password</td>
			<td style='padding-left:0;'><input id=txtPass class=InputSmall type=password style='width:100;' onkeypress='if (event.keyCode == 13){Connect();return false;}'></td>
			<td class=Button style='padding:1 3;' onclick='Connect()'><img id=imgCon src='images/_connect.gif' height=16 width=16 alt='Connect'><img id=imgDisc src='images/_disconnect.gif' height=16 width=16 alt='Disconnect' style='display:none;'></td>
			<td id=tdChk class=Button style='padding:1 3;display:none;' onclick='ExecSQL(true)'><img src='images/_check.gif' height=16 width=16 alt='Parse Query (ALT+E)'></td>
			<td id=tdBlank1 style='padding:1 4;display:none;'><img src='images/_check_dis.gif' height=16 width=16 alt='Parse Query (ALT+E)'></td>
			<td id=tdExec class=Button style='padding:1 3;display:none;' onclick='ExecSQL(false)'><img src='images/_execute.gif' height=16 width=16 alt='Execute Query (CTRL+E)'></td>
			<td id=tdStop class=Button style='padding:1 3;display:none;' onclick='StopSQL()'><img src='images/_stop.gif' height=16 width=16 alt='Stop Query (CTRL+S)'></td>
			<td id=tdBlank2 style='padding:1 4;display:none;'><img src='images/_execute_dis.gif' height=16 width=16 alt='Stop Query (CTRL+S)'></td>
			<td id=tdDB></td>
			<td id=tdEnc style='padding:1 4;display:none;' nowrap title='Results will be returned slower' style='cursor:default;'><input id=chkEnc type=checkbox class=InputCheck><span onclick='chkEnc.click();'> Encrypt Results<span></td>
	</tr></table>
</div>

<table id=tblMain border=0 cellpadding=0 cellspacing=0 height=100% width=100% >
<tr>
	<td id=tdList bgcolor=#ffffff valign=top>
		<table border=0 cellpadding=0 cellspacing=0 height=100% width=100% >
		<tr><td class=ColumnTitle style='padding:4 5;' height=20 nowrap><img src=images/_info.gif class=hand onclick='Info();' style='float:right' height=13 width=13>Object Browser</td></tr>
		<tr height=100% ><td><div id=divList style='padding:2 0; position:relative; width:100%; height:100%; overflow-y:auto; overflow-x:auto;'>&nbsp;</div></td></tr>
		</table>
	</td>
	<td bgcolor=menu class=ColumnTitle style='cursor:col-resize;' onmousedown='bHResize=true;' valign=top width=6><img src=images/_vertend.gif width=4 height=4 style=visibility:hidden;></td>
	<td id=tdQuery width=80%>

		<table border=0 cellpadding=0 cellspacing=0 height=100% width=100% >
		<tr id=trQuery height=30% valign=top><td><textarea id=txtQuery style='width:100%;height:100%;border:0 solid menu;font-family:terminal;font-size:6pt;position:relative;top:-1;' onkeydown='return CheckKey();'></textarea></td></tr>
		<tr><td bgcolor=menu class=ColumnTitle style='cursor:row-resize;' onmousedown='bVResize=true;' valign=top height=6><img src=images/_vertend.gif width=4 height=4 style=visibility:hidden;></td></tr>
		<tr height=100% ><td><div id=divDetail style='position:relative; width:100%; height:100%; overflow-x:auto; overflow-y:scroll;' onscroll='FixTitle()'></div></td></tr>
		</td></tr>
		</table>

	</td>
</tr>
</table>
<script>
	GetCon();
	divDetail.style.width = divDetail.offsetWidth;
	divList.style.width = tdList.offsetWidth;
	WindowResize();
</script>
<script Language=Javascript1.3>
	//added to support ie4 ... ugh
	selServer.style.display='inline';
	txtServer.style.top = -2;
</script>

<iframe id=frmRequest style='display:none;position:absolute;top:50;left:0;'></iframe>
<iframe id=frmKeepAlive style='display:none;position:absolute;top:50;left:0;'></iframe>
</body>
</html>
