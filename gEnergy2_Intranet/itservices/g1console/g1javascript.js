// JScript source code
function fillup(name){
  //document.location="usrdetail.asp?username="+name
  //document.site.src="usrsite.asp?username="+name
  alert(document.site.src);
}
function submitform(choice){
    if (document.form1.isDirty.value){
        if (document.form1.passwd.value != document.form1.Repasswd.value) {
            dispayWarning();
            return false;
        }
    }
	document.form1.choice.value=choice
	document.form1.submit()
}
function viewconsole(mode){
  document.site.location="./g1nav/g1nav.asp?mode="  + mode;
}
if (screen.width > 1024) {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/largestyles.css\" type=\"text/css\">")
} else {
  document.write ("<link rel=\"Stylesheet\" href=\"/gEnergy2_Intranet/styles.css\" type=\"text/css\">")
}
function cloneid(userid){
	var url
	url = "cloneid.asp?uid=" + userid
	openwin(url, 400,150)

}
function launchoptions(uid){

	var selection = document.form1.glboptList.value
	var cid = selection.split('_')[0];
	var label = selection.split('_')[1];
	var url = '/um/security/optionsList.asp?username='+uid+'&csid='+cid+'&label=' + label;
	openwin(url, 250,400)
}
function openwin(url,mwidth,mheight){
window.open(url,"","statusbar=no, menubar=no, HEIGHT="+mheight+", WIDTH="+mwidth)
}
function setDirty(){
    document.form1.isDirty.value = true;
    var lbl = document.getElementById("repass_lbl");
    var txt = document.getElementById("repass_txt");
    lbl.style.display = "block";
    txt.style.display = "block";
}
function dispayWarning(){
    var obj = document.getElementById("warning_lbl");
    obj.style.display = "block";
}