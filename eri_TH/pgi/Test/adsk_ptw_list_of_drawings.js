/* 
	Copyright 1988-2000 by Autodesk, Inc.

	Permission to use, copy, modify, and distribute this software
	for any purpose and without fee is hereby granted, provided
	that the above copyright notice appears in all copies and
	that both that copyright notice and the limited warranty and
	restricted rights notice below appear in all supporting
	documentation.

	AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
	AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
	MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK, INC.
	DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
	UNINTERRUPTED OR ERROR FREE.

	Use, duplication, or disclosure by the U.S. Government is subject to
	restrictions set forth in FAR 52.227-19 (Commercial Computer
	Software - Restricted Rights) and DFAR 252.227-7013(c)(1)(ii) 
	(Rights in Technical Data and Computer Software), as applicable.

    Autodesk Publish to Web JavaScript 
    
    Template element id: adsk_ptw_list_of_drawings
    Publishing content:  label

    Template element id: adsk_ptw_image_description
    Publishing content:  description

    Template element id: adsk_ptw_image
    Publishing content:  image

    Template element id: adsk_ptw_idrop
    Publishing content:  idrop

    Template element id: adsk_ptw_summary_frame
    Publishing content:  drawing summary    
*/

adsk_ptw_list_of_drawings_main();

function adsk_ptw_list_of_drawings_main() {
    var xmle=adsk_ptw_xml.getElementsByTagName("publish_to_web").item(0);
    xmle=xmle.getElementsByTagName("contents").item(0);
    var xmles=xmle.getElementsByTagName("content");
    var e = document.getElementById("adsk_ptw_list_of_drawings");

    table = document.createElement("table");
    tbody = document.createElement("tbody");
    e.appendChild(table);
    table.appendChild(tbody);
    tr = document.createElement("tr");
    tbody.appendChild(tr);
    tr.align="left";
    tr.vAlign="top"

    td = document.createElement("td"); 
    tr.appendChild(td);
    table2=document.createElement("table"); 
    td.appendChild(table2);
    table2.cellPadding=1;
    table2.cellSpacing=5;

    tbody2=document.createElement("tbody");
    table2.appendChild(tbody2);
    td2 = document.createElement("td"); 
    tr.appendChild(td2);
    for (i=0; i < xmles.length; i++) {
        content=xmles.item(i);      
        title = content.getElementsByTagName("title").item(0);
        a=document.createElement("a"); 
        if (null == title.firstChild) {
            a.appendChild(document.createTextNode(" "));
        }
        else {
            a.appendChild(document.createTextNode(title.firstChild.text));
        }
        table2_tr=document.createElement("tr");
        tbody2.appendChild(table2_tr);
        table2_td=document.createElement("td");
        table2_tr.appendChild(table2_td);
        table2_td.appendChild(a);
        a.className="DRAWING_LABEL";
        var fileName = content.getElementsByTagName("image").item(0).firstChild.text;
        a.id=fileName;
        a.value=i;   
        if (adsk_ptw_list_of_drawings_is_image_dwf(fileName)) {  
            a.href="javascript:adsk_ptw_list_of_drawings_onClickVVE()";
        } else {
            a.href="javascript:adsk_ptw_list_of_drawings_onClickImage()";
        }
    }
}

function adsk_ptw_list_of_drawings_createVVEControl (i) {
    dwg_img_desc=parent.adsk_ptw_image_frame.document.getElementById("adsk_ptw_image_description");
	dwg_img_desc.removeChild(dwg_img_desc.firstChild);

    var xmle=adsk_ptw_xml.getElementsByTagName("publish_to_web").item(0);
    xmle=xmle.getElementsByTagName("contents").item(0);
    xmles=xmle.getElementsByTagName("content");
    desc = xmles.item(i).getElementsByTagName("description").item(0);
    p = parent.adsk_ptw_image_frame.document.createElement("p");
    if (null == desc.firstChild) {
        p.appendChild(parent.adsk_ptw_image_frame.document.createTextNode(""));
    }
    else {
        p.appendChild(parent.adsk_ptw_image_frame.document.createTextNode(desc.firstChild.text));
    }
    dwg_img_desc.appendChild(p);

    dwg_img=parent.adsk_ptw_image_frame.document.getElementById("adsk_ptw_image");
	dwg_img.removeChild(dwg_img.firstChild); 
    activex = parent.adsk_ptw_image_frame.document.createElement("object");
    dwg_img.appendChild(activex);
    activex.classid="clsid:8718C658-8956-11D2-BD21-0060B0A12A50";

    var URL=document.location.href;
    var fileName=xmles.item(i).getElementsByTagName("image").item(0).firstChild.text;
    activex.src=(URL.substring(0, URL.lastIndexOf('/') + 1)) + fileName;
    activex.id="adsk_ptw_vve";
    activex.border="1";
    activex.width="500";
    activex.height="360";

    adsk_ptw_validate_vview_gfunc(activex.Version);

    adsk_ptw_list_of_drawings_setiDrop(i);
}

function adsk_ptw_list_of_drawings_onClickVVE() {
    adsk_ptw_list_of_drawings_createVVEControl (document.activeElement.value);
    adsk_ptw_list_of_drawings_set_summary_info(document.activeElement.value);
}

function adsk_ptw_list_of_drawings_createImageElement(i) {
    dwg_img_desc=parent.adsk_ptw_image_frame.document.getElementById("adsk_ptw_image_description");
	dwg_img_desc.removeChild(dwg_img_desc.firstChild);

    var xmle=adsk_ptw_xml.getElementsByTagName("publish_to_web").item(0);
    xmle=xmle.getElementsByTagName("contents").item(0);
    var xmles=xmle.getElementsByTagName("content");
    desc = xmles.item(i).getElementsByTagName("description").item(0);
    p = parent.adsk_ptw_image_frame.document.createElement("p");
    if (null == desc.firstChild) {
        p.appendChild(parent.adsk_ptw_image_frame.document.createTextNode(""));
    }
    else {
        p.appendChild(parent.adsk_ptw_image_frame.document.createTextNode(desc.firstChild.text));
    }
    dwg_img_desc.appendChild(p);

    dwg_img=parent.adsk_ptw_image_frame.document.getElementById("adsk_ptw_image");
	dwg_img.removeChild(dwg_img.firstChild); 
    image = parent.adsk_ptw_image_frame.document.createElement("img");
    dwg_img.appendChild(image);
    image.src=xmles.item(i).getElementsByTagName("image").item(0).firstChild.text;
    image.border=1;

    adsk_ptw_list_of_drawings_setiDrop(i);
}

function adsk_ptw_list_of_drawings_onClickImage() {
    adsk_ptw_list_of_drawings_createImageElement(document.activeElement.value);
    adsk_ptw_list_of_drawings_set_summary_info(document.activeElement.value);
}

function adsk_ptw_list_of_drawings_setiDrop(i) {
    dwg_idrop = parent.adsk_ptw_image_frame.document.getElementById("adsk_ptw_idrop");
    if (null != dwg_idrop.firstChild) {
	    dwg_idrop.removeChild(dwg_idrop.firstChild); 
    }
    var xmle=adsk_ptw_xml.getElementsByTagName("publish_to_web").item(0);
    xmle=xmle.getElementsByTagName("contents").item(0);
    var xmles=xmle.getElementsByTagName("content");
    idrop = xmles.item(i).getElementsByTagName("iDropXML").item(0);
    if (null != idrop.firstChild) {
        activex = parent.adsk_ptw_image_frame.document.createElement("object");
        dwg_idrop.appendChild(activex);
        activex.codeBase=xmsg_adsk_ptw_all_idrop_url;
        activex.classid="clsid:21E0CB95-1198-4945-A3D2-4BF804295F78";
        activex.package=idrop.firstChild.text;
        activex.background="iDropButton.gif";
        activex.width="16";
        activex.height="16";
    }
}

function adsk_ptw_list_of_drawings_set_summary_info(i) {
    if (null == parent) return;
    if (null == parent.adsk_ptw_summary_frame) return;

    body=parent.adsk_ptw_summary_frame.document.getElementsByTagName("body").item(0);
    n=body.childNodes.length;
    for (index=0; index < n; index++) {
        body.removeChild(body.firstChild);
    }

    var xmle=adsk_ptw_xml.getElementsByTagName("publish_to_web").item(0);
    xmle=xmle.getElementsByTagName("contents").item(0);
    var xmles=xmle.getElementsByTagName("content");
    sum_info = xmles.item(i).getElementsByTagName("summary_info").item(0);

    if (null != sum_info) {
        body=parent.adsk_ptw_summary_frame.document.getElementsByTagName("body").item(0);

        title=sum_info.getElementsByTagName("title").item(0);
        adsk_ptw_list_of_drawings_summary(body, xmsg_adsk_ptw_all_summaryTitle, title);

        subject=sum_info.getElementsByTagName("subject").item(0);
        adsk_ptw_list_of_drawings_summary(body, xmsg_adsk_ptw_all_summarySubject, subject);

        author=sum_info.getElementsByTagName("author").item(0);
        adsk_ptw_list_of_drawings_summary(body, xmsg_adsk_ptw_all_summaryAuthor, author);

        keywords=sum_info.getElementsByTagName("keywords").item(0);
        adsk_ptw_list_of_drawings_summary(body, xmsg_adsk_ptw_all_summaryKeywords, keywords);

        comments=sum_info.getElementsByTagName("comments").item(0);
        adsk_ptw_list_of_drawings_summary(body, xmsg_adsk_ptw_all_summaryComments, comments);
        
        hyperlink_base=sum_info.getElementsByTagName("hyperlink_base").item(0);
        adsk_ptw_list_of_drawings_summary(body, xmsg_adsk_ptw_all_summaryHyperlinkBase, hyperlink_base);
    }
}

function adsk_ptw_list_of_drawings_summary(rootNode, nameString, valueNode) {
    if (null == valueNode) return;
    if (null == valueNode.firstChild) return;

    b=parent.adsk_ptw_summary_frame.document.createElement("b");
    div=parent.adsk_ptw_summary_frame.document.createElement("div");
    rootNode.appendChild(div);
    div.appendChild(b);
    str = nameString + valueNode.firstChild.text;
    b.appendChild(parent.adsk_ptw_summary_frame.document.createTextNode(str));
}

function adsk_ptw_list_of_drawings_is_image_dwf(file_name) {
    var ext = file_name.substring(file_name.lastIndexOf('.') + 1, (file_name.length));
    return("dwf" == ext);
}
