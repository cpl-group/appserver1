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
    
    Template element id: adsk_ptw_array_of_thumbnails
    Publishing content:  image
                         label
                         description
                         idrop

    Template element id: adsk_ptw_summary_frame
    Publishing content:  drawing summary
*/

adsk_ptw_array_of_thumbnails_main();

function adsk_ptw_array_of_thumbnails_main() {

    var n=4; 
    var table_width="700"; 
    
    var xmle=adsk_ptw_xml.getElementsByTagName("publish_to_web").item(0);
    xmle=xmle.getElementsByTagName("contents").item(0);
    var xmles=xmle.getElementsByTagName("content");
    var e = document.getElementById("adsk_ptw_array_of_thumbnails");
    table = document.createElement("table");
    table.cellSpacing=10; 
    tbody = document.createElement("tbody");
    e.appendChild(table);
    table.appendChild(tbody);
    table.width="100%";
    tbody.appendChild(document.createElement("tr")); 

    var isDWF=false;
    for (i=0; i < xmles.length; i++) {
        if ((0==i) || (0 == (i % n))) { 
            tr = document.createElement("tr");
            tr2 = document.createElement("tr");
            tr3 = document.createElement("tr"); 
            tbody.appendChild(tr);
            tbody.appendChild(tr2);
            tbody.appendChild(document.createElement("tr"));
            tbody.appendChild(document.createElement("tr"));
            tr.align="left";
            tr.vAlign="top"
            tr2.align="left";
            tr2.vAlign="top"
        }  
        td = document.createElement("td"); 
        td2 = document.createElement("td"); 
        tr.appendChild(td);
        tr2.appendChild(td2);  
        td.width=table_width / n; 
        td2.width=table_width / n; 
        td2.align="left";
        td2.vAlign="top";
        content=xmles.item(i);     
        image=content.getElementsByTagName("image").item(0);
        a=document.createElement("a"); 
        td.appendChild(a);
        a.value=image.firstChild.data;
        a.href="javascript:adsk_ptw_array_of_thumbnails_onClickImage()";
        thumb=content.getElementsByTagName("thumbnail").item(0);
        img = document.createElement("image");
        a.appendChild(img);
        img.src = thumb.firstChild.data; 
        img.border=1; 
        img.onmouseover=function() { adsk_ptw_array_of_thumbnails_mouse_over(); }
        img.onmouseout=function() { adsk_ptw_array_of_thumbnails_mouse_out(); }
        img.name=i;

        idrop = content.getElementsByTagName("iDropXML").item(0);
        if (null != idrop.firstChild) {
            br = document.createElement("br");
            td.appendChild(br);
            activex = document.createElement("object");
            td.appendChild(activex);
            activex.codeBase=xmsg_adsk_ptw_all_idrop_url;
            activex.classid="clsid:21E0CB95-1198-4945-A3D2-4BF804295F78";
            activex.package=idrop.firstChild.text;
            activex.background="iDropButton.gif";
            activex.width="16";
            activex.height="16";
        }

        a2=document.createElement("a"); 
        td2.appendChild(a2);
        a2.className="DRAWING_LABEL";   
        a2.value=image.firstChild.data;
        a2.href="javascript:adsk_ptw_array_of_thumbnails_onClickImage()";

        if (!isDWF) {
            isDWF=adsk_ptw_validate_vview_is_dwf_file(image.firstChild.text);
        }

        title=content.getElementsByTagName("title").item(0);
        if (null == title.firstChild) {
            a2.appendChild(document.createTextNode(" "));
        }
        else {
            a2.appendChild(document.createTextNode(title.firstChild.text));
        }
        td2.appendChild(document.createElement("br"));

        div=document.createElement("div");
        td2.appendChild(div);
        div.className="DRAWING_DESCRIPTION";  
        desc = content.getElementsByTagName("description").item(0);
        if (null == desc.firstChild) {
            div.appendChild(document.createTextNode(""));
        }
        else {
            div.appendChild(document.createTextNode(desc.firstChild.data));
        }
    }
    if (isDWF) {
        activex = document.createElement("object");
        activex.classid="clsid:8718C658-8956-11D2-BD21-0060B0A12A50";
        adsk_ptw_validate_vview_gfunc(activex.Version);
    }
}

function adsk_ptw_array_of_thumbnails_onClickImage() {
    parent.window.navigate(document.activeElement.value);
}

function adsk_ptw_array_of_thumbnails_mouse_over() {
    if (null == parent) return;
    if (null == parent.adsk_ptw_summary_frame) return;

    i=window.event.srcElement.name;

    var xmle=adsk_ptw_xml.getElementsByTagName("publish_to_web").item(0);
    xmle=xmle.getElementsByTagName("contents").item(0);
    var xmles=xmle.getElementsByTagName("content");
    sum_info = xmles.item(i).getElementsByTagName("summary_info").item(0);

    if (null != sum_info) {
        body=parent.adsk_ptw_summary_frame.document.getElementsByTagName("body").item(0);

        title=sum_info.getElementsByTagName("title").item(0);
        adsk_ptw_array_of_thumbnails_summary(body, xmsg_adsk_ptw_all_summaryTitle, title);

        subject=sum_info.getElementsByTagName("subject").item(0);
        adsk_ptw_array_of_thumbnails_summary(body, xmsg_adsk_ptw_all_summarySubject, subject);

        author=sum_info.getElementsByTagName("author").item(0);
        adsk_ptw_array_of_thumbnails_summary(body, xmsg_adsk_ptw_all_summaryAuthor, author);

        keywords=sum_info.getElementsByTagName("keywords").item(0);
        adsk_ptw_array_of_thumbnails_summary(body, xmsg_adsk_ptw_all_summaryKeywords, keywords);

        comments=sum_info.getElementsByTagName("comments").item(0);
        adsk_ptw_array_of_thumbnails_summary(body, xmsg_adsk_ptw_all_summaryComments, comments);
        
        hyperlink_base=sum_info.getElementsByTagName("hyperlink_base").item(0);
        adsk_ptw_array_of_thumbnails_summary(body, xmsg_adsk_ptw_all_summaryHyperlinkBase, hyperlink_base);
    }
}

function adsk_ptw_array_of_thumbnails_mouse_out() {
    if (null == parent) return;
    if (null == parent.adsk_ptw_summary_frame) return;

    body=parent.adsk_ptw_summary_frame.document.getElementsByTagName("body").item(0);
    n=body.childNodes.length;
    for (i=0; i < n; i++) {
        body.removeChild(body.firstChild);
    }
}

function adsk_ptw_array_of_thumbnails_summary(rootNode, nameString, valueNode) {
    if (null == valueNode) return;
    if (null == valueNode.firstChild) return;

    b=parent.adsk_ptw_summary_frame.document.createElement("b");
    div=parent.adsk_ptw_summary_frame.document.createElement("div");
    rootNode.appendChild(div);
    div.appendChild(b);
    str = nameString + valueNode.firstChild.text;
    b.appendChild(parent.adsk_ptw_summary_frame.document.createTextNode(str));
}


