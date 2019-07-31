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
    
    Template element id: adsk_ptw_image_description
    Publishing content:  description

    Template element id: adsk_ptw_image
    Publishing content:  image

    Template element id: adsk_ptw_idrop
    Publishing content:  idrop

    Template element id: adsk_ptw_summary_frame
    Publishing content:  drawing summary    
*/

adsk_ptw_image_and_idrop_main();

function adsk_ptw_image_and_idrop_main() {

    var i=0; 

    var xmle=adsk_ptw_xml.getElementsByTagName("publish_to_web").item(0);
    xmle=xmle.getElementsByTagName("contents").item(0);
    var xmles=xmle.getElementsByTagName("content");

    dwg_img_desc=parent.adsk_ptw_image_frame.document.getElementById("adsk_ptw_image_description");
    desc = xmles.item(i).getElementsByTagName("description").item(0);
    p = document.createElement("p");
    if (null == desc.firstChild) {
        p.appendChild(document.createTextNode(""));
    }
    else {
        p.appendChild(document.createTextNode(desc.firstChild.text));
    }
    dwg_img_desc.appendChild(p);

    dwg_img=document.getElementById("adsk_ptw_image");
    var URL=document.location.href;
    var fileName=xmles.item(i).getElementsByTagName("image").item(0).firstChild.text;
    if (adsk_ptw_image_and_idrop_is_image_dwf(fileName)) {
        activex = document.createElement("object");
        dwg_img.appendChild(activex);
        activex.classid="clsid:8718C658-8956-11D2-BD21-0060B0A12A50";
        activex.src=(URL.substring(0, URL.lastIndexOf('/') + 1)) + fileName;
        activex.id="adsk_ptw_vve";
        activex.border="1";
        activex.width="500";
        activex.height="360";

        adsk_ptw_validate_vview_gfunc(activex.Version);
    }
    else {
        image = document.createElement("img");
        dwg_img.appendChild(image);
        image.src=fileName;
        image.border=1;
    }

    dwg_idrop = document.getElementById("adsk_ptw_idrop");
    idrop = xmles.item(0).getElementsByTagName("iDropXML").item(0);
    if (null != idrop.firstChild) {
        activex = document.createElement("object");
        dwg_idrop.appendChild(activex);
        activex.codeBase=xmsg_adsk_ptw_all_idrop_url;
        activex.classid="clsid:21E0CB95-1198-4945-A3D2-4BF804295F78";
        activex.package=idrop.firstChild.text;
        activex.background="iDropButton.gif";
        activex.width="16";
        activex.height="16";
    }

    if (null == parent.adsk_ptw_summary_frame) return;
    
    sum_info = xmles.item(i).getElementsByTagName("summary_info").item(0);
    if (null != sum_info) {
        body=parent.adsk_ptw_summary_frame.document.getElementsByTagName("body").item(0);

        title=sum_info.getElementsByTagName("title").item(0);
        adsk_ptw_image_and_idrop_summary(body, xmsg_adsk_ptw_all_summaryTitle, title);

        subject=sum_info.getElementsByTagName("subject").item(0);
        adsk_ptw_image_and_idrop_summary(body, xmsg_adsk_ptw_all_summarySubject, subject);

        author=sum_info.getElementsByTagName("author").item(0);
        adsk_ptw_image_and_idrop_summary(body, xmsg_adsk_ptw_all_summaryAuthor, author);

        keywords=sum_info.getElementsByTagName("keywords").item(0);
        adsk_ptw_image_and_idrop_summary(body, xmsg_adsk_ptw_all_summaryKeywords, keywords);

        comments=sum_info.getElementsByTagName("comments").item(0);
        adsk_ptw_image_and_idrop_summary(body, xmsg_adsk_ptw_all_summaryComments, comments);
        
        hyperlink_base=sum_info.getElementsByTagName("hyperlink_base").item(0);
        adsk_ptw_image_and_idrop_summary(body, xmsg_adsk_ptw_all_summaryHyperlinkBase, hyperlink_base);
    }
}

function adsk_ptw_image_and_idrop_summary(rootNode, nameString, valueNode) {
    if (null == valueNode) return;
    if (null == valueNode.firstChild) return;

    b=parent.adsk_ptw_summary_frame.document.createElement("b");
    div=parent.adsk_ptw_summary_frame.document.createElement("div");
    rootNode.appendChild(div);
    div.appendChild(b);
    str = nameString + valueNode.firstChild.text;
    b.appendChild(parent.adsk_ptw_summary_frame.document.createTextNode(str));
}

function adsk_ptw_image_and_idrop_is_image_dwf(file_name) {
    var ext = file_name.substring(file_name.lastIndexOf('.') + 1, (file_name.length));
    return("dwf" == ext);
}



