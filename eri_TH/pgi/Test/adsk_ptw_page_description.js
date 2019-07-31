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
    
    Template element id: adsk_ptw_page_description
    Publishing content:  page description
*/

adsk_ptw_page_description_main();

function adsk_ptw_page_description_main() {
    var xmle=adsk_ptw_xml.getElementsByTagName("publish_to_web").item(0);
    xmle=xmle.getElementsByTagName("header").item(0);
    xmle=xmle.getElementsByTagName("description").item(0);
    var e = document.getElementById("adsk_ptw_page_description");
    if (null == xmle.firstChild) {
        e.appendChild(document.createTextNode(""));
    } 
    else {
        e.appendChild(document.createTextNode(xmle.firstChild.data));
    }
}