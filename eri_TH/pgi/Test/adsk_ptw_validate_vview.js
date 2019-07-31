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

    Volo View (Express) detection    
*/

parent.document.adsk_ptw_vv_validate = true;

function adsk_ptw_validate_vview_gfunc(version)
{ 
    if (parent != null)  {
        if (!parent.document.adsk_ptw_vv_validate) 
            return;

        parent.document.adsk_ptw_vv_validate = false;
    }

	if (!adsk_ptw_validate_vview_checkVoloViewVersion (version)) {
        parent.window.navigate(xmsg_adsk_ptw_all_validate_vview_url);
        return;
	}
	
 	return;
}

function adsk_ptw_validate_vview_checkVoloViewVersion(version)
{ 
	if ((version == null) || (version < 1.13)) {
		return !window.confirm(xmsg_adsk_ptw_all_validate_vview);
	}
	
	return true;
}

function  adsk_ptw_validate_vview_is_dwf_file(file_name) {
    var ext = file_name.substring(file_name.lastIndexOf('.') + 1, (file_name.length));
    return("dwf" == ext);
}
