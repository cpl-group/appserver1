/***********************************************************************************
** GENERGY ONLINE JAVASCRIPT                                                      **
** Filename: genergyonline.js                                                     **
** General Energy Services, Inc., Copyright 2008                                  **
** The contents of this file my not be copied, duplicated, or redistributed       **
** in any form without the prior written consent of General Energy Services, Inc. **
************************************************************************************/

//Run Application
function Run(app)
{
	alert("Application : "+app);
}

//Load Page
function LoadPage(url)
{
	parent.document.frames.main.location = url
}

