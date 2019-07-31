<%option explicit%>
<!--METADATA TYPE="typelib" FILE="\Program Files\Common Files\System\ado\msado15.dll"-->
<!--#INCLUDE VIRTUAL="/genergy2/secure.inc"-->
<%
dim tenantnum, bldg
tenantnum = request("tenantnum")
if instr(tenantnum,".")>0 then
  bldg = split(tenantnum,".")(1)
  tenantnum = split(tenantnum,".")(0)
end if
dim fMsg, loggedin
fMsg = "Welcome."
loggedin = False

loadNewXML tenantnum 'Tweaked to get to next page 

'response.write getxmlusername()

'response.write session("xmlUserObj")
'response.end

' Determine which button the Web visitor clicked and take
' appropriate action.

if not isempty(Request.Form("tenantnum")) then
ProcessLogin
end if



Sub ProcessLogin()
  'Verify that the Web visitor submitted both an email address and a passwd.
  If isempty(tenantnum) Then
    fMsg ="Enter Tenant Number."
    Exit Sub    	
  End If
  
  If (bldg = "TG") or (bldg = "LGA") or (bldg = "PABT") or (bldg = "GWB") or (bldg = "BATH") or (bldg = "LT") or (bldg = "JFKCORR") or (bldg = "tg") or (bldg = "lga") or (bldg = "pabt") or (bldg = "gwb") or (bldg = "bath") or (bldg = "lt") or (bldg = "jfkcorr")  Then
    fMsg ="Tenant Number suspended."
    Exit Sub    	
  End If
  
  dim cnn1, rsVis, sql
  Set cnn1 = Server.CreateObject("ADODB.Connection")
  Set rsVis = Server.CreateObject("ADODB.Recordset")
  cnn1.Open getConnect(0,0,"dbCore")
 'sql = "SELECT * FROM "&makeIPUnionDB("tblleases","")&" t  WHERE (tenantnum= '" & tenantnum & "')"
  sql = "SELECT * FROM super_main WHERE bldgnum='"&bldg&"'"
  'SQL= "SELECT * FROM tblleases  WHERE tenantnum= '" & tenantnum & "'"
  dim ip, loggedin, user, path, pid, billlink
  'rsVis.open sql, getConnect(0,0,"dbCore")
 'response.write  getConnect(0,bldg,"Billing")
 'response.end
  rsVis.open sql, getConnect(0,bldg,"Billing")
  if not rsVis.eof then
    ip = rsVis("ip")
    pid = rsVis("pid")
 else
    fMsg = "System update ongoing, try again later."
    exit sub
  end if
  rsVis.close
 
  sql = "SELECT * FROM tblleases WHERE tenantnum='"&tenantnum&"'"
  
  rsVis.open sql, getConnect(0,bldg,"billing")
   
  If rsVis.EOF Then
    fMsg = "System update ongoing, try again later."
  exit sub
  Else
    path = "tenantpage.asp?tenantnum="&tenantnum&"&pid="&pid
    loggedin = True
    user = rsVis("tenantnum")
 
	End If
	rsVis.Close
     
     'response.write "here1"
     'response.end     
  DIM SERVERIP,PORT
  rsVis.open "SELECT location,SERVERIP FROM portfolio p, billtemplates bt WHERE bt.id=p.templateid AND p.id='"&pid&"'", cnn1
  if not rsVis.eof then 
  billlink = rsVis("location")
  SERVERIP = rsVis("SERVERIP")
  end if
  rsVis.close

 If loggedin Then
  ' response.write user
  ' response.end
	loadNewXML user
	IF SERVERIP <> "" THEN 
	PORT ="1433"
	else
	PORT =""
    END IF

	setBuilding bldg, ip ,null,"",PORT
    setKeyValue "bldg", bldg
    setKeyValue "billlink", billlink
    setKeyValue "pid", pid
    response.redirect path
  End If
End Sub
%>
<html>
<head>
<title>gEnergyOne Login</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="Stylesheet" href="http://appserver1.genergy.com/genergy2/styles.css" type="text/css">
    <style>
        body {
            font-family: "Lucida Sans Unicode", tahoma, serif;
            -webkit-font-smoothing: antialiased;
        }

        .container {
            max-width: 650px;
            margin: 25px auto 0 auto;
            padding-bottom: 0px;
        }

        .main-box {
            max-width: 800px;
            margin: 0 auto 0 auto;
            text-align: center;
            background: #ffffff;
            overflow: hidden;
            border-radius: 6px;
            box-shadow: 0 0 20px #000;
            border: 1px solid #d2492a;
        }

        .form-group {
            margin-top: 20px;
        }

        .error-message {
            font-size: 14px;
            color: #E33A3A;
            letter-spacing: 1px;
        }

        .logo-holder {
            padding-top: 50px;
        }

        .version {
            font-size: 10px;
            margin-top: 30px;
            margin-bottom: 5px;
            display: inline-block;
            color: gray;
        }

        .message-holder {
            width: 100%;
            padding: 1px 0 1px 0;
            background: #fff;
        }

        .message-holder h4 {
            font-size: 14px;
            letter-spacing: 1px;
            margin: 20px;
        }

        .form-holder {
            background: #f05330;
            padding-bottom: 50px
        }

        .form-group {
            display: inline-block;
            width: 625px;
            text-align: center;
            border-bottom: 1px solid gray;
        }

        .icon-holder {
            display: inline-block;
            position: relative;
            float: left;
            width: 30px;
            margin-top: 5px;
        }

        .input-control {
            display: inline-block;
            width: 80%;
            background: none;
            box-shadow: none;
            border: none;
            color: #fff;
            font-size: 22px;
            outline: none;
            float: left;
            margin-top: 8px;
            padding-left: 10px;
        }

        .input-control input:focus {
            box-shadow: none !important;
            border: none;
        }

        .input-control::-webkit-input-placeholder {
            color: #1e0a06;
        }

        .blue-btn {
            width: 150px;
            height: 50px;
            background: #36BFD6;
            border: none;
            color: #FFF;
            font-size: 24px;
            border-radius: 6px;
            cursor: pointer;
        }

        .blue-btn:hover {
            background: #4BD3EA;
        }

        .check-box-style {
            color: #fff;
            margin-top: 15px;
            margin-bottom: 25px;
            font-size: 12px;
        }

        .check-box-style input {
            margin-right: 10px;
        }

        @media (max-width: 500px) {
            .container {
                margin: 20px auto 0 auto;
            }

            .blue-btn {
                width: 80%;
            }
        }

        @media (max-width: 768px) {
            .container {
                margin: 20px auto 0 auto;
            }

            .blue-btn {
                width: 80%;
            }
        }
    </style>
</head>
<script language="javascript">
loaded = 0;
function preloadImg(){
  btnLoginOn = new Image(); btnLoginOn.src = "/images/login/login-1.gif";
  btnLoginOff = new Image(); btnLoginOff.src = "/images/login/login.gif";
  ResetOn = new Image(); ResetOn.src = "/images/login/reset-1.gif";
  ResetOff = new Image(); ResetOff.src = "/images/login/reset.gif";
  loaded = 1;
}

mywidth = screen.availWidth - 8;
myheight = screen.availHeight - 28;
function sizeandcenter(){
  desiredwidth = 580;
  desiredheight = 430;
  window.moveTo(((mywidth/2) - (desiredwidth/2)),((myheight/2) - (desiredheight/2))); 
  window.resizeTo(desiredwidth,desiredheight);
}
function processlogin(){
login.submit();
document.getElementById('progressbar').style.display = 'block';
document.getElementById('slideshow').style.display = 'none';
}
</script>
<body bgcolor="#FFFFFF" link="#000000" vlink="#000000" alink="#000000" leftmargin="0" topmargin="0">
<div class="container">
    <div class="main-box">
        <div class="logo-holder"><img src="images/logo.png" alt="logo"/></div>
        <!-- /.logo-holder -->
        <span class="version"></span>
        <div id="invalid" class="message-holder" hidden>
            <h4 class="error-message">Invalid username and password provided! please try again.</h4>
        </div>
        <!-- /.message-holder -->

        <div class="form-holder">
             <form name="form1" method="post" action="tlogin.asp">
                <div class="form-group">
                    <div class="icon-holder"><img src="images/user.png" alt="" width="15"/></div>
                    <!-- /.icon-holder -->
							<font size="5" face="Arial, Helvetica, sans-serif">
			Please Enter Your Tenant Access Code </br> As Seen On The Bottom Of Your Invoice 
		</font>

                    <input class="input-control" id="tenantnum" name="tenantnum" type="text" placeholder="tenantnum"/>
                </div>
                <!-- /.form-group -->


                <!-- /.form-group -->

                <div class="check-box-style remember-box">
                    &nbsp;
                </div>
                <!-- /.form-group -->
                <button id="submit" class="blue-btn">Login</button>
				
            </form>
        </div>
        <!-- /.form-holder -->
    </div>
    <!-- /.login-box -->
</div>
</br></br>
<font size="1" face="Arial, Helvetica, sans-serif">NOTE: 
                                  <a href="http://www.adobe.com/products/acrobat/readstep2.html" target="_blank">ADOBE 
                                  ACROBAT READER</a> IS REQUIRED TO VIEW INVOICES</font>
</br>
<font face="Arial, Helvetica, sans-serif" size="1" color="#000000">Our 
                                myEnergyPlatform offers tenants of buildings 
                                serviced by our Reading &amp; Billing Services 
                                instant online access to their current and historical 
                                invoices.</font>
								</br>
<font face="Arial, Helvetica, sans-serif" size="1" color="#000000">Should 
                                you experience problems with the login process, 
                                contact our IT department at 212 664 
                                7600 ext. 103, or send us an <a href="mailto:admin@cplgroupusa.com.com">email</a>. 
                                </font>
								</br>
<font face="Arial, Helvetica, sans-serif" size="1" color="#000000">Should 
                                you have questions concerning your account information, 
                                please contact our Reading &amp; Billing Department 
                                at 212 664 7600 ext. 137, or send us an <a href="mailto:rb@cplgroupusa.com">email</a>. 
                                </font>								
								
</body>
</html>
