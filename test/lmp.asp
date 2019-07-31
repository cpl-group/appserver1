<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN"
        "http://www.w3.org/TR/1999/REC-html401-19991224/loose.dtd">
<html>
<head>
	<title>Untitled</title>
  <link rel="Stylesheet" href="/genergy2/styles.css" type="text/css">
  <script language="javascript">
  </script>
</head>
<body bgcolor="#336699">

<form name="buttons">
<table border=0 cellpadding="3" cellspacing="0" bgcolor="#ffffff" width="100%">
<tr>
  <td bgcolor="#336699" style="border:1px solid #000000;padding:0px;margin:0px;">
  <!-- begin LMP button bar -->
  <table border=0 cellpadding="0" cellspacing="0" width="100%" style="border-bottom:1px solid #000000;border-right:1px solid #000000;">
  <tr valign="middle" bgcolor="#003366">
    <td align="right" style="border-bottom:1px solid #000000;" colspan="7">
    <table border=0 cellpadding="0" cellspacing="0">
    <tr valign="middle">
      <td><input type="image" src="images/show_prefs-4.gif" id="prefstoggle" onclick="return togglePrefs();" border="0"></td>
      <td><a href="#"><img src="images/print_pdf2.gif" border="0"></a></td>
    </tr>
    </table>
    </td>
    <td width="5" style="border-bottom:1px solid #000000;">&nbsp;</td>
  </tr>
  <tr valign="middle">
    <td align="center" width="35" style="border-top:1px solid #99ccff;border-right:1px solid #000000;"><a href="#"><img src="images/pop_cal_sm.gif" alt="Pop Up Calendar" title="Pop Up Calendar" width="25" height="22" hspace="10" border"0" border="0"></a></td>
    <td align="center" style="border-top:1px solid #99ccff;border-left:1px solid #99ccff;border-right:1px solid #000000;">
    <table border=0 cellpadding="2" cellspacing="0">
    <tr valign="middle">
      <td><a href="#"><img src="images/prev_year_aro.gif" alt="&lt;&lt; Previous Year" title="Previous Year" width="16" height="22" border="0"></a></td>
      <td><a href="#"><img src="images/prev_mo_aro.gif" alt="&lt; Previous Month" title="Previous Month" width="16" height="22" border="0"></a></td>
      <td>
      <select name="month">
      <option>Month</option>
      <option value="1">Jan</option>
      <option value="2">Feb</option>
      <option value="3">Mar</option>
      <option value="8">Aug</option>
      <option value="9" selected>Sep</option>
      <option value="10">Oct</option>
      </select>
      </td>
      <td>
      <select name="day">
      <option>Day</option>
      <option value="all">All</option>
      <option value="1">1</option>
      <option value="2">2</option>
      <option value="3">3</option>
      <option value="4">4</option>
      <option value="20" selected>20</option>
      <option value="27">27</option>
      <option value="28">28</option>
      <option value="29">29</option>
      <option value="30">30</option>
      <option value="30">31</option>
      </select>
      </td>
      <td>
      <select name="year">
      <option>Year</option>
      <option value="2002" selected>2002</option>
      </select>
      </td>
      <td><a href="#"><img src="images/next_mo_aro.gif" alt="&lt; Next Month" title="Next Month" width="16" height="22" border"0" border="0"></a></td>
      <td><a href="#"><img src="images/next_year_aro.gif" alt="&lt;&lt; Next Year" title="Next Year" width="16" height="22" border="0"></a></td>
    </tr>
    </table>    
    </td>
    <td align="center" style="border-top:1px solid #99ccff;border-left:1px solid #99ccff;">
    <select onchange="switchLMP(this.selectedIndex);">
    <option>Interval
    <option selected>Hourly
    <option>15 minute
    <option>1 minute
    </select>
    </td>
    <!--[[td style="border-top:1px solid #99ccff;border-right:1px solid #000000;"]][[input type="image" src="images/go.gif" border="0"]][[/td]]-->      
    <td align="center" style="border-top:1px solid #99ccff;border-right:1px solid #000000;">
    <select>
    <option>Utility
    <option selected>Electricity
    <option>Gas
    <option>Steam
    <option>Chilled Water
    </select>
    </td>
    <!--[[td style="border-top:1px solid #99ccff;"]][[input type="image" src="images/go.gif" border="0"]][[/td]]-->
     <td align="center" style="border-top:1px solid #99ccff;border-left:1px solid #99ccff;">
      <select onchange="parent.location='projected_frameset.html';">
      <option>Cost
      <option>Actual Cost
      <option>DAM Cost
      <option>Full Service Rate SC4 1
      <option>Full Service Rate SC4 2
      <option>Full Service Rate SC9
      <option>Other
      </select> 
    </td>
    <td style="border-top:1px solid #99ccff;border-right:1px solid #000000;"><input type="image" src="images/go.gif" onclick="loadPage('profile.html','uppercontent');" border="0"></td>
    <td style="border-top:1px solid #99ccff;border-left:1px solid #99ccff;"><a href="javascript:loadPage('monthly.html','lowercontent')"><img src="images/compare.gif" hspace="3" border="0"></a></td>
    <td width="5" style="border-top:1px solid #99ccff;">&nbsp;</td>
  </tr>
  </table> 
  <!-- end LMP button bar -->
  </td>
</tr>
</form>

</body>
</html>
