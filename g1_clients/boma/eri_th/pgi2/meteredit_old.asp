<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="setup.css" type="text/css" rel="stylesheet" />
<title>Untitled Document</title>
</head>

<body>
	<table style="border-top: 1px groove #EEEEEE; border-left: 1px groove #EEEEEE; border-right: 1px solid #000000" cellspacing="0" cellpadding="3" border="0" width="400">
  		<tr class="blueBack">
        	<td style="text-align: center"><h1 style="color: #FFFFFF">METER DETAILS - METER ID</h1></td>
        </tr>
        <tr class="greyBack">
        	<td style="border-bottom: 1px solid #000000; height: 20px">&nbsp;</td>
        </tr>
        <tr class="greyBack">
        	<td style="border-bottom: 1px solid #CCCCCC"><table cellspacing="2" cellpadding="2" border="0" align="center">
            	<tr>
                	<td><label>Meter Name:</label></td>
                    <td><%MeterName%></td>
                    <td width="10"></td>
                    <td><label>Serial Number:</label></td>
                    <td><%SerialNumber%></td>
                </tr>
            </table>
        </tr>
        <tr class="greyBack">
        	<td style="border-bottom: 1px solid #000000;"><table style="margin-top: 10px; margin-bottom: 10px" cellspacing="2" cellpadding="2" border="0" align="center">
            		<tr>
                    	<td style="text-align: right"><label>Charge Code:</label></td>
                        <td><%ChargeCode%></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Factor:</label></td>
                        <td><label><%Factor%></label></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Usage.x:</label></td>
                        <td><%UsageX%></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Capacity.x:</label></td>
                        <td><%CapacityX%></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Meter Reference:</label></td>
                        <td><%MeterReference%></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>CT Ratio</label></td>
                        <td><%CTatio%></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Manufacturer:</label></td>
                        <td><%ManuFact%></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Model:</label></td>
                        <td><%Model%></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Voltage:</label></td>
                        <td><%Voltage%></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Location:</label></td>
                        <td><%Location%></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Floor:</label></td>
                        <td><%Floor%></td>
                    </tr>
                    <tr>
                    	<td style="text-align: right"><label>Meter Type:</label></td>
                        <td><%MeterType%></td>
                    </tr>
            </table></td>
        </tr>
    </table>
</body>
</html>