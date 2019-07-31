<script>
function popUp()
{
  this.setTimeout("transferURL()",5000);
 // parent.document.location = "meterinfo.asp";
}

function transferURL()
{
	parent.document.location="meterinfo.asp";
}

</script>

<link rel="Stylesheet" href="styles.css" type="text/css">
<body bgcolor="#FFFFFF" onload="popUp();">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="100%" align="center">
  <tr>
    <td>
      <div name="output" align="center">
	  The Load has been disconnected <br />
	  Transfering back to the meter page in 5 seconds
	  </div>
    </td>
  </tr>
</table>
</body>