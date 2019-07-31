/**********************************
	JavaScript 64Bit RSA Class

	Author: benwhite@columbus.rr.com

	Description:
		- This is a 64bit RSA class that allows you to encrypt and decrypt strings
		- While 64 bit is default, you can lower the encryption to 24 bit
		if you like. this will of course be easier to crack, but the encrypted data
		string will be shorter. (for low bandwidth applications?)

	Usage:
		var RSA = new clsRSA;
		RSA.Bytes = 8; //64bit (optional, 8 is default)
		RSA.KeyGen();
		alert(RSA.Dec(RSA.Enc("Works!!!")));

		or

		var RSA = new clsRSA;
		RSA.N = 76513061;
		RSA.E = 19000901;
		RSA.D = 17148101;
		alert(RSA.Dec(RSA.Enc("Works!!!")));

	Resources:
		William Gerard Griffiths (http://www.webdreams.org.uk)
			VB RSA module
		Lewis Edward Moten III (http://www.lewismoten.com)
			ASP RSA include
		Nick Radford
			hex -> dec routine
			dec -> hex routine

**********************************/


function clsRSA() {

/**********************************
	Private Functions
**********************************/

	function Mult(x, pg, n) {
		var y = 1;
		x = Math.floor(x);
	    while (pg > 0) {
	        while ((pg / 2) == Math.floor(pg / 2)) {
	            x = nMod((x * x), n);
	            pg = pg / 2;
	        }
	        y = nMod((x * y), n);
	        pg--;
	    }
	    return y;
	}

	function nMod(x, y) {
		if (y==0) {return 0}
	  	return x - (Math.floor(x / y) * y);
	}

	function GCD(nPHI) {
	    //generates a random number relatively prime to PHI
	    var x = 0;
	    var y = 0;
	    var N_UP = Math.pow(10,size)-1; //set upper limit of random number For E (8 = 99999999)
	    var N_LW = Math.pow(10,size-1); //set lower limit of random number For E (8 = 10000000)
	    var nE = Math.floor((N_UP - N_LW + 1) * Math.random() + N_LW)
		while (true) {
		    x = nPHI % nE
		    y = x % nE
		    if (y != 0 && IsPrime(nE)) {
		        return nE;
		    } else {
		        nE = nE + 1
		    }
		}
	}

	function Euler(E3, PHI3) {
	    //genetates D from (E and PHI) using the Euler algorithm
	    var u1=1, u2=0, u3=PHI3;
	    var v1=0, v2=1, v3=E3;
	    while (v3 != 0) {
	        q = Math.floor(u3 / v3);
	        t1 = u1 - q * v1; t2 = u2 - q * v2; t3 = u3 - q * v3;
	        u1 = v1; u2 = v2; u3 = v3;
	        v1 = t1; v2 = t2; v3 = t3;
	        z = 1;
	    }
	    if (u2 < 0) {
	        return u2 + PHI3;
	    } else {
	        return u2;
	    }
	}

	function IsPrime(pLngNumber) {
		var lLngSquare = 0;
		var lLngIndex = 0;
	    if (pLngNumber < 2) {return false}
	    if (pLngNumber % 2 == 0) {return false}
	    lLngSquare = Math.sqrt(pLngNumber)
	    for (lLngIndex=3;lLngIndex<=lLngSquare;lLngIndex+=2){
	        if (pLngNumber % lLngIndex == 0) {return false}
	    }
		return true;
	}

	function H2D(HexVal) {
		return parseInt(HexVal.toUpperCase(),16);
	}

	function D2H(DecVal) {
    	var HexChars = '0123456789ABCDEF';
    	var HexStr = ''
    	while (DecVal>0){
    		var HexStr = HexChars.charAt( DecVal%16 ) + HexStr;
    		var DecVal = Math.floor(DecVal/16);
    	}
    	return HexStr
    }

    function SetSize(intSize) {
    	if (intSize < 3 || intSize > 8) {
    		size = 8;
    	} else {
    		size = intSize;
    	}
    }

/**********************************
	Public Functions
**********************************/

	function Enc(strText) {
		var strEnc = '';
		var strHEX = '';
		SetSize(this.Bytes);
		if (strText == '' || this.E==0 || this.N==0) {return ''}
	    for (var i=0;i<strText.length;i++) {
		    strHEX = D2H(Mult(strText.charCodeAt(i), this.E, this.N));
	        while (strHEX.length < size) {
	            strHEX = "0" + strHEX;
	        }
		    strEnc += strHEX;
		}
		return strEnc;
	}

	function Dec(strEnc) {
		var strDec = '';
		SetSize(this.Bytes);
		if (strEnc=='' || this.D==0 || this.N==0) {return ''}
		if (strEnc.length % size != 0) {return ''}
		for (z=0;z<strEnc.length;z+=size) {
		    tok = H2D(strEnc.slice(z, z+size));
		    strDec += String.fromCharCode(Mult(tok, this.D, this.N))
		}
		return strDec;
	}

	function KeyGen() {
	    var p = 0, q = 0, E = 0, D = 0, N = 0, PHI = 0;
	    SetSize(this.Bytes);
	    var PQ_UP = Math.ceil(Math.pow(10,size/2))-1; //set upper limit of random number (8 = 99999999)
	    var PQ_LW = Math.floor(Math.pow(10,(size-1)/2))+1; //set upper limit of random number (8 = 3163)
	    var KEY_LW = Math.pow(10,size-1); //set minimum value (8 = 10000000)
	    while (D <= KEY_LW || D==E) { //makes sure keys meet minimum bitlength
	        while (!IsPrime(p))
	        	{p = Math.floor((PQ_UP - PQ_LW + 1) * Math.random() + PQ_LW)}
	        while (!IsPrime(q) || p==q)
	        	{q = Math.floor((PQ_UP - PQ_LW + 1) * Math.random() + PQ_LW)}
	        PHI = (p - 1) * (q - 1);
	        E = GCD(PHI);
	        D = Euler(E, PHI);
	        N = p * q;
	    }
        this.N = N;
        this.E = E;
        this.D = D;
	}

/**********************************
	Initialize Class Variables
**********************************/

	//Internal Vars
	var size = 0;

	//External Vars
	this.E = 0;
	this.D = 0;
	this.N = 0;
	this.Bytes = 8;

	//External Functions
	this.Enc = Enc;
	this.Dec = Dec;
	this.KeyGen = KeyGen;
}