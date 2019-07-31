<%
' Compiled by Lewis Edward Moten III
' lewis@moten.com
' http://www.lewismoten.com
' Wednesday, May 09, 2001 05:42 PM GMT +5

' RSA Encryption Class
'
' .KeyEnc
'		Key for others to encrypt data with.
'
' .KeyDec
'		Your personal private key.  Keep this hidden.
'
' .KeyMod
'		Used with both public and private keys when encrypting and decrypting data.
'
' .KeyGen
'		Used to generate both public and private keys for encrypting and decrypting data.
'
' .Encode(pStrMessage)
'		Encrypts message and returns in numeric format
'
' .Decode(pStrMessage)
'		Decrypts message and returns a string
'
Class clsRSA

	Public KeyEnc
	Public KeyDec
	Public KeyMod

	Private Function Mult(ByVal x, ByVal pg, ByVal m)
		dim y
		y = 1
	    Do While pg > 0
	        Do While (pg / 2) = Int((pg / 2))
	            x = nMod((x * x), m)
	            pg = pg / 2
	        Loop
	        y = nMod((x * y), m)
	        pg = pg - 1
	    Loop
	    Mult = y
	End Function


	Private Function nMod(x, y)
		nMod = 0
		if y = 0 then Exit Function
		nMod = x - (Int(x / y) * y)
	End Function


	Public Function Encode(ByVal tIp)
		Dim encSt, z
		Dim strMult
		    If tIp = "" Then Exit Function
		    For z = 1 To Len(tIp)
		        encSt = encSt & NumberToHex(Mult(CLng(Asc(Mid(tIp, z, 1))), KeyEnc, KeyMod),8)
		    Next
		Encode = encSt
	End Function


	Public Function Decode(ByVal tIp)
		Dim decSt, z
		if Len(tIp) Mod 8 <> 0 then Exit Function
		For z = 1 To Len(tIp) Step 8
		    decSt = decSt + Chr(Mult(HexToNumber(Mid(tIp, z, 8)), KeyDec, KeyMod))
		Next
		Decode = decSt
	End Function

	Public Sub KeyGen()
	    'Generates the keys for E, D and N
	    Dim E, D, N
	    Const PQ_UP = 9999 'set upper limit of random number
	    Const PQ_LW = 3170 'set lower limit of random number
	    Const KEY_LOWER_LIMIT  = 10000000 'set For 64bit minimum
	    p = 0: q = 0
	    Randomize
	    Do Until D > KEY_LOWER_LIMIT 'makes sure keys are 64bit minimum
	        Do Until IsPrime(p) And IsPrime(q) ' make sure q and q are primes
	            p = clng((PQ_UP - PQ_LW + 1) * Rnd + PQ_LW)
	            q = clng((PQ_UP - PQ_LW + 1) * Rnd + PQ_LW)
	        Loop
	        N = clng(p * q)
	        PHI = (p - 1) * (q - 1)
	        E = clng(GCD(PHI))
	        D = clng(Euler(E, PHI))
	    Loop
	    KeyEnc = E
	    KeyDec = D
	    KeyMod = N
	End Sub

	Private Function Euler(E3, PHI3)
	    'genetates D from (E and PHI) using the Euler algorithm
	    On Error Resume Next
	    Dim u1, u2, u3, v1, v2, v3, q
	    Dim t1, t2, t3, z, vv, inverse
	    u1 = 1
	    u2 = 0
	    u3 = PHI3
	    v1 = 0
	    v2 = 1
	    v3 = E3
	    Do Until (v3 = 0)
	        q = Int(u3 / v3)
	        t1 = u1 - q * v1: t2 = u2 - q * v2: t3 = u3 - q * v3
	        u1 = v1: u2 = v2: u3 = v3
	        v1 = t1: v2 = t2: v3 = t3
	        z = 1
	    Loop
	    If (u2 < 0) Then
	        inverse = u2 + PHI3
	    Else
	        inverse = u2
	    End If
	    Euler = inverse
	End Function

	Private Function GCD(nPHI)
	    On Error Resume Next
	    Dim nE, y
	    Const N_UP = 99999999 'set upper limit of random number For E
	    Const N_LW = 10000000 'set lower limit of random number For E
	    Randomize
	    nE = Int((N_UP - N_LW + 1) * Rnd + N_LW)
		Do
		    x = nPHI Mod nE
		    y = x Mod nE
		    If y <> 0 And IsPrime(nE) Then
		        GCD = nE
		        Exit Function
		    Else
		        nE = nE + 1
		    End If
		Loop
	End Function

	Private Function IsPrime(lngNumber)
	    On Error Resume Next
	    Dim lngCount
	    Dim lngSqr
	    Dim x
	    lngSqr = Int(Sqr(lngNumber)) ' Get the int square root
	    If lngNumber < 2 Then
	        IsPrime = False
	        Exit Function
	    End If
	    lngCount = 2
	    IsPrime = True
	    If lngNumber Mod lngCount = 0 Then
	        IsPrime = False
	        Exit Function
	    End If
	    lngCount = 3
	    For x = lngCount To lngSqr Step 2
	        If lngNumber Mod x = 0 Then
	            IsPrime = False
	            Exit Function
	        End If
	    Next
	End Function

	Private Function NumberToHex(ByRef pLngNumber, ByRef pLngLength)
		NumberToHex = Right(String(pLngLength, "0") & Hex(pLngNumber), pLngLength)
	End Function


	Private Function HexToNumber(ByRef pStrHex)
		HexToNumber = CLng("&h" & pStrHex)
	End Function

End Class
%>