<%

'Return a javascript frendly string
Private Function FixJS(strInput)
	FixJS = replace(replace(replace(replace(strInput,"\","\\"),"'","\'"),chr(10),"<br>"),chr(13),"")
End Function

'Return the path portion of the file path
Private Function IIF(Exp,TrueState,FalseState)
	if Exp then
		IIF = TrueState
	Else
		IIF = FalseState
	End If
End Function

%>