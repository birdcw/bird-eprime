Sub prevSeed(fileID As String, Optional path As Variant, Optional ext As Variant)

	'This is an alternative to using a startup parameter to use RandomSeed for
	'running random selection experiments at an identical state as a previous session.
	
	'Similar to my other two functions "autoSample" and "recovery" except the experiment
	'is run from the top here.
	
	'fileID: e-recovery text file for acquiring a previously used RandomSeed.
	
	'path: optional path to fileID, if not provided in fileID.
	
	'ext: optional extension (with .) for fileID, if not provided in fileID.
	
	If IsMissing(path) Then path = ebNullString
	If IsMissing(ext) Then ext = ebNullString
	
	Dim lis As List, itm As Integer, lin As String
	
	Open CStr(path) & fileID & CStr(ext) For Input Access Read As #1
	While InStr(lin, "RandomSeed:") = 0
		Line Input #1, lin
	Wend
	Close #1
	
	PRNG.SetSeed CLng(Mid(lin, InStr(lin, ":") + 2))
	
	For itm = 1 To Rte.GetObjectCount
			
		Set lis = CList(Rte.GetObject(itm))
		
		If lis Is Not Nothing Then
			
			If lis.Order.TypeName = "RandomOrder" Then 
				Set lis.Order = New RandomOrder
			ElseIf lis.Order.TypeName = "RandomReplaceOrder" Then 
				Set lis.Order = New RandomReplaceOrder
			End If
			
		End If
		
	Next itm
	
	Set lis = Nothing
	
End Sub
