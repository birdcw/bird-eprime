Sub autoSample(theList As List, fileID As String, _
Optional path As Variant, Optional ext As Variant)

	'Searches E-recovery text file fileID for all previously sampled items
	'of list object theList. Those items are all re-sampled
	
	'Good for preventing items from getting sampled twice during a two session experiment.
	'See my other script "recovery" for a more robust tool that can re-sample and
	'a partially completed experiment for resuming using an incomplete E-recovery file.
	
	'theList: Items from this list found in the recovery file are re-sampled.
	
	'fileID e-recovery text file used to determine which items were previously sampled.
	
	'path: optional path to fileID, if not provided in fileID.
	
	'ext: optional extension (with .) for fileID, if not provided in fileID.
	
	'path and ext are optional for when fileID is a complete filename.
	If IsMissing(path) Then path = ebnullstring
	If IsMissing(ext) Then ext = ebnullstring
	
	
	
	Open CStr(path) & fileID & CStr(ext) For Input Access Read As #1
	Dim lin As String, attrib As Variant, prevSeed As Long
	
	prevSeed = PRNG.GetSeed
			
	While Not EOF(1)
			
		Line Input #1, lin
		
		If InStr(lin, theList.Name & ":") <> 0 Then
			attrib = theList.GetNextAttrib(theList.Attribs.Item(1).Name)	
		ElseIf Instr(lin, "RandomSeed:") <> 0 And _
		CStr(Mid(lin, InStr(lin, ":") + 2)) <> CStr(PRNG.GetSeed) Then
			PRNG.SetSeed CLng(Mid(lin, InStr(lin, ":") + 2))
			If theList.Order.TypeName = "RandomOrder" Then 
				Set theList.Order = New RandomOrder
			ElseIf theList.Order.TypeName = "RandomReplaceOrder" Then 
				Set theList.Order = New RandomReplaceOrder
			End If
		End If
				
	Wend
	
	Close #1
	
	PRNG.SetSeed prevSeed
	
End Sub
