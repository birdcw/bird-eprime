Sub recovery(fileID As String, Optional path As Variant, Optional ext As Variant)

	' Re-creates the conditions of an experiment using any incomplete E-Recovery data file.
	' The experiment will seamlessly return to the original break point as if nothing happened.

	' IMPORTANT! This script uses the Tag property with every list object in your experiment.
	' The script requires you to set the Tag for every list object to the total # of samples
	' See "cycles x samples/cycle in the list's summary window. Cycles x samples/cycle is 
	' is needed, but it is not available at runtime. Users of this script need to store that 
	' number in the Tag property. If PST ever adds support for the equivalent of
	' Get cycles x samples/cycle, I will update this script and remove the tag requirement.
	
	' CAUTION! If you need to make any changes to any list objects at runtime, such as SetWeight,
	' Set TerminateCondition, Set ResetCondition, or Reset, you must do so BEFORE calling this script. 
	' This script will not be compatible with any experiments that require you to make changes after 
	' it has been called.
	
	' CAUTION! This script resumes the experiment inside the last running procedure.
	' i.e. if your trial procedure requires displayed items that are written to the display in
	' a prior procedure (instruction, key reminders written in an instruction procedure 
	' that are not cleared during the trial procedure, etc), those displays might be bypassed 
	' when resuming the experiment inside the last running procedure.
	
	'fileID: e-recovery text file used to determine which items were previously sampled.
	
	'path: optional path to fileID, if not provided in fileID.
	
	'ext: optional extension (with .) for fileID, if not provided in fileID.

	'path and ext are optional for when attrib is a complete filename.
	If IsMissing(path) Then path = ebNullString
	If IsMissing(ext) Then ext = ebNullString
	
	Dim lis As List, itm As Integer, lin As String, attrib As Variant
	
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

			Open CStr(path) & fileID & CStr(ext) For Input Access Read As #1
		
			While Not EOF(1)
			
				Line Input #1, lin
			
				If InStr(lin, lis.Name & ":") <> 0 Then
					attrib = lis.GetNextAttrib(lis.Attribs.Item(1).Name)
					lis.Tag = lis.Tag - 1
				End If
					
			Wend
		
			Close #1
			
			Set lis.TerminateCondition = Samples(lis.Tag)
			
		End If
	
	Next itm

	Set lis = Nothing
	
	For itm = 10 To 1 Step -1
		Display.Canvas.Text Display.XRes/2 -50, Display.YRes/2, "Restarting in " & Format(CInt(itm),"00")
		Sleep(1000)
	Next itm
	
End Sub
