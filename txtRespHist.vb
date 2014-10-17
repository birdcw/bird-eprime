Sub txtRespHist(c As Context, theDevice As InputDevice, theStart As RteRunnableObject, _
theEnd As RteRunnableObject, Optional resp As Variant, Optional fileID As Variant)

	'Writes to a text file all target response times relative to experiment start.
	
	'In detail: writes to a text file, fileID, the "onset time" (RTTime) 
	'of each "theDevice" input defined in resp (a string of target response values) between 
	'the end of runnable object (procedure, input object, etc) "theStart" and 
	'the begining Of runnable object "theEnd". Disable theStart and/or theEnd by setting them to SessionProc.
	'All responses are target responses if resp is not provided.
	
	'MUST SET A SUFFICIENTLY HIGH theDevice.History.MaxCount AT THE TOP Of THE SessionProc.
	'Maximum History.MaxCount values are 4096 (Eprime 2) or 1048576 (Eprime 2 Professional)
	
	'Default history max count of 64 indicates misuse.
	If theDevice.History.MaxCount = 64 Then _
		Debug.Print "CAUTION: You are using the default " & theDevice.Name & ".History.MaxCount = 64\n" &_
		"This might not be a sufficiently high MaxCount For txtRespHist.\n" &_
		"Maximum Values: " & theDevice.Name & ".History.MaxCount = 4096 (standard) or 1048576 (professional)"
		
	'Default value
	If IsMissing(fileID) Then fileID = "txtRespHist"
	
	'Avoids file name collisions	
	While FileExists(CStr(fileID) & "-" & c.DataFile.Filename)
		fileID = "Copy - " & fileID
	Wend
	'Writes to this file
	Open CStr(fileID) & "-" & c.DataFile.Filename  For Output As #1
	
	'file headers
	Print #1, theDevice.Name & ebtab & "RTTime"
	
	'Eprime pre-release might cause this script to run before the "beginning" or "ending" time.
	While Clock.Read < theStart.StartTime Or Clock.Read < theEnd.FinishTime
		DoEvents
	Wend
	
	'If the script is called before theEnd is run, it's FinishTime = 0
	Dim ending As Long
	ending = theEnd.FinishTime
	If ending = 0 Then ending = Clock.Read
		
	Dim theResponseData As ResponseData, theHistory As RteCollection, index As Long
	Set theHistory = theDevice.History.Clone

	For index = 1 To theHistory.Count
	
		Set theResponseData = CResponseData(theHistory(index))
	
		'Writes to the text file all target response times relative to experiment start.
		If theResponseData.RTTime >= theStart.StartTime And theResponseData.RTTime <= ending _
		And (IsMissing(resp) Or Instr(CStr(resp), theResponseData.Resp)) Then _
			Print #1, theResponseData.RESP & ebtab & theResponseData.RTTime
				
	Next index

	Close #1

	Set theResponseData = Nothing
	Set theHistory = Nothing

End Sub
