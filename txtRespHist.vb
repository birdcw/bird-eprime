Sub txtRespHist(c As Context, theDevice As InputDevice, Optional theStart As RteRunnableObject, _
Optional theEnd As RteRunnableObject, Optional resp As Variant, Optional fileID As Variant)

	' Eprime provides an input device history for storing the history of all 
	' device input regardless of input masks used during the experiment. 
	' This subroutine can be used to selectively copy this history information 
	' to a text file for later analysis. The input values are written to the text file 
	' along with the RTTime (timestamp relative to the start of the experiment).

	' MUST SET A SUFFICIENTLY HIGH theDevice.History.MaxCount AT THE TOP Of THE SessionProc.
	' The default value is 64, meaning that your input device history will start to be 
	' flushed after 64 input events. Maximum History.MaxCount values are 4096 (Eprime 2) 
	' or 1048576 (Eprime 2 Professional)

	' theDevice: Any input device such as Keyboard, SRBox, etc.

	' theStart: Any runnable object such as a stimulus display or procedure. 
	' The text file will only include input device activity that occurred after the 
	' start time of this object.
	
	' theEnd: The other temporal bound. The text file will only include input device activity 
	' that occurred before the finish time of this object.

	' resp: any string of eprime-style input "qwerty{SPACE}1234", etc. Optional - when resp 
	' is omitted, input is written to the text regardless of the value.

	' fileID: this id is used along with the experiments data file name to create a 
	' unique file name "fileID + data file name.txt". In the event that the new file name 
	' is still not available, it will append "Copy - " to prevent overwriting any existing files. 
	' If fileID is omitted, "txtRespHist" is used ion the file name instead.
	
	'Default history max count of 64 indicates potential misuse.
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
	
	Dim beginning As Long, ending As Long
	If IsMissing(theStart) Or theStart Is Nothing Then 
		beginning = 0
	Else
		beginning = theStart.StartTime
	End If
	If IsMissing(theEnd) Or theEnd Is Nothing Or theEnd.FinishTime = 0 Then
		ending = Clock.Read
	Else
		ending = theEnd.FinishTime
	End If
	'Eprime pre-release might cause this script to run before theStart.StartTime or theEnd.FinishTime
	While Clock.Read < beginning Or Clock.Read < ending
	DoEvents
	Wend
		
	Dim theResponseData As ResponseData, theHistory As RteCollection, index As Long
	Set theHistory = theDevice.History.Clone

	For index = 1 To theHistory.Count
	
		Set theResponseData = CResponseData(theHistory(index))
	
		'Writes to the text file all target response times relative to experiment start.
		If theResponseData.RTTime >= beginning And theResponseData.RTTime <= ending _
		And (IsMissing(resp) Or Instr(CStr(resp), theResponseData.Resp) <> 0) Then _
			Print #1, theResponseData.RESP & ebtab & theResponseData.RTTime
				
	Next index

	Close #1

	Set theResponseData = Nothing
	Set theHistory = Nothing

End Sub
