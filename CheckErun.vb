Function checkErun() As Boolean

	'For EGI Netstation (EEG) Users
	
	'This function Returns FALSE if the experiment is configured to run 
	'without the EGI custom clock. Accurate EEG timing is not to be expected 
	'when this function returns FALSE. For example, you can prevent running a 
	'session with the incorrect clock setup by adding at the top of your SessionProc:
	'Debug.Assert checkErun
	
	'The Eprime Extensions for Netstation package includes an E-Run.ini file.
	'Some users comment out options in this file to switch between the
	'standard clock (for running eprime without netstation) and the custom
	'clock (for accurate timing when paired with netstation).

	If Not FileExists(Basic.HomeDir$ & "\\E-Run.ini") Then 
		
		checkErun = False
		
		Debug.Print "\n\nWARNING - File Does Not Exist:\n\n" &_
		Basic.HomeDir$ & "\\E-Run.ini\n\n"
		
		Exit Function
	
	End If
	
	
	Dim corOptions As String, corClock As String, corClockSNTP As String, _
	corIndex As String
	
	'Correct E-Run.ini
	corOptions = "[Options]"
	corClock = "CustomClock=""EgiClockExtension.ebn"""
	corClockSNTP = "CustomClock=""SNTPClockExtension.ebn"""
	corIndex = "CustomClockIndex=0"
	
	Dim lnOptions As String, lnClock As String, lnIndex As String
	
	'E-Run.ini file available
	Open Basic.HomeDir$ & "\\E-Run.ini" For Input Access Read As #1
	Input #1, lnOptions, lnClock, lnIndex
	Close #1

	Dim allCor As Boolean
	
	'is the available file all correct
	allCor = ( _
	lnOptions = corOptions And _
	(lnClock = corClock Or lnClock = corClockSNTP) And _
	lnIndex = corIndex)
	
	If allCor = False Then Debug.Print "\n\nWARNING: " & Basic.HomeDir$ &_
	"\\E-Run.ini Is not configured for correct use with Netstation.\n\n" &_
	"If you are using SNTP, the correct file contents are:\n" &_
	corOptions & "\n" & corClockSNTP & "\n" & corIndex &_
	"\n\nIf you are a single clock timing box, the correct file contents are:\n" &_
	corOptions & "\n" & corClock & "\n" & corIndex & "\n\n"
	
	checkErun = allCor

End Function
