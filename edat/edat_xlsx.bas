Attribute VB_Name = "edat_xlsx"
Sub edat_xlsx()

	'Comparability issues with the file dialogue in Apple OSX.

	'Eprime E-DataAid Excel Export saves files as tab delimited text.

	'This Excel Module will copy that exported text data to xlsx format.

	'Supports batching files.

	'When given E-Merge Data, each session is saved to a separate xlsx file.

	'The only interaction with this script is a file dialogue at the beginning
	'for selecting which file(s) to process.
			
	Dim path As String, file As Variant, files() As Variant, folder As String
	Dim NewWorkbook As Workbook, bookName As String
	Dim header As Range, top As Range, bottom As Range

	On Err GoTo ErrorMessage

	files = Application.GetOpenFilename(FileFilter:="Text Files (*.txt), *.txt", _
	Title:="Select File(s) Exported From E-DataAid (txt)", _
	MultiSelect:=True)

	'macro runs faster with this
	Application.ScreenUpdating = False
	Application.DisplayAlerts = False

	'Changes working directory to the file's location
	If InStr(Application.OperatingSystem, "Windows") Then
		folder = "\"
	Else
		folder = ":"
	End If

	path = files(LBound(files))

	path = Left(path, _
			InStrRev(path, folder))
		  
	ChDir path

	For Each file In files

		fileName = Right(file, Len( _
			file) - InStrRev(file, folder))

		'opens the file
		Workbooks.OpenText fileName:=(fileName), Origin:= _
			xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
			xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
			Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1)
				
		Workbooks(fileName).Activate
		
		'The header row of edat "variable" names. SessionTime is used to identify separate data sets.
		Set header = ActiveWorkbook.ActiveSheet.Range("A1", Cells(Rows.Count, Columns.Count)) _
		.Find("SessionTime", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
		
		Set top = header.Offset(1, 0)
		Set header = header.EntireRow
	   
	   'rows are removed after copied over to a separate excel sheet.
	   'We're finished with each text file once the row below the header row is blank.
		Do While top.Value <> vbNullString
	   
			'End of the data file
			Set bottom = top.End(xlDown)
	   
			'Each data set is identified by a unique start time, logged as "SessionTime"
			'We set top-to-bottom to cover one data set at a time.
			Do While top.Value <> bottom.Value
				Set bottom = bottom.Offset(-1, 0)
			Loop
			
			'Each data set is copied to its own excel file. experiment-subject-session.xls
			bookName = CStr(header.Find("ExperimentName", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Offset(1, 0).Value) & "-" & _
			CInt(header.Find("Subject", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Offset(1, 0).Value) & "-" & _
			CInt(header.Find("Session", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Offset(1, 0).Value)
			
			Set NewBook = Workbooks.Add
			
			With NewBook
				.Title = bookName
				.SaveAs fileName:=bookName & ".xlsx"
			End With
			
			Workbooks(fileName).Activate
			
			'Data is copied from the text file to excel file
			ActiveWorkbook.ActiveSheet.Range("A1", top.Offset(-1, 0).End(xlToRight).Offset(bottom.Row - 2)).Copy _
			Destination:=NewBook.Sheets("Sheet1").Range("A1")
			
			NewBook.Close SaveChanges:=True
			
			Workbooks(fileName).Activate
			
			'Temporarily deleted rows to keep track of our progress
			Range(top, bottom).EntireRow.Delete
			
			Set top = header.Find("SessionTime", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True).Offset(1, 0)
			
		Loop
		
		'Important to not save, because all we did was remove everything in this file
		'to keep track as it is transferred to the new file.
		Workbooks(fileName).Close SaveChanges:=False
		
	Next file
	   
	Application.ScreenUpdating = True
	Application.DisplayAlerts = True

	Beep
	MsgBox "Macro completed successfully."

	Exit Sub

	ErrorMessage:

	Application.ScreenUpdating = True
	Application.DisplayAlerts = True

	Beep
	MsgBox "Macro failed while processing " & fileName
        
End Sub

