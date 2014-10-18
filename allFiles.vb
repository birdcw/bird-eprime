Function allFiles(c As Context, theList As List, attrib As String, _
path As Variant, Optional ext As Variant) As Boolean

	'This Function Returns True if all files (.wav, .jpg, etc) required
	'by theList are available. Add at the top of your procedure to prevent 
	'crashes mid-trial: Debug.Assert allFiles 
	
	'theList: This is the list object that contains an attribute (attrib) which 
	'points to a file that your experiment will need when sampled.
	
	'attrib: this is the name of the attribute that is a full or partial 
	'reference to a file that you expect to exist.
	'See path and ext for using partial filenames.
	
	'path: Optional. You can specify the path to your attrib files here.
	
	'ext: include the file extension (with .) here, if not already provided by attrib.

	'path and ext are optional for when attrib is a complete filename.
	If IsMissing(path) Then path = ebnullstring
	If IsMissing(ext) Then ext = ebnullstring
	
	Dim index As Integer, files As String, nextFile As String
	files = ebnullstring
	
	'Checks theList for any files (path + attrib + ext) that are missing.
	For index = 1 To theList.Size
	
		nextFile = CStr(path) & theList.GetAttrib(index, attrib) & CStr(ext)
		
	 	If Not FileExists(nextFile) Then _
		files = files & index & ebtab & nextFile & "\n"
		
	Next index
	
	If files = ebnullstring Then
	
		allFiles = True
		
	Else
	
		allFiles = False
		
		Dim fileID As String
		fileID = "missing-" & theList.Name & "-" & c.DataFile.Filename
		
		While FileExists(fileID)
			fileID = "Copy - " & fileID
		Wend
		
		'Writes a missing file report to this file
		Open fileID For Output As #1
		Print #1, "WARNING: The following files were not available:\n\n"
		Print #1, theList.Name & ebtab & attrib
		Print #1, files
		close #1
		
		'Estudio Debug Tab
		Debug.Print "\n\nWARNING: The following files are not available:\n\n"
		Debug.Print theList.Name & ebtab & attrib
		Debug.Print files
		
	End If

End Function
