Sub tagitNS(c As Context, tags() As String, theStim As RteRunnableObject, _
Optional tagSpace As Variant)

	'For EGI Netstation and Brain Vision Analyzer (EEG) Users
	'If you record EEG using Netstation and Eprime, the EENS TRSP
	'information is not readable when your EEG recordings are exported
	'to some Non-Netstation analysis software such as Brain Vision Analyzer.
	
	'All 4 character length RteRunnableObject tags will be readable in Brain
	'Vision Analyzer when sent using NetStation_SendTrialEvent, provided that 
	'the tags do not overlap temporally.
	
	'This sequence sends an array of 4 character tags to Netstation 
	'(NetStation_SendTrialEvent) and spaces them a milliseconds apart from each other.
	
	'tags(): All strings in this array will be sent as events to Netstation.
	'For each item ,any characters more than 4 will be ignored.
	
	'theStim: This is the trial event that will determine approximately where the 
	'new tags will be placed (see tagSpace below).
	
	'tagSpace: This is the number in milliseconds that determines the spacing interval
	'for each new tag. The default value is 50 ms. For example, the first item in tags()
	'will be placed 50 ms after theStim, the second is placed 100 ms after theStim, etc.
	'50 ms is a compatible distance for Brain Vision Analyzer users.

	If IsMissing(tagSpace) Then tagSpace = 50
	
	Dim tag As Variant, prevTag As String
	prevTag = theStim.tag
	
	For Each tag In tags
	
		theStim.tag = Left(tag, 4)
		theStim.OnsetTime = theStim.OnsetTime + CInt(tagSpace)
		NetStation_SendTrialEvent c, theStim
		
	Next tag
	
	theStim.OnsetTime = theStim.OnsetTime - _
	(UBound(tags) - LBound(tags) + 1) * CInt(tagSpace)
	theStim.tag = prevTag
	
End Sub
