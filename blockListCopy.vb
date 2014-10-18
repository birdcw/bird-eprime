Sub blockListCopy(c As Context, theList As List, attrib As String, _
Optional prevLevels As Variant)

	'This sequence can be used in a random stimulus selection, blocked recognition
	'memory paradigm (study-test-study-test, etc. Where previously studied words are sampled 
	'along with new words during each test phase.). 
	'blockListCopy is called once during each study phase.
	'Other uses might be possible with modification.
	
	'theList: Stimuli sampled during the study phase of the recognition memory task
	'are copied to this "old" list to be sampled for a second time during the recognition
	'test. theList should be the same size as the sample size of each study phase block.
	
	'attrib: This is the name of your stimulus attribute. It must exist both in the 
	'context (c.GetAttrib(attrib)) And in thList.
	
	'prevLevels: This is the number of levels in your top level (block) list that are sampled
	'in between each run of study and test trials. For example, if your block begins with one 
	'level (i.e. and instructions procedure) and then the Study Phase trials, you should use prevLevels = 1.

	If IsMissing(prevLevels) Then prevLevels = 0
	
	Dim theBlockList As List
	Set theBlockList = CList(Rte.GetObject(CStr(c.GetAttrib("Running"))))

	theList.SetAttrib c.GetAttrib(c.GetAttrib("Running") & ".Sample") Mod _
	theBlockList.SizeWithWeight - CInt(prevLevels), attrib, c.GetAttrib(attrib)
	
	Set theBlockList = Nothing

End Sub
