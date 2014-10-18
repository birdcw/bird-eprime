Sub superReset(theList As List)

	'Before Resetting theList, this sequence sets
	'ResetCondition to all samples and TerminateCondition to 1 cycle.
	
	'Often ".Reset" alone will lead to unexpected ResetCondition and/or
	'TerminateCondition behavior.
	
	Set theList.ResetCondition = Samples(theList.SizeWithWeight)
	Set theList.TerminateCondition = Cycles(1)
	theList.Reset
	
End Sub
