bird-eprime
===========

Subroutines and Functions for use in Eprime (Psychology Software Tools, Inc)

1. txtRespHist(c As Context, theDevice As InputDevice, theStart As RteRunnableObject, _
theEnd As RteRunnableObject, Optional resp As Variant, Optional fileID As Variant)

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
	' start time of this object. Set theStart to your top-level procedure (i.e. SessionProc) 
	' if you want the text file to include input from the very beginning of your experiment.

	' theEnd: The other temporal bound. The text file will only include input device activity 
	' that occurred before the finish time of this object. Set theEnd to your top-level 
	' procedure (i.e. SessionProc) if you want the text file to include input all the way up 
	' until the moment this subroutine is called.

	' resp: any string of eprime-style input "qwerty{SPACE}1234", etc. Optional - when resp 
	' is omitted, input is written to the text regardless of the value.

	' fileID: this id is used along with the experiments data file name to create a 
	' unique file name "fileID + data file name.txt". In the event that the new file name 
	' is still not available, it will append "Copy - " to prevent overwriting any existing files. 
	' If fileID is omitted, "txtRespHist" is used ion the file name instead.





