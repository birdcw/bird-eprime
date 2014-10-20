bird-eprime
===========

####Subroutines and Functions for use in Eprime (Psychology Software Tools, Inc)



#####edat/edat_xlsx()

Eprime E-DataAid Excel Export saves files as tab delimited text. This Excel Module will copy that exported text data to xlsx format.

Comparability issues with the file dialogue in Apple OSX.

Supports batching files.

When given E-Merge Data, each session is saved to a separate xlsx file.

The only interaction with this script is a file dialogue at the beginning for selecting which file(s) to process.

#####allFiles(c As Context, theList As List, attrib As String, path As Variant, Optional ext As Variant) As Boolean

This Function Returns True if all files (.wav, .jpg, etc) required by theList are available. Add at the top of your procedure to prevent crashes mid-trial: Debug.Assert allFiles.

theList: This is the list object that contains an attribute (attrib) which points to a file that your experiment will need when sampled.

attrib: this is the name of the attribute that is a full or partial reference to a file that you expect to exist. See path and ext for using partial filenames.

path: Optional. You can specify the path to your attrib files here.

ext: include the file extension (with .) here, if not already provided by attrib.

#####autoSample(theList As List, fileID As String, Optional path As Variant, Optional ext As Variant)

Searches E-recovery text file fileID for all previously sampled items of list object theList. Those items are all re-sampled

Good for preventing items from getting sampled twice during a two session experiment. See my other script "recovery" for a more robust tool that can re-sample and a partially completed experiment for resuming using an incomplete E-recovery file.

theList: Items from this list found in the recovery file are re-sampled.

fileID e-recovery text file used to determine which items were previously sampled.

path: optional path to fileID, if not provided in fileID.

ext: optional extension (with .) for fileID, if not provided in fileID.

#####blockListCopy(c As Context, theList As List, attrib As String, Optional prevLevels As Variant)

This sequence can be used in a random stimulus selection, blocked recognition memory paradigm (study-test-study-test, etc. Where previously studied words are sampled along with new words during each test phase.). blockListCopy is called once during each study phase. Other uses might be possible with modification.

theList: Stimuli sampled during the study phase of the recognition memory task are copied to this "old" list to be sampled for a second time during the recognition test. theList should be the same size as the sample size of each study phase block.

attrib: This is the name of your stimulus attribute. It must exist both in the context (c.GetAttrib(attrib)) And in theList.

prevLevels: This is the number of levels in your top level (block) list that are sampled in between each run of study and test trials. For example, if your block begins with one level (i.e. and instructions procedure) and then the Study Phase trials, you should use prevLevels = 1.

#####checkErun() As Boolean

For EGI Netstation (EEG) Users

This function Returns FALSE if the experiment is configured to run without the EGI custom clock. Accurate EEG timing is not to be expected when this function returns FALSE. For example, you can prevent running session with the incorrect clock setup by adding at the top of your SessionProc: Debug.Assert checkErun

The Eprime Extensions for Netstation package includes an E-Run.ini file. Some users comment out options in this file to switch between the standard clock (for running eprime without netstation) and the custom clock (for accurate timing when paired with netstation).

#####prevSeed(fileID As String, Optional path As Variant, Optional ext As Variant)

This is an alternative to using a startup parameter to use RandomSeed for running random selection experiments at an identical state as a previous session.

Similar to my other two functions "autoSample" and "recovery" except the experiment is run from the top here.

fileID: e-recovery text file for acquiring a previously used RandomSeed.

path: optional path to fileID, if not provided in fileID.

ext: optional extension (with .) for fileID, if not provided in fileID.

#####recovery(fileID As String, Optional path As Variant, Optional ext As Variant)

Re-creates the conditions of an experiment using any incomplete E-Recovery data file. The experiment will seamlessly return to the original break point as if nothing happened.

IMPORTANT! This script uses the Tag property with every list object in your experiment. The script requires you to set the Tag for every list object to the total # of samples See "cycles x samples/cycle in the list's summary window. Cycles x samples/cycle is needed, but it is not available at runtime. Users of this script need to store that number in the Tag property. If PST ever adds support for the equivalent of Get cycles x samples/cycle, I will update this script and remove the tag requirement.

IMPORTANT EXCEPTION: If your experiment uses a Counterbalance list, set its Tag to 2, not 1.

CAUTION! If you need to make any changes to any list objects at runtime, such as SetWeight, Set TerminateCondition, Set ResetCondition, or Reset, you must do so BEFORE calling this script. This script will not be compatible with any experiments that require you to make changes after it has been called.

CAUTION! This script resumes the experiment inside the last running procedure. i.e. if your trial procedure requires displayed items that are written to the display in a prior procedure (instruction, key reminders written in an instruction procedure that are not cleared during the trial procedure, etc), those displays might be bypassed when resuming the experiment inside the last running procedure.

fileID: e-recovery text file used to determine which items were previously sampled.

path: optional path to fileID, if not provided in fileID.

ext: optional extension (with .) for fileID, if not provided in fileID.

#####superReset(theList As List)

Before Resetting theList, this sequence sets ResetCondition to all samples and TerminateCondition to 1 cycle.

Often ".Reset" alone will lead to unexpected ResetCondition and/or TerminateCondition behavior.

#####tagitNS(c As Context, tags() As String, theStim As RteRunnableObject, Optional tagSpace As Variant)

For EGI Netstation and Brain Vision Analyzer (EEG) Users If you record EEG using Netstation and Eprime, the EENS TRSP information is not readable when your EEG recordings are exported to some Non-Netstation analysis software such as Brain Vision Analyzer.

All 4 character length RteRunnableObject tags will be readable in Brain Vision Analyzer when sent using NetStation_SendTrialEvent, provided that the tags do not overlap temporally.

This sequence sends an array of 4 character tags to Netstation (NetStation_SendTrialEvent) and spaces them a milliseconds apart from each other.

tags(): All strings in this array will be sent as events to Netstation. For each item ,any characters more than 4 will be ignored.

theStim: This is the trial event that will determine approximately where the new tags will be placed (see tagSpace below).

tagSpace: This is the number in milliseconds that determines the spacing interval for each new tag. The default value is 50 ms. For example, the first item in tags() will be placed 50 ms after theStim, the second is placed 100 ms after theStim, etc. 50 ms is a compatible distance for Brain Vision Analyzer users.

#####txtRespHist(c As Context, theDevice As InputDevice, theStart As RteRunnableObject, _ theEnd As RteRunnableObject, Optional resp As Variant, Optional fileID As Variant)

Eprime provides an input device history for storing the history of all device input regardless of input masks used during the experiment. This subroutine can be used to selectively copy this history information to a text file for later analysis. The input values are written to the text file along with the RTTime (timestamp relative to the start of the experiment).

MUST SET A SUFFICIENTLY HIGH theDevice.History.MaxCount AT THE TOP Of THE SessionProc. The default value is 64, meaning that your input device history will start to be flushed after 64 input events. Maximum History.MaxCount values are 4096 (Eprime 2) or 1048576 (Eprime 2 Professional)

theDevice: Any input device such as Keyboard, SRBox, etc.

theStart: Any runnable object such as a stimulus display or procedure. The text file will only include input device activity that occurred after the start time of this object. To make theStart optional (no time constraint), set it to your top level procedure object (i.e. Session Proc)
	
theEnd: The other temporal bound. The text file will only include input device activity 
that occurred before the finish time of this object. to make theEnd optional (no time constraint),
set it to your top level procedure object (i.e. Session Proc)

resp: any string of eprime-style input "qwerty{SPACE}1234", etc. Optional - when resp is omitted, input is written to the text regardless of the value.

fileID: this id is used along with the experiments data file name to create a unique file name "fileID + data file name.txt". In the event that the new file name is still not available, it will append "Copy - " to prevent overwriting any existing files. If fileID is omitted, "txtRespHist" is used ion the file name instead.
