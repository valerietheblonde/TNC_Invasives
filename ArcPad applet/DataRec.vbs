''''This sub is devoted to getting the data recorder's name so that they don't have to enter it each time
'the following sub, LoadDataRecForm, sets the variable GetOutofHere so that the form can exit peacefully
'from multiple places in the sign-in process.

Sub LoadDataRecForm

Dim strArcPadAPX                  '
Dim objFSO                        '
Dim pXMLtxt                       '
Dim strContents                   '
Dim versiontext                   '
Dim strFirstLine                  '
Dim strNewContents                '
Dim pXML                          '
Dim blnExists                     ' these are all variables used in the path-ing process below
Dim PATHSNode                     '                   
Dim dataPath                      '
Dim appletPath                    '
Dim GetOutofHere 'GetOutofHere is the variable that's the exit flag
Dim oRS 'object recordset doncha know

'Attempt to create a reference to the MSXML DOM
On Error Resume Next
Set pXML = CreateObject("MSXML2.DOMDocument.3.0")
pXML.async = false
If Err.Number <> 0 Then
Application.MessageBox "MSXML is not present on this device.", vbCritical, "No MSXML"
Exit Sub
End If
On Error GoTo 0
'Path to the ArcPadPrefs.apx file
strArcPadAPX = Application.System.Properties("PERSONALFOLDER") & "\My ArcPad\ArcPadPrefs.apx"
'Everytime ArcPad restarts, it removes the processing instruction at the top of the XML document,
'which means that my pXML object can't read the apx file. So here's to editing it as a text file!
Const ForReading = 1 'constants for working with the text files
Const ForWriting = 2
'get FSO out so we can read the 'text' file
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set pXMLtxt = objFSO.OpenTextFile("strArcPadAPX", ForReading)
'read it and close it
strContents = pXMLtxt.ReadAll
objFile.Close
set versiontext = """1.0"""
strFirstLine = "<?xml version=" & "versiontext" & " ?>"
strNewContents = strFirstLine & vbCrLf & strContents
Set pXMLtxt = objFSO.OpenTextFile("strArcPadAPX"), ForWriting)
pXMLtxt.WriteLine strNewContents
pXMLtxt.Close
'Everytime ArcPad restarts, it removes the processing instruction at the top of the XML document,
'which means that my pXML object can't read the apx file.
Set pXMLIntro = pXML.createProcessingInstruction("xml","version=\"1.0\"")  
pXML.insertBefore(pXMLIntro,pXML.childNodes(0)) 
'Read in ArcPadPrefs.apx
If pXML.load("strArcPadAPX") Then
Application.MessageBox "I found arcpad's preferences file on your device, so now I know where everything is."
Else
Application.MessageBox "Can't find arcpad's preferences file. Don't know where anything is."
End If
Set blnExists = pXML.load("strArcPadAPX")
'go down the XML hierarchy from the PREFERENCES to the PATHS element to the nodes of the path element
Set PATHSNode = pXML.SelectSingleNode("/PREFERENCES/PATHS")
'Get the value from the "data" and "applets" attributes and set them as global variables for later use
Application.UserProperties("dataPath") = PATHSNode.getAttribute("data")
Application.UserProperties("appletPath") = PATHSNode.getAttribute("applets")

'here's how we exit this without crashing ArcPad
GetOutofHere = False
If (True = GetOutofHere) Then 
	Application.Quit 'if exit flag is set - then quit
Exit Sub
End If

'opens a recordset and populates it with the people.dbf located in the appletPath folder
Set oRS = Application.CreateAppObject("recordset")
oRS.Open appletPath & "people.dbf"

End Sub

''''this sub fills in the first name from the initials they select in the combobox cboInitials
Sub FillinFullName 

Dim oControls 'is here to fill in my text boxes

Set oControls = ThisEvent.Object.Pages("pgDataRec").Controls 
If oControls("txtFirstName").Value = "" Then
   oControls("txtFirstName").Value = ReturnFirstName
End If
If oControls("txtLastName").Value = "" Then
   oControls("txtLastName").Value = ReturnLastName
End If

Set oControls = nothing
End Sub

'Here are two functions that are used by combobox to return first and last names so the user can 
'verify that the initials match them perhaps inc. the ability to add a name to the dbf and/or edit your entry?

Function ReturnFirstName(sInitials)
	
	Dim oRecord
	Dim oQuery
	Dim DataRecorderFirstName

	oRS.Open appletPath & "people.dbf"
	oQuery = "[INITIALS]="""&sInitials&""""
	oRecord = oRS.Find(oQuery)

	If (oRecord > 0) Then 'found it, yeehaw
  	 oRS.MoveFirst
  	 oRS.Move(oRecord-1)
  	 ReturnFirstName = oRS.Fields("FIRSTNAME").Value
	 Application.UserProperties("DataRecorderFirstName") = oRS.Fields("FIRSTNAME").Value
  	Else
		ReturnFirstName = ""	
	End If

	set oQuery = Nothing
	Set oRecord = Nothing

End Function

Function ReturnLastName(sInitials)
	Dim oRecord
	Dim oQuery

	oRS.Open appletPath & "people.dbf"
	oQuery = "[INITIALS]="""&sInitials&""""
	oRecord = oRS.Find(oQuery)

	If (oRecord > 0) Then 'found it, yeehaw
  	 oRS.MoveFirst
  	 oRS.Move(oRecord-1)
  	 ReturnLastName = oRS.Fields("LASTNAME").Value
	 Application.UserProperties("DataRecorderLastName") = oRS.Fields("LASTNAME").Value
  	Else
		ReturnLastName = ""	
	End If

	Set oQuery = Nothing
	Set oRecord = Nothing

End Function

Sub DataRecForm_OnOk()

Dim pControls
'get page 1 controls 
Set pControls = ThisEvent.Object.Pages("pgDataRec").Controls
'store the selected employee in a user property so it can be retrieved from any script
Application.UserProperties("DataRecorder") = pControls("cboInitials").Value

' let's use a message box to welcome the logged in user
'Application.MessageBox "Hello & DataRecorderInitials & aka , & DataRecorderFirstName, &  DataRecorderLastName, " ", "You are in.", apNewLine, "Happy 'fieldwork!", apOKOnly, "Data recorder signed in"")

Set pControls = Nothing

End Sub

Sub DataRecName_OnValidate()
    '++ get the control ptr
    Dim pControl
    Set pControl = ThisEvent.Object
    
    '++ an employee must be selected
    If (-1 = pControl.ListIndex) Then
		ThisEvent.Result = False
		ThisEvent.MessageText = "The dropdown is up there. If you need to add yourself, do so on the next page, ok?"
    End If
End Sub

Sub DataRecForm_OnQueryCancel()
	'++ prompt the user before cancelling the DataRec and exiting the application
	Dim answer
	answer = Application.MessageBox("You must exist in order to do field work. Do you want to be outside or not?" ,apYesNo,"Quit?")
	If (apYes = answer) Then
		ThisEvent.Result = False
	End If
End Sub

Sub DataRecForm_OnCancel()
	'user cancelled - set flag that we need to exit
	GetOutofHere = True
End Sub