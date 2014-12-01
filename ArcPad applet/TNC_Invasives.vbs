Option Explicit
'This script is intended for use within the ArcPad applet for the geodatabase of invasive species 
'occurrences on The Nature Conservancy's Disney Wilderness Preserve. The original script was written 
'by someone at TNC in the Washington, DC office for the WIMS MS Access database. Debi Tharp-Stone asked 
'me to migrate the database to Esri's file geodatabase format and make these scripts work with the new geodatabase. 
'I don't know who will be maintaining this script.
'Valerie Anderson, volunteer, TNC-DWP, Kissimmee, FL, 
'July 2014 - October 2014 vca@hush.com

Sub ControlToolbarSettings()
	Dim pCallingControl
	Dim pToolbar
	Dim pOtherControl1
	Dim pOtherControl2
	Dim sCaption
	
	Set pCallingControl = ThisEvent.Object
	Set pToolbar = Application.Toolbars("tlbInvasivesDB")
	
	Select Case pCallingControl.Name 'the variable pCallingControl.Name
		Case "btnOccur"
			Set pOtherControl1 = pToolBar.Item("btnAssess")
			Set pOtherControl2 = pToolBar.Item("btnTreat")
			sCaption = "Occurrences"
		Case "btnAssess"
			Set pOtherControl1 = pToolBar.Item("btnOccur")
			Set pOtherControl2 = pToolBar.Item("btnTreat")
			sCaption = "Assess Polygons"
		Case "btnTreat"
			Set pOtherControl1 = pToolBar.Item("btnAssess")
			Set pOtherControl2 = pToolBar.Item("btnOccur")
			sCaption = "Treat Polygons"
	End Select
	
	If  pCallingControl.Checked = True Then
		pCallingControl.Checked = False
		pToolbar.Caption = "Invasives Database"
	Else
		pCallingControl.Checked = True
		pOtherControl1.Checked = False
		pOtherControl2.Checked = False
		pToolbar.Caption = sCaption
	End If
	
	Set pCallingControl = Nothing
	Set pToolbar = Nothing
	Set pOtherControl1 = Nothing
	Set pOtherControl2 = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub SetLayerEditing(sLayerName)
	
	If ThisEvent.Object.Checked = True Then
		'Turn editing off for all layers
		Dim i
		map.select
		
		
		For i = 1 To Map.Layers.Count
			If Map.Layers(i).CanEdit = True Then
				Map.Layers(i).Editable = False
			End If
		Next

		'Turn on editing for the selected layer
		Map.Layers(sLayerName).Editable = True
		Application.Toolbars("draw").Visible = True

	Else
		
		Map.Layers(sLayerName).Editable = False
		Application.Toolbars("draw").Visible = False

	End If
End Sub
'''''
Sub FillinFullName 
	'this sub fills in the first name from the initials they select in cboInitials
	Dim pControls 'is here to fill in my text boxes
	
	Set pControls = Applet.Forms("frmDataRec").Pages("pgDataRec").Controls
	If pControls("txtFirstName").Value = "" Then
		ReturnFirstName()
		pControls("txtFirstName").Value = ReturnFirstName()
	End If
	If pControls("txtLastName").Value = "" Then
		ReturnLastName()
		pControls("txtLastName").Value = ReturnLastName()
	End If
	
	Set pControls = nothing
End Sub
'''''''
'Here are two functions that are used by combobox to return first and last names so the user can 
'verify that the initials match them perhaps inc. the ability to add a name to the dbf and/or edit your entry?

Function ReturnFirstName()
	
	Dim oRecord
	Dim oQuery
	Dim oRS
	Dim mControls
	Dim DataRecInitials
	
	Set oRS = Application.CreateAppObject("recordset")
	Set mControls = Applet.Forms("frmDataRec").Pages("pgDataRec").Controls
	DataRecInitials = mControls("cboInitials").Value
	oRS.Open "K:\Invasivesdb.transfer\ArcPad applet\People.dbf"
	oQuery = "[INITIALS] = """&DataRecInitials&""""
	oRecord = oRS.Find(oQuery)

	If (oRecord > 0) Then 'found it, yeehaw
  	 oRS.MoveFirst
  	 oRS.Move(oRecord-1)
  	 ReturnFirstName = oRS.Fields("FIRSTNAME").Value
  	Else
		ReturnFirstName = ""	
	End If
	
	oRS.Close

	Set oQuery = Nothing
	Set oRecord = Nothing
	Set oRS = Nothing
	Set mControls = Nothing
	Set DataRecInitials = Nothing

End Function
'''''''
Function ReturnLastName()
	
	Dim oRecord
	Dim oQuery
	Dim oRS
	Dim mControls
	Dim DataRecInitials
	
	Set oRS = Application.CreateAppObject("recordset")
	Set mControls = Applet.Forms("frmDataRec").Pages("pgDataRec").Controls
	DataRecInitials = mControls("cboInitials").Value
	oRS.Open "K:\Invasivesdb.transfer\ArcPad applet\People.dbf"
	oQuery = "[INITIALS]="""&DataRecInitials&""""
	oRecord = oRS.Find(oQuery)

	If (oRecord > 0) Then 'found it, yeehaw
		oRS.MoveFirst
		oRS.Move(oRecord-1)
		ReturnLastName = oRS.Fields("LASTNAME").Value
  	Else
		ReturnLastName = ""	
	End If
	
	oRs.Close

	Set oQuery = Nothing
	Set oRecord = Nothing
	Set mControls = Nothing
	Set DataRecInitials = Nothing
	Set oRS = Nothing

End Function
'''''

Sub frmDataRec_OnOk()
	
	Dim pControls
	Dim oRS
	Dim mControls
	Dim DataRecInitials
	Dim DataRecFirstName
	Dim DataRecLastName
	Dim appletpath 'full path to applet folder
	Dim mainpath 'full path to main folder
	Dim peoplepath 'full path to people.dbf
	Dim iControls
	Dim pPath
	
	appletpath = Preferences.Properties("AppletsPath")
	mainpath = Preferences.Properties("DataPath")
	Application.UserProperties("mPath") = mainpath
	Application.UserProperties("aPath") = appletpath
	Set oRS = Application.CreateAppObject("recordset")
	Set pControls = Applet.Forms("frmDataRec").Pages("pgDataRec").Controls
	Set iControls = Applet.Forms("frmDataRec").Pages("pgDataRec2").Controls
	If iControls("txtInitials2") = "" Then 
		Application.UserProperties("DataRecorder") = pControls("cboInitials").Value
		Application.UserProperties("GPSAccuracy") = pControls("cboAccuracy").Value
	Else
		Set DataRecInitials = iControls("txtInitials2").Value
		Set DataRecFirstName = iControls("txtFirstName2").Value
		Set DataRecLastName = iControls("txtLastName2").Value
		Set peoplepath = appletpath & "\People.dbf"
		Application.UserProperties("pPath") = peoplepath
		oRS.Open peoplepath,2
		oRS.AddNew
		oRS.Fields("INITIALS").Value = DataRecInitials
		oRS.Fields("FIRSTNAME").Value = DataRecFirstName
		oRS.Fields("LASTNAME").Value = DataRecLastName
		oRS.Update
		oRS.Close
		Application.UserProperties("DataRecorder") = iControls("txtInitials2").Value
		Application.UserProperties("GPSAccuracy") = pControls("cboAccuracy").Value
	End If
	
	Set pControls = Nothing
	Set iControls = Nothing
	Set oRS = Nothing
	Set DataRecInitials = Nothing
	Set DataRecLastName = Nothing
	Set appletpath = Nothing
	Set mainpath = Nothing
	Set peoplepath = Nothing
	Set pPath = Nothing

End Sub

'Sub DataRecName_OnValidate()
'    
'	Dim pControl
'	
'	Set pControl = ThisEvent.Object
'	
'	If (-1 = pControl.ListIndex) Then 'have to select an employee
'		ThisEvent.Result = False
'		ThisEvent.MessageText = "The dropdown is up there. If you need to add yourself, do so on the next page, ok?"
'	End If
'
'End Sub
''''''''
Sub frmDataRec_OnQueryCancel()
	'prompt the user before cancelling the DataRec and exiting the application
	
	Dim answer
	
	answer = Application.MessageBox("You must exist in order to do field work. Do you want to be outside or not?" ,apYesNo,"Quit?")
	
	If (apYes = answer) Then
		ThisEvent.Result = False
	End If
	
	Set answer = nothing

End Sub
''''''
Sub frmDataRec2_OnQueryCancel()
	'prompt the user before cancelling the DataRec and exiting the application
	
	Dim answer
	
	answer = Application.MessageBox("You must exist in order to do field work. Do you want to be outside or not?" ,apYesNo,"Quit?")
	
	If (apYes = answer) Then
		ThisEvent.Result = False
	End If
	
	Set answer = nothing

End Sub
'''''''''
Sub frmDataRec_OnCancel()
	
	'user cancelled - set flag that we need to exit
	GetOutofHere = True

End Sub
''''''''
Sub frmDataRec2_OnCancel()
	
	Applet.Forms("frmDataRec2").Close
	
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub OnOffTools() 'sub called by btnOnOff
	
	Dim pToolbar
	Set pToolbar = Application.Toolbars("tlbInvasivesDB")

	If ThisEvent.Object.Checked = True Then
		ThisEvent.Object.Checked = False
		pToolbar.Item("btnOccur").Checked = False
		pToolbar.Item("btnOccur").Enabled = False
		If Map.Layers("Occurrences").Editable = True Then Map.Layers("Occurrences").Editable = False	

		pToolbar.Item("btnAssess").Checked = False
		pToolbar.Item("btnAssess").Enabled = False
		If Map.Layers("Assessments").Editable = True Then Map.Layers("Assessments").Editable = False

		pToolbar.Item("btnTreat").Checked = False
		pToolbar.Item("btnTreat").Enabled = False
		If Map.Layers("Treatments").Editable = True Then Map.Layers("Treatments").Editable = False

		Application.Toolbars("draw").Visible = False

	Else
		'First check for the layers...
		Dim i
		Dim iCounter
		iCounter = 0

		For i = 1 To Map.Layers.Count
			If Map.Layers(i).Name = "Occurrences.shp" Then iCounter = iCounter + 1
			If Map.Layers(i).Name = "Assessments.shp" Then iCounter = iCounter + 1
			If Map.Layers(i).Name = "Treatments.shp" Then iCounter = iCounter + 1
		Next

		If Not iCounter = 3 Then
			Application.Messagebox "At least one of the following layers is not loaded: Occurrences, Assessments, and Treatments", vbInformation, "Layer Missing"
			Set pToolbar = Nothing
			Exit Sub
		End If

		'Then activate the controls
		ThisEvent.Object.Checked = True
		pToolbar.Item("btnOccur").Checked = False
		pToolbar.Item("btnOccur").Enabled = True
		pToolbar.Item("btnAssess").Checked = False
		pToolbar.Item("btnAssess").Enabled = True
		pToolbar.Item("btnTreat").Checked = False
		pToolbar.Item("btnTreat").Enabled = True
	End If

	Set pToolbar = Nothing

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SetControlStatus()
	
	Dim pToolbar
	
	Set pToolbar = Application.Toolbars("tlbInvasivesDB")
	pToolbar.Item("btnOccur").Enabled = False
	pToolbar.Item("btnAssess").Enabled = False
	pToolbar.Item("btnTreat").Enabled = False
	
	Set pToolbar = Nothing

End Sub
