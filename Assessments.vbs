Option Explicit
'This script is intended for use on the polygon layer 
'of the geodatabase of invasive species occurrences on 
'The Nature Conservancy's Disney Wilderness Preserve. 
'the original script was written by someone at TNC 
'in the Washington, DC office for the WIMS MS Access
'database. Debi Tharp-Stone asked me to migrate the database
'to Esri's file geodatabase format and make these scripts work
'with the new geodatabase. 
'I don't know who will be maintaining this script.
'Valerie Anderson, volunteer, TNC-DWP, Kissimmee, FL, 
'July 2014 - September 2014 vca@hush.com

Sub PopulateWO()
	'aka the "Populate the Weed Occurrence Key (sWOKey) box sub procedure"
	'Use this sub procedure to screen the existing weed polygons 
	'to find only those occurrences that are within the envelope 
	'of the assessment and populate the cboWOKey box accordingly.
	
	'Find the envelope of the shape drawn, declare our variables!
	Dim oLayerRS 'aka occurrence layer recordset
	Dim oExtentRect 'aka occurrence extent rectangle
	Set oLayerRS = Map.SelectionLayer.Records 'set the occurrence layer recordset to the current selected feature
	
	oLayerRS.Bookmark = Map.SelectionBookmark 'bookmark the occurrence layer recordset to match the bookmark of the currently selected feature
	Set oExtentRect = oLayerRS.Fields.Shape.Extent 'set the occurrence extent rectangle variable 
	
	'Create the recordset and filter by envelope
	Dim pOccRS
	Dim pOccLayer
	Dim lRec
	Dim lBookmark
	
	Set pOccLayer = Map.Layers("Occurrences") 'set the variable pOccLayer as the "Occurrences" map layer
	Set pOccRS = pOccLayer.Records 'retrieve the records from the map layer and set them to the variable pOccRS
	lBookmark = 0
	lRec = pOccRS.Find("[WOKEY] <> 0", oExtentRect, lBookmark)
	
	Dim sWOKey
	
	If Not oLayerRS.Fields("WOKEY").Value = "" Then
		sWOKey = oLayerRS.Fields("WOKEY").Value
	End If
	
	'If no records within extent, then abort creation of feature
	If lRec = 0 Then
		Dim iYN
		If sWOKey = "" Then
			iYN = Application.Messagebox("No occurrences fall within the extent of this assessment boundary. Do you want to see all occurrence records in the map extent? No will cancel feature.", vbYesNo, "Assessment Creation Error")
			If iYN = 7 Then
				Map.Layers("Assessments").Forms("EDITFORM").Close(False)
				Set oLayerRS = Nothing
				Set pOccRS = Nothing
				Set pOccLayer = Nothing
				Set oExtentRect = Nothing
				Exit Sub
			Else
				Set oExtentRect = Map.Extent
				lRec = pOccRS.Find("[WOKEY] <> 0", oExtentRect, lBookmark)
			End If
		Else
			Set oExtentRect = Map.Extent
			lRec = pOccRS.Find("[WOKEY] <> 0", oExtentRect, lBookmark)
		End If
	End If


	Dim i, iIndex
	Dim str
	
	Dim arr(400)
	i = 0
	'
	'   store locations/keys in the array, delimited by a '^'
	'
	Do While Not lRec = 0
		str = pOccRS.Fields("ALTLOCINFO")	
		arr(i) = str & "^" & pOccRS.Fields("WOKEY")
		i = i + 1
		lBookmark = lRec
		lRec = pOccRS.Find("[WOKEY] > 0", oExtentRect, lBookmark)
	Loop
	
	Dim arrMax
	arrMax = i - 1
	'
	'   sort the array (by location)
	'
	Dim j
	Dim temp
	dim arg1
	dim arg2
	dim res1
	
	If arrMax > 0 Then
		For i = arrMax - 1 To 0 Step -1
			For j = 0 To i
				arg1 = left(arr(j),instr(arr(j),"^") - 1) 
				arg2 = left(arr(j+1),instr(arr(j+1),"^") - 1)

				res1 = StrComp(arg1,arg2,vbTextCompare)
				if res1 = 1 then
					temp = arr(j + 1)
					arr(j + 1) = arr(j)
					arr(j) = temp	
				End If
			Next
		Next
	End If
	'
	'   build combo-box with sorted Locations and WOKeys - note when you match the pre-existing value of WOKey
	'
	Dim theLoc
	Dim theKey
	Dim delimiterPos
	Dim oCBO
	Set oCBO = Map.Layers("Assessments").Forms("EDITFORM").Pages("pgLoc").Controls("cboOccurrence")
	oCBO.Clear   'Clear contents of the control

	Dim foundIt
	foundIt = "N"
	iIndex = 0

	For i = 0 To 399
		If arr(i) = "" Then
			Exit For
		End If

		delimiterPos = instr(arr(i),"^")
		theLoc = left(arr(i),delimiterPos - 1)
		theKey = right(arr(i), len(arr(i)) - delimiterPos)
		'
		Call oCBO.AddItem(theKey ,theLoc) ' show Location, but value is WOKey
		If theKey = sWOKey Then
			iIndex = i
			foundIt = "Y"
		End If
	Next
	If foundIt = "Y" Then
		oCBO.ListIndex  = iIndex 
	End If


	'Clean-up
	'Set oCBO = Nothing
	Set oLayerRS = Nothing
	Set pOccRS = Nothing
	Set pOccLayer = Nothing
	Set oExtentRect = Nothing
	
	
	If sWOKey <> "" Then
		Call FillWOFields(sWOKey)
	Else
		If Map.Layers("Assessments").Forms("EDITFORM").Pages("pgLoc").Controls("cboOccurrence").value <> "" Then
			Call FillWOFields(Map.Layers("Assessments").Forms("EDITFORM").Pages("pgLoc").Controls("cboOccurrence").value)
		End If
	End If

End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ValidateWeed()
Dim pPage
Set pPage = Map.Layers("Assessments").Forms("EDITFORM").Pages("pgLoc")	
If pPage.Controls("txtWeedName") = "" Then
	Application.Messagebox "Please specify the related Occurrence from the pick-list"
End If
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub FillWOFields(sWOKey)
If sWOKey = "" Then
	Application.Messagebox "Please specify the related Occurrence from the pick-list"
	Exit Sub
End If

Dim pOccRS
Dim sPath
Dim iEndSlash
Dim pSelLayer
Dim sQuery
Dim pRec
Dim pPage

Set pPage = Map.Layers("Assessments").Forms("EDITFORM").Pages("pgLoc")

Set pSelLayer = Map.Layers("Assessments")
sPath = pSelLayer.FilePath
iEndSlash = InStrRev(sPath, "\")
sPath = mid(sPath, 1, iEndSlash)
Set pSelLayer = Nothing

Set pOccRS = Application.CreateAppObject("Recordset")
pOccRS.Open sPath & "Occurrences.dbf"
sQuery = "[WOKEY] = """ & sWOKey & """"
pRec = pOccRS.Find(sQuery)

If (pRec > 0) Then 'should always be
	pOccRS.MoveFirst
	pOccRS.Move(pRec - 1)
	pPage.Controls("txtWeedName").Value = pOccRS.Fields("WEEDNAME").Value
	pPage.Controls("txtDataRec").Value = pOccRS.Fields("DATAREC").Value
End If

Set pOccRS = Nothing
Set pPage = Nothing

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CalcSize()

Dim pPage
Set pPage = Map.Layers("Assessments").Forms("EDITFORM").Pages("pgSize")

If pPage.Controls("chkCalcFromShape").Value = False Then 'This means the box was unchecked
	pPage.Controls("txtLength").Enabled = True
	pPage.Controls("txtWidth").Enabled = True
	pPage.Controls("cboUOM").Enabled = True
	pPage.Controls("txtAcres").Enabled = True
Else  'The box was checked and we need to calc the area
	Dim pRS
	Set pRS = Map.SelectionLayer.Records
	pRS.Bookmark = Map.SelectionBookmark

	'Return the polygon
	Dim pPoly
	Set pPoly = pRS.Fields.Shape

	'Get area of the polygon - will be sq meters if UTM or geographic, sq ft. otherwise
	Dim dArea
	Dim dAcres
	dArea = pPoly.Area
	dAcres = dArea * 0.000247
    pPage.Controls("txtAcres").Value = round(dAcres,4)

	'Get the Coordinate System and the projection unit
	Dim pCS
	Dim dPrjUnit
	Set pCS = pPoly.CoordinateSystem
	dPrjUnit = pCS.LookupConstant(pCS.ProjectionUnit) 'conversion factor between CS units and meters

	'Get the length and width of the polygon
	Dim dWidth
	Dim dHeight
	dWidth = abs(pPoly.Extent.Left - pPoly.Extent.Right) * dPrjUnit
	dHeight = abs(pPoly.Extent.Top - pPoly.Extent.Bottom) * dPrjUnit

	'Write the calc'd values to the form
	pPage.Controls("txtLength").Value = round(dHeight,0)
	pPage.Controls("txtWidth").Value = round(dWidth,0)
	pPage.Controls("cboUOM").Value = "m"
	pPage.Controls("txtAcres").Value = round(dAcres,4)
	'Enable the controls
	pPage.Controls("txtLength").Enabled = False
	pPage.Controls("txtWidth").Enabled = False
	pPage.Controls("cboUOM").Enabled = False
	pPage.Controls("txtAcres").Enabled = False

	'Intermediate clean-up
	Set pPoly = Nothing
	Set pCS = Nothing
	Set pRS = Nothing

End If

Set pPage = Nothing

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function getInvasivesDBKey()
    '
    '  this is the shared routine for generating a (hopefully) unique record key
    '      it's a combination of the current Date/Time + a random 4-digit number
    '
    Dim Mon1
    Dim Day1
    Dim Hour1
    Dim Min1
    Dim Sec1
    '
    '   make sure we get 2-digit entries for each of these
    '
    If Len(DatePart("m", Now())) = 1 Then
        Mon1 = "0"
    End If
    If Len(DatePart("d", Now())) = 1 Then
        Day1 = "0"
    End If
    If Len(DatePart("h", Now())) = 1 Then
        Hour1 = "0"
    End If
    If Len(DatePart("n", Now())) = 1 Then
        Min1 = "0"
    End If
    If Len(DatePart("s", Now())) = 1 Then
        Sec1 = "0"
    End If
    '
    Randomize
    getInvasivesDBKey = Mid(CStr(DatePart("yyyy", Now()) & _
            Mon1 & DatePart("m", Now()) & _
            Day1 & DatePart("d", Now()) & _
            Hour1 & DatePart("h", Now()) & _
            Min1 & DatePart("n", Now()) & _
            Sec1 & DatePart("s", Now())), 3) & _
            CStr(Int((9999 * Rnd()) + 1))
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub getInvasivesDBKeyApplet

Dim pToolbars
Dim A

Set pToolbars = Application.Toolbars
For Each A in pToolbars
	If A.Name = "tlbInvasivesDB" Then
		Dim pToolbar
		Set pToolbar = Application.Toolbars("tlbInvasivesDB")
		pToolbar.Item("btnOccur").Enabled = False
		pToolbar.Item("btnAssess").Enabled = False
		pToolbar.Item("btnTreat").Enabled = False
		pToolbar.Item("btnOnOff").Checked = False
		Set pToolbar = Nothing
	End If
Next
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ValidateTime()
	Dim pPage
	Dim holdStart
	Dim holdEnd
	Dim holdHH
	Dim holdMM
	Set pPage = Map.Layers("Assessments").Forms("EDITFORM").Pages("pgTime")
	holdStart = pPage.Controls("txtTimeStart").Value 
	holdEnd = pPage.Controls("txtTimeEnd").Value
	If holdStart <> "" And holdEnd <> "" Then  ' both are present
		If instr(holdStart,":") > 0 Or instr(holdEnd,":") > 0 Then
			Application.Messagebox "Please key Times in 24-hour format (no colons)"
		Else
			If Cint(holdEnd) < cint(holdStart) Then
				Application.Messagebox "Start time must be prior to End time"
			Else
				If holdStart = "0" And holdEnd = "0" Then  ' no input
				Else
					If len(holdStart) > 4 Or len(holdEnd) > 4 _
					Or len(holdStart) < 3 Or len(holdEnd) < 3 Then
						Application.Messagebox "Please key times in 24-hour format (HHMM)"
					Else  ' make sure that we have 4 digit times
						If len(holdStart) = 3 Then
							holdStart = "0" & holdStart
							pPage.Controls("txtTimeStart").Value = holdStart 
						End If						
						If len(holdEnd) = 3 Then
							holdEnd = "0" & holdEnd
							pPage.Controls("txtTimeEnd").Value = holdEnd 
						End If
						'
						holdHH = left(holdStart,2)
						holdMM = right(holdStart,2)
						If holdHH >= "00" And holdHH <= "23" _
						And holdMM >= "00" And holdMM <= "59" Then
							holdHH = left(holdEnd,2)
							holdMM = right(holdEnd,2)
							If holdHH >= "00" And holdHH <= "23" _
							And holdMM >= "00" And holdMM <= "59" Then
							Else
								Application.Messagebox "Invalid End time"
							End If
						Else
							Application.Messagebox "Invalid Start time"
						End If
					End If
				End If
			End If 
		End If
	Else
		If holdStart <> "" Or holdEnd <> "" Then  ' only one is present
			Application.Messagebox "Please supply both Start and End times"
		End If
	End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ValidateNumeric(holdType, holdValue)
	If holdValue <> "" Then	
		If Not Isnumeric(holdValue) Then
			Application.Messagebox holdType & " must be numeric"
		End If
	End If
End Sub

''''''''''''''''''''''''''''''''''
Sub CheckCover()
	Dim pPage
	Set pPage = Map.Layers("Assessments").Forms("EDITFORM").Pages("pgCovDen")

	If Not Isnumeric(pPage.Controls("txtPctCover").Value) Then
		pPage.Controls("txtPctCover").Value = 0	
	End If

	If pPage.Controls("txtPctCover").Value = 0 Then
        Select Case pPage.Controls("cboCovClass").Value
            Case "< 1%"
                pPage.Controls("txtPctCover").Value = 0.5
            Case "1 - 10%"
                pPage.Controls("txtPctCover").Value = 5
            Case "11 - 25%"
                pPage.Controls("txtPctCover").Value = 18
            Case "26 - 50%"
                pPage.Controls("txtPctCover").Value = 38
            Case "51 - 100%"
                pPage.Controls("txtPctCover").Value = 75
            Case Else
                Application.Messagebox "unrecognized value for Cover Class"
        End Select
	End If
	Set pPage = Nothing
End Sub
''''''''''''''''''''''''''''''''
Sub ObjectControlsAssessment 'sub called by the onload event within the Assessments EDITFORM

Dim objControls

Set objControls = ThisEvent.Object.Pages("pgLoc").Controls
If  objControls("txtVisitID").Value = "" Then
    objControls("txtVisitID").Value = getInvasivesDBKey
End If
    objControls("txtDateMod").Value = Now
If Not objControls("txtDate_").Value = "" Then
   Dim dDate
    dDate = cDate(objControls("txtDate_").Value)
    objControls("dtpDate").Value = dDate
End If
Set objControls = Nothing
If Not Map.PointerMode = "modeidentify" Then
	Call PopulateWO
End If

End Sub
''''''''''''''''''''''''''''''''
Sub TurnOffApplet 'sub called by the arcpad layer function onclose

Dim pToolbars
Dim A

Set pToolbars = Application.Toolbars
For Each A in pToolbars
	If A.Name = "tlbInvasivesDB" Then
		Dim pToolbar
		Set pToolbar = Application.Toolbars("tlbInvasivesDB")
		pToolbar.Item("btnOccur").Enabled = False
		pToolbar.Item("btnAssess").Enabled = False
		pToolbar.Item("btnTreat").Enabled = False
		pToolbar.Item("btnOnOff").Checked = False
		Set pToolbar = Nothing
	End If
Next

End Sub