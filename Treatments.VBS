Option Explicit

Sub DisablePages()

Dim sTreatType
sTreatType = ThisEvent.Object.Pages("pgBasic").Controls("cboTreatType").Value

If Not sTreatType = "" Then
	Select Case sTreatType
		Case "Chemical"
			Call EnableDisable("pgMech", False)
			Call EnableDisable("pgFire", False)
			Call EnableDisable("pgBio1", False)
			Call EnableDisable("pgGraz", False)	
		Case "Mechanical"
			Call EnableDisable("pgChem1", False)
			Call EnableDisable("pgChem2", False)
			Call EnableDisable("pgChem3", False)
			Call EnableDisable("pgFire", False)
			Call EnableDisable("pgBio1", False)
			Call EnableDisable("pgGraz", False)		
		Case "Fire"
			Call EnableDisable("pgChem1", False)
			Call EnableDisable("pgChem2", False)
			Call EnableDisable("pgChem3", False)
			Call EnableDisable("pgMech", False)
			Call EnableDisable("pgBio1", False)
			Call EnableDisable("pgGraz", False)		
		Case "BioAgent"
			Call EnableDisable("pgChem1", False)
			Call EnableDisable("pgChem2", False)
			Call EnableDisable("pgChem3", False)
			Call EnableDisable("pgMech", False)
			Call EnableDisable("pgFire", False)
			Call EnableDisable("pgGraz", False)
		Case "Grazing"
			Call EnableDisable("pgChem1", False)
			Call EnableDisable("pgChem2", False)
			Call EnableDisable("pgChem3", False)
			Call EnableDisable("pgMech", False)
			Call EnableDisable("pgFire", False)
			Call EnableDisable("pgBio1", False)
	End Select
End If

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub EnableControls()
'This sub enables and disables the controls on certain pages of the treatment form based
'on what value the user selects for the treatment type. For example, if the user selects
'Chemical as the treatment type, then the controls on the Chemical page are enabled and
'the controls on all the other treatment type pages are disabled. This prevents the user
'from entering erroneous information for a particular treatment type

Dim sTreatType

sTreatType = ThisEvent.Object.Value

Select Case sTreatType
	Case "Chemical"
		Call EnableDisable("pgChem1", True)
		Call EnableDisable("pgChem2", True)
	    Call EnableDisable("pgChem3", True)
		Call EnableDisable("pgMech", False)
		Call EnableDisable("pgFire", False)
		Call EnableDisable("pgBio1", False)
		Call EnableDisable("pgGraz", False)

	Case "Mechanical"
		Call EnableDisable("pgChem1", False)
		Call EnableDisable("pgChem2", False)
	   Call EnableDisable("pgChem3", False)
		Call EnableDisable("pgMech", True)
		Call EnableDisable("pgFire", False)
		Call EnableDisable("pgBio1", False)
		Call EnableDisable("pgGraz", False)
	
	Case "Fire"
		Call EnableDisable("pgChem1", False)
		Call EnableDisable("pgChem2", False)
		Call EnableDisable("pgChem3", False)		
		Call EnableDisable("pgMech", False)
		Call EnableDisable("pgFire", True)
		Call EnableDisable("pgBio1", False)
		Call EnableDisable("pgGraz", False)
	
	Case "BioAgent"
		Call EnableDisable("pgChem1", False)
		Call EnableDisable("pgChem2", False)
		Call EnableDisable("pgChem3", False)
		Call EnableDisable("pgMech", False)
		Call EnableDisable("pgFire", False)
		Call EnableDisable("pgBio1", True)
		Call EnableDisable("pgGraz", False)
	
	Case "Grazing"
		Call EnableDisable("pgChem1", False)
		Call EnableDisable("pgChem2", False)
		Call EnableDisable("pgChem3", False)
		Call EnableDisable("pgMech", False)
		Call EnableDisable("pgFire", False)
		Call EnableDisable("pgBio1", False)
		Call EnableDisable("pgGraz", True)
		
End Select

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub EnableDisable(sPageName, bEnValue)

	Dim pPage
	Dim iNumControls
	Dim i
	
	Set pPage = Map.Layers("Treatments").Forms("EDITFORM").Pages(sPageName)
	iNumControls = pPage.Controls.Count
	If iNumControls > 0 Then
		For i = 1 To iNumControls
			pPage.Controls.Item(i).Enabled = bEnValue
		Next
	End If
	
	Set pPage = Nothing

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub CalcAcres

Dim pPage
Set pPage = Map.Layers("Treatments").Forms("EDITFORM").Pages("pgAreas")

Dim pRS
Set pRS = Map.SelectionLayer.Records
pRS.Bookmark = Map.SelectionBookmark
	
'Return the polygon
Dim pPoly
Set pPoly = pRS.Fields.Shape
	
'Get the area of the polygon - will be sq meters if UTM or geographic
Dim dArea
Dim dAcres
dArea = pPoly.Area
dAcres = dArea * 0.000247

'Write acres to the form (first 'Acres Treated' slot)
pPage.Controls("txtTotAcTreat1").Value = round(dAcres,4)

'Clean-up
Set pRS = Nothing
Set pPoly = Nothing
Set pPage = Nothing
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub PopulateWO
'Use this sub to screen the occurrences to only
'those that are within the envelope of the treatment
'and populate the cboWOKey boxes accordingly

'Find the envelope of the polygon
Dim pLayerRS
Dim pExtentRect
Set pLayerRS = Map.SelectionLayer.Records
pLayerRS.Bookmark = Map.SelectionBookmark
Set pExtentRect = pLayerRS.Fields.Shape.Extent

'Create the recordset and filter by the envelope
Dim pOccRS
Dim pOccLayer
Dim lRec
Dim lBookmark

Set pOccLayer = Map.Layers("Occurrences")
Set pOccRS = pOccLayer.Records
lBookmark = 0
lRec = pOccRS.Find("[WOKEY] <> 0", pExtentRect, lBookmark)

'If no records within the extent, then abort creation of feature
If lRec = 0 Then
	Dim iYN
	If Map.Layers("Treatments").Forms("EDITFORM").Pages("pgAreas").Controls("cboArea1").Value = "" Then
		iYN = Application.Messagebox("No occurrences fall within the extent of this treatment boundary. Do you want to see all occurrence records in the map extent? No will cancel feature.", vbYesNo, "Treatment Creation Error")
		If iYN = 7 Then
			Map.Layers("Treatments").Forms("EDITFORM").Close(False)
			Set pLayerRS = Nothing
			Set pOccRS = Nothing
			Set pOccLayer = Nothing
			Set pExtentRect = Nothing
			Exit Sub
		Else
			Set pExtentRect = Map.Extent
			lRec = pOccRS.Find("[WOKEY] <> 0", pExtentRect, lBookmark)
		End If
	Else
		Set pExtentRect = Map.Extent
		lRec = pOccRS.Find("[WOKEY] <> 0", pExtentRect, lBookmark)
	End If
End If

'Get references to the controls we want to update
Dim pCBO1, pCBO2, pCBO3, pCBO4,pCBO5, pCBO6,pCBO7, pCBO8, pCBO9, pCBO10,pCBO11, pCBO12
Set pCBO1 = Map.Layers("Treatments").Forms("EDITFORM").Pages("pgWeedA").Controls("cboWeed1")
Set pCBO2 = Map.Layers("Treatments").Forms("EDITFORM").Pages("pgWeedA").Controls("cboWeed2")
Set pCBO3 = Map.Layers("Treatments").Forms("EDITFORM").Pages("pgWeedB").Controls("cboWeed3")
Set pCBO4 = Map.Layers("Treatments").Forms("EDITFORM").Pages("pgWeedB").Controls("cboWeed4") 
Set pCBO5 = Map.Layers("Treatments").Forms("EDITFORM").Pages("pgWeedC").Controls("cboWeed5")
Set pCBO6 = Map.Layers("Treatments").Forms("EDITFORM").Pages("pgWeedC").Controls("cboWeed6")
Set pCBO7 = Map.Layers("Treatments").Forms("EDITFORM").Pages("pgWeedD").Controls("cboWeed7")
Set pCBO8 = Map.Layers("Treatments").Forms("EDITFORM").Pages("pgWeedD").Controls("cboWeed8")
Set pCBO9 = Map.Layers("Treatments").Forms("EDITFORM").Pages("pgWeedE").Controls("cboWeed9")
Set pCBO10 = Map.Layers("Treatments").Forms("EDITFORM").Pages("pgWeedE").Controls("cboWeed10")
Set pCBO11 = Map.Layers("Treatments").Forms("EDITFORM").Pages("pgWeedF").Controls("cboWeed11")
Set pCBO12 = Map.Layers("Treatments").Forms("EDITFORM").Pages("pgWeedF").Controls("cboWeed12")
pCBO1.Clear
pCBO2.Clear
pCBO3.Clear
pCBO4.Clear
pCBO5.Clear
pCBO6.Clear
pCBO7.Clear
pCBO8.Clear
pCBO9.Clear
pCBO10.Clear
pCBO11.Clear
pCBO12.Clear

'Populate the comboboxes
Dim theLoc
Call pCBO1.AddItem("","<none>")  ' this permits them to erase an entry if they change their mind
Call pCBO2.AddItem("","<none>")
Call pCBO3.AddItem("","<none>")
Call pCBO4.AddItem("","<none>")
Call pCBO5.AddItem("","<none>")
Call pCBO6.AddItem("","<none>")
Call pCBO7.AddItem("","<none>")
Call pCBO8.AddItem("","<none>")
Call pCBO9.AddItem("","<none>")
Call pCBO10.AddItem("","<none>")
Call pCBO11.AddItem("","<none>")
Call pCBO12.AddItem("","<none>")
Do While Not lRec = 0	
	theLoc = pOccRS.Fields("ALTLOCINFO")
	Call pCBO1.AddItem(pOccRS.Fields("WOKEY"), theLoc)
	Call pCBO2.AddItem(pOccRS.Fields("WOKEY"), theLoc)
	Call pCBO3.AddItem(pOccRS.Fields("WOKEY"), theLoc)
	Call pCBO4.AddItem(pOccRS.Fields("WOKEY"), theLoc)
	Call pCBO5.AddItem(pOccRS.Fields("WOKEY"), theLoc)
	Call pCBO6.AddItem(pOccRS.Fields("WOKEY"), theLoc)
	Call pCBO7.AddItem(pOccRS.Fields("WOKEY"), theLoc)
	Call pCBO8.AddItem(pOccRS.Fields("WOKEY"), theLoc)
	Call pCBO9.AddItem(pOccRS.Fields("WOKEY"), theLoc)
	Call pCBO10.AddItem(pOccRS.Fields("WOKEY"), theLoc)
	Call pCBO11.AddItem(pOccRS.Fields("WOKEY"), theLoc)
	Call pCBO12.AddItem(pOccRS.Fields("WOKEY"), theLoc)
		
	lBookmark = lRec
	lRec = pOccRS.Find("[WOKEY] > 0", pExtentRect, lBookmark)
Loop

'Clean-up
Set pCBO1 = Nothing
Set pCBO2 = Nothing
Set pCBO3 = Nothing
Set pCBO4 = Nothing
Set pCBO5 = Nothing
Set pCBO6 = Nothing
Set pCBO7 = Nothing
Set pCBO8 = Nothing
Set pCBO9 = Nothing
Set pCBO10 = Nothing
Set pCBO11 = Nothing
Set pCBO12 = Nothing
Set pLayerRS = Nothing
Set pOccRS = Nothing
Set pOccLayer = Nothing
Set pExtentRect = Nothing

End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub FillWOFields(sWOKey, iWeedNum, sPage)
	Dim pOccRS
	Dim sPath
	Dim iEndSlash
	Dim pSelLayer
	Dim sQuery
	Dim pRec
	Dim pPage

	Set pPage = Map.Layers("Treatments.shp").Forms(1).Pages(sPage)

	Set pSelLayer = Map.Layers("Treatments.shp")
	sPath = pSelLayer.FilePath
	iEndSlash = InStrRev(sPath, "\")
	sPath = mid(sPath, 1, iEndSlash)
	Set pSelLayer = Nothing

	If sWOKey = "" Then  ' they are clearing out prior selection
		pPage.Controls("txtWeed" & iWeedNum & "Name").Value = ""
		pPage.Controls("txtWeed" & iWeedNum & "Pct").Value = 0
		Exit Sub
	Else
		Set pOccRS = Application.CreateAppObject("Recordset")
		pOccRS.Open sPath & "Occurrences.dbf"
		sQuery = "[WOKEY] = """ & sWOKey & """"
		pRec = pOccRS.Find(sQuery)

		If (pRec > 0) Then 'should always be
			pOccRS.MoveFirst
			pOccRS.Move(pRec - 1)

			pPage.Controls("txtWeed" & iWeedNum & "Name").Value = pOccRS.Fields("WEEDNAME").Value
			If Map.PointerMode = "modeidentify" Then
				pPage.Controls("cboWeed" & iWeedNum).Clear
				pPage.Controls("cboWeed" & iWeedNum).Additem sWOKey, pOccRS.Fields("ALTLOCINFO").Value
				pPage.Controls("cboWeed" & iWeedNum).ListIndex = 0
			End If
		End If
	End If

	Set pOccRS = Nothing
	Set pPage = Nothing
	Set pRec = Nothing
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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

sub getInvasivesDBKeyApplet

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
	Set pPage = Map.Layers("Treatments").Forms("EDITFORM").Pages("pgTime")
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

Sub ComputeQty(holdNum)
	Dim pPage
	dim TankMixUsed
	Dim TankMixUOM
	Dim PctSolHerb
	Dim PctSolAdj
	Dim AdjuvantName
	
	Dim QuantityUsed
	Dim QuantityUOM
	Dim AdjuvantQuantity
	Dim AdjuvantUOM
	
	if holdNum = 1 then
		Set pPage = Map.Layers("Treatments").Forms("EDITFORM").Pages("pgChem2")
	else
		Set pPage = Map.Layers("Treatments").Forms("EDITFORM").Pages("pgChem3")
	end if
	'  set up variables to hold the field names for the current Page
	TankMixUsed = "txtHerb" & holdNum & "Tank"
	TankMixUOM = "cboHerb" & holdNum & "MixUOM" 
	PctSolHerb = "txtHerb" & holdNum & "PctSol" 
	PctSolAdj = "txtAdj" & holdNum & "PctSol" 
	AdjuvantName = "cboAdj" & holdNum
	QuantityUsed = "txtHerbQty" & holdNum
	QuantityUOM = "cboHerb" & holdNum & "UOM"
	AdjuvantQuantity = "txtAdjQty" & holdNum
	AdjuvantUOM = "cboAdj" & holdNum & "UOM"
	'   check for missing data
	if pPage.Controls.item(TankMixUsed).Value = "0" then
		Application.Messagebox "Please specify amount of Tank Mix Used"
		exit sub	
	end if
    If pPage.Controls.item(PctSolHerb).Value = "0" Then
        Application.Messagebox "Please specify the Herbicide % Solution"
		exit sub	
	end if
	If pPage.Controls.item(AdjuvantName).Value <> "" And pPage.Controls.item(PctSolAdj).Value = "0" Then
        Application.Messagebox "Please specify the Adjuvant % Solution"
		exit sub	
	end if
	'
	'   looks good - let's proceed
	'
	pPage.Controls.item(QuantityUOM).Value = "oz"   ' everything gets converted to oz
	pPage.Controls.item(AdjuvantUOM).Value = "oz"
    '
	'  set the conversion factor
	'
	Dim CF
	CF = 0
	Select Case pPage.Controls.item(TankMixUOM).Value
    	Case "oz"
			CF = 1
        Case "qt"  '  1 quart = 32 oz
			CF = 32
        Case "gal"  '  1 gallon = 128 oz
			CF = 128
        Case "ml"   ' 1 ml = 0.033814023 oz
			CF = 0.033814023
        Case "l"    ' 1 l = 33.814022559 oz
			CF = 33.814022559
    End Select
	'
	'   it all boils down to this
	'
    pPage.Controls.item(QuantityUsed).Value = Round((pPage.Controls.item(TankMixUsed).value * CF) * (pPage.Controls.item(PctSolHerb).value / 100), 1)
    If pPage.Controls.item(PctSolAdj).value > 0 Then
    	pPage.Controls.item(AdjuvantQuantity).value = Round((pPage.Controls.item(TankMixUsed).value * CF) * (pPage.Controls.item(PctSolAdj).value / 100), 1)
    End If
End Sub

''''''''''''''''''''''''''''''''''''''''
Sub ObjectControlTreatments 'sub called by the onload event within the Treatments EDITFORM
Dim objControls
Set objControls = ThisEvent.Object.Pages("pgBasic").Controls
If objControls("txtTreatKey").Value = "" Then
   objControls("txtTreatKey").Value = getInvasivesDBKey
End If
objControls("txtDateMod").Value = Now
If Not objControls("txtDate_").Value = "" Then
   Dim dDate
   dDate = cDate(objControls("txtDate_").Value)
   objControls("dtpTreatDate").Value = dDate
End If
Set objControls = Nothing

If Not Map.PointerMode = "modeidentify" Then
	Call PopulateWO
End If

End Sub

''''''''''''''''''''''''''''''''''''''''
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