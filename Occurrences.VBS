Option Explicit

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

    'Randomize
    getInvasivesDBKey = Mid(CStr(DatePart("yyyy", Now()) & _
            Mon1 & DatePart("m", Now()) & _
            Day1 & DatePart("d", Now()) & _
            Hour1 & DatePart("h", Now()) & _
            Min1 & DatePart("n", Now()) & _
            Sec1 & DatePart("s", Now())), 3) & _
            CStr(Int((9999 * Rnd()) + 1))
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GetLat()
	Dim ObjRS, objSH, dHectares

'     If Map.Layers("Occurrences") Is Nothing Then
	If Map.SelectionLayer Is Nothing Then
		GetLat = -1
 		Exit Function
	End If

'    Set objRS = Map.Layers("Occurrences").Records
    Set objRS = Map.SelectionLayer.Records
	objRS.Bookmark = Map.SelectionBookmark
	GetLat = objRS.Fields.Shape.Y	
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function GetLon()
	Dim ObjRS, objSH, dHectares
'	If Map.Layers("Occurrences") Is Nothing Then
	If Map.SelectionLayer Is Nothing Then
		GetLon = -1
		Exit Function
	End If
'    Set objRS = Map.Layers("Occurrences").Records
	Set objRS = Map.SelectionLayer.Records
	objRS.Bookmark = Map.SelectionBookmark
GetLon = objRS.Fields.Shape.X
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function ReturnCName(sSciName)
	Dim oRS
	Dim oRecord
	Dim oQuery
	Dim sPath
	Dim oSelLayer
	Dim iEndSlash	

	Set oSelLayer = Map.Layers("Occurrences")
	Set oRS = Application.CreateAppObject("recordset")
	
	sPath = oSelLayer.FilePath
	iEndSlash = InStrRev(sPath, "\")
	sPath = mid(sPath, 1, iEndSlash)
	oRS.Open sPath & "plants.dbf"
	oQuery = "[SCINAME] = """&sSciName&""""
	oRecord = oRS.Find(oQuery)

	If (oRecord > 0) Then 'found
  	 oRS.MoveFirst
  	 oRS.Move(oRecord-1)
	 ReturnCName = oRS.Fields("COMNAME").Value
	 Else
	 ReturnCName = ""
	End If
	
	Set oRS = Nothing
	Set oSelLayer = Nothing
	Set oRecord = Nothing

End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ValidateAreas()
	
	Dim pPage
	Set pPage = Map.Layers("Occurrences.shp").Forms("EDITFORM").Pages("PgLocation")
	'Check that at least one area is selected
	
	If pPage.Controls("cboArea1").Value = "" And pPage.Controls("cboArea2").Value = "" And pPage.Controls("cboArea3").Value = "" Then
	 ThisEvent.Result = False
	 ThisEvent.MessageText = "You must pick at least one area."
	 ThisEvent.MessageType = 48
	End If
	
	'Primary areas.........
	Dim iPrimSum
	Dim fieldName1, fieldName2
	Dim int
	
	iPrimSum = 0
	for int = 1 to 4
	 If pPage.Controls("cboArea" & int).ListIndex > 0 Then  'fieldName1 = "chkPrimary" & int; an area is specified		
		If pPage.Controls("chkPrimary" & int).Value = 1 or pPage.Controls("chkPrimary" & int).Value = True Then
		 iPrimSum = iPrimSum + 1
		End If		
	 End If
	next
	
	'Check that at least one area is selected as primary area
	If iPrimSum = 0 Then
	 ThisEvent.Result = False
	 ThisEvent.MessageText = "You must specify an area as primary area."
	 ThisEvent.MessageType = 48
	End If

'Check that more than one area isn't checked as primary area
	If iPrimSum > 1 Then
	 ThisEvent.Result = False
	 ThisEvent.MessageText =  "You can only specify one area as primary area."
	 ThisEvent.MessageType = 48
	End If

End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub InitializeOccurrencesEditForm 'sub called by the onload event within the Occurrences EDITFORM
	'automatically populates some fields
	
	Dim pLayer 'current layer
	Dim oControls 'object controls variable
	
	Set pLayer = Map.Layers("Occurrences.shp")
	Set oControls = pLayer.Forms("EDITFORM").Pages("pgBasic").Controls
	If oControls("txtWOKey").Value = "" Then
		oControls("txtWOKey").Value = getInvasivesDBKey
	End If
	
	oControls("txtLatitude").Value = round(GetLat,8)
	If oControls("txtLatitude").Value = "-1" Then
		oControls("txtLatitude").Value = "See 'Geography'"		
		oControls("txtLongitude").Value = "tab"
	Else
	   oControls("txtLongitude").Value = round(GetLon,8)
	End If
	
	If oControls("txtInitials").Value = "" Then
		oControls("txtInitials").Value = Application.UserProperties("DataRecorder")
	Else
		oControls("txtInitials").Value = "no recorder"
	End If
	
	If oControls("txtAccuracy").Value = "" Then
		oControls("txtAccuracy").Value = Application.UserProperties("GPSAccuracy")
	Else
		oControls("txtAccuracy").Value = "na"
	End If
	
	If oControls("txtPDOP").Value = "" Then
		oControls("txtPDOP").Value = GPS.Properties("PDOP")
	End If
	
	If oControls("txtSatellites").Value = "" Then
		oControls("txtSatellites").Value = GPS.Properties("SATS_USED")
	End If
	
	If oControls("txtQuality").Value = "" Then
		oControls("txtQuality").Value = GPS.Properties("QUALITY")
	End If
	
	If oControls("txtDifferentialAge").Value = "" Then
		oControls("txtDifferentialAge").Value = GPS.Properties("DIFF_AGE")
	End If
	
	Set pLayer = Nothing
	Set oControls = Nothing
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub TurnOffApplet 'sub called by the arcpad layer function onclose
	
	Dim pToolbars
	Dim A
	
	Set pToolbars = Application.Toolbars
	For Each A in pToolbars
	 If A.Name = "tlbInvasivesDB" Then
	  pToolbar = Application.Toolbars("tlbInvasivesDB")
	  pToolbar.Item("btnOccur").Enabled = False
	  pToolbar.Item("btnAssess").Enabled = False
	  pToolbar.Item("btnTreat").Enabled = False
	  pToolbar.Item("btnOnOff").Checked = False
	  Set pToolbar = Nothing
	 End If
	Next

End Sub
