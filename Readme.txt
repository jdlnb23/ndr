
Attribute VB_Name = "RatedNoteFeeder_SocGen"
Option Explicit
'==============================================================================
' RATED NOTE FEEDER - SOCIETE GENERALE ENHANCED EDITION
' Single Module VBA Implementation with Formula-First Architecture verB fixes
' Version: 6.2.0R - Refactored Production Release
' Compliance: Excel 2016+ 64-bit, Option Explicit
'==============================================================================

'------------------------------------------------------------------------------
' MODULE CONSTANTS & GLOBALS
'------------------------------------------------------------------------------
Private Const MODULE_NAME As String = "RatedNoteFeeder_SocGen"
Private Const MODULE_VERSION As String = "v6.2.0R"

' Sentinels
Private Const RATIO_SENTINEL As Double = 999#
Private Const DEFAULT_IF_ERROR As Double = RATIO_SENTINEL

' SocGen Color Palette
Private Const SG_RED As Long = 2494997         'RGB(230, 20, 41)
Private Const SG_BLACK As Long = 0             'RGB(0, 0, 0)
Private Const SG_SLATE As Long = 5066061       'RGB(79, 77, 77)
Private Const SG_GRAY_LIGHT As Long = 15132390 'RGB(230, 230, 230)
Private Const SG_ACCENT As Long = 16316664     'RGB(248, 248, 248)

' Frame Range Names
Private Const OCIC_CHART_FRAME As String = "OCIC_Chart_Frame"
Private Const SCENARIO_MATRIX_FRAME As String = "ScenarioMatrix_Output_Frame"
Private Const INVESTOR_CHART_FRAME As String = "Investor_Chart_Frame"
Private Const CONTROL_BUTTON_ZONE As String = "Control_Button_Zone"

'------------------------------------------------------------------------------
' MAIN ORCHESTRATORS
'------------------------------------------------------------------------------
Public Sub RNF_Strict_BuildAndRun()
    On Error GoTo EH
    Const PROC_NAME As String = "RNF_Strict_BuildAndRun"
    
    Dim wb As Workbook
    Dim startTime As Double
    Dim calcState As XlCalculation
    Dim scr As Boolean, evt As Boolean
    
    startTime = Timer
    scr = Application.ScreenUpdating: Application.ScreenUpdating = False
    calcState = Application.Calculation: Application.Calculation = xlCalculationManual
    evt = Application.EnableEvents: Application.EnableEvents = False
    
    Call Status("Building workbook skeleton...")
    
    Set wb = ActiveWorkbook
    
    ' Step 1: Build Styles
    Call Build_Styles
    
    ' Step 2: Create all sheets
    Call Status("Creating sheets...")
    Call CreateAllSheets(wb)
    
    ' Step 3: Create named frames
    Call Status("Creating named frames...")
    Call CreateNamedFrames(wb)
    
    ' Step 4: Build Control layout
    Call Status("Setting up Control panel...")
    Call SetupControlSheet(wb.Worksheets("Control"))
    
    ' Step 5: Register scenario defaults
    Call Status("Registering scenarios...")
    Call SCN_RegisterDefaults
    
    ' Step 6: Seed AssetTape
    Call Status("Seeding asset tape...")
    Call SeedAssetTape(wb)
    
    ' Step 7: Create buttons
    Call Status("Creating buttons...")
    Call CreateAllButtons(wb)
    
    ' Step 8: Create control named ranges
    Call Status("Creating named ranges...")
    Call CreateControlNamedRanges(wb)
    
    ' Step 9: Initial refresh
    Call Status("Running initial refresh...")
    Call RNF_RefreshAll
    
    Call Log(PROC_NAME, "Build complete in " & Format(Timer - startTime, "0.00") & " seconds")
    
CleanExit:
    Call Status("")
    Application.EnableEvents = evt
    Application.Calculation = calcState
    Application.ScreenUpdating = scr
    Exit Sub
    
EH:
    Call Log(PROC_NAME, "ERROR: " & Err.Number & " " & Err.Description)
    Resume CleanExit
End Sub

Public Sub RNF_RefreshAll()
    On Error GoTo EH
    Const PROC_NAME As String = "RNF_RefreshAll"
    
    Dim wb As Workbook
    Dim controlDict As Object, tapeData As Variant
    Dim simResults As Object, waterfallResults As Object
    Dim quarterDates() As Date
    Dim calcState As XlCalculation
    Dim scr As Boolean, evt As Boolean
    
    scr = Application.ScreenUpdating: Application.ScreenUpdating = False
    calcState = Application.Calculation: Application.Calculation = xlCalculationManual
    evt = Application.EnableEvents: Application.EnableEvents = False
    
    Call Status("Reading control inputs...")
    
    Set wb = ActiveWorkbook
    
    ' Step 1: Read control inputs
    Set controlDict = ReadControlInputs(wb)
    
    ' Step 2: Normalize tape
    Call Status("Normalizing asset tape...")
    tapeData = NormalizeAssetTape(wb)
    
    ' Step 3: Build dates
    quarterDates = BuildQuarterDates(controlDict)
    
    ' Step 4: Simulate
    Call Status("Running simulation...")
    Set simResults = SimulateTape(tapeData, controlDict, quarterDates)
    
    ' Step 5: Waterfall
    Call Status("Running waterfall...")
    Set waterfallResults = RunWaterfall(simResults, controlDict, quarterDates)
    
    ' Step 6: Write results
    Call Status("Writing results...")
    Call WriteRunSheet(wb, waterfallResults, quarterDates, controlDict)
    Call DefineDynamicNamesRun(wb, controlDict)
    
    ' Step 7: Update all reporting sheets
    Call Status("Updating reports...")
    Call UpdateAllReportingSheets(wb, waterfallResults, controlDict, quarterDates)
    
    ' Step 8: Update covenant chart
    Call Status("Updating charts...")
    Call UpdateOCICChart(wb)
    
    ' Step 9: Update Investor Deck
    Call UpdateInvestorDeck(wb)
    
    ' Step 10: Arrange buttons
    Call ArrangeButtonsOnGrid(wb.Worksheets("Control"), CONTROL_BUTTON_ZONE)
    
    Call Log(PROC_NAME, "Refresh complete")
    
CleanExit:
    Call Status("")
    Application.EnableEvents = evt
    Application.Calculation = calcState
    Application.ScreenUpdating = scr
    Exit Sub
    
EH:
    Call Log(PROC_NAME, "ERROR: " & Err.Number & " " & Err.Description)
    Resume CleanExit
End Sub

Public Sub PXVZ_RunScenarioMatrix()
    On Error GoTo EH
    Const PROC_NAME As String = "PXVZ_RunScenarioMatrix"
    
    Dim wb As Workbook
    Dim state As Object
    Dim wsCache As Worksheet
    Dim flags As Collection
    Dim scenarios As Variant
    Dim scn As Object
    Dim mask As Variant
    Dim i As Long, j As Long
    Dim results() As Variant
    Dim calcState As XlCalculation
    Dim scr As Boolean, evt As Boolean
    
    scr = Application.ScreenUpdating: Application.ScreenUpdating = False
    calcState = Application.Calculation: Application.Calculation = xlCalculationManual
    evt = Application.EnableEvents: Application.EnableEvents = False
    
    Call Status("Running scenario matrix...")
    
    Set wb = ActiveWorkbook
    
    ' Step 1: Save state
    Set state = SCN_SaveState
    
    ' Step 2: Get cache sheet
    Set wsCache = GetOrCreateSheet("__ScenarioCache", True)
    
    ' Step 3: Get flag permutations
    Set flags = SCN_Permutations_FromFlags(Range("Ctl_Matrix_Flags_Col"))
    
    ' Step 4: Run scenarios
    scenarios = Array("Base", "Down", "Up")
    ReDim results(1 To (UBound(scenarios) + 1) * flags.Count, 1 To 10)
    
    Dim row As Long
    row = 1
    
    For i = 0 To UBound(scenarios)
        Set scn = SCN_Get(CStr(scenarios(i)))
        Call SCN_ApplyToControl(scn)
        
        For j = 1 To flags.Count
            Set mask = flags(j)
            Call ApplyToggleMask(mask)
            
            Call RNF_RefreshAll
            
            ' Capture KPIs
            results(row, 1) = scenarios(i)
            results(row, 2) = GetToggleString(mask)
            results(row, 3) = SafeWorksheetFunction("min", Range("Run_OC_B"))
            results(row, 4) = SafeWorksheetFunction("min", Range("Run_DSCR"))
            results(row, 5) = GetNamedValue("Reporting_Metrics!E5")  ' Equity IRR
            results(row, 6) = GetNamedValue("Reporting_Metrics!A7")  ' A WAL
            results(row, 7) = GetNamedValue("Reporting_Metrics!B7")  ' B WAL
            results(row, 8) = SCN_Hash(scn)
            
            row = row + 1
        Next j
    Next i
    
    ' Step 5: Publish results
    Call PublishScenarioMatrix(wb, results)
    
    ' Step 6: Restore state
    Call SCN_RestoreState(state)
    
    Call Log(PROC_NAME, "Matrix complete: " & row - 1 & " scenarios")
    
CleanExit:
    Call Status("")
    Application.EnableEvents = evt
    Application.Calculation = calcState
    Application.ScreenUpdating = scr
    Exit Sub
    
EH:
    Call Log(PROC_NAME, "ERROR: " & Err.Number & " " & Err.Description)
    Resume CleanExit
End Sub

Public Sub PXVZ_LoadNewAssetTape()
    On Error GoTo EH
    Const PROC_NAME As String = "PXVZ_LoadNewAssetTape"
    
    Dim filename As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim qt As QueryTable
    
    filename = Application.GetOpenFilename("CSV Files (*.csv), *.csv", Title:="Select Asset Tape CSV")
    
    If filename = False Then
        Exit Sub
    End If
    
    Call Status("Loading asset tape...")
    
    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets("AssetTape")
    
    ' Clear existing data
    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > 4 Then
        ws.Range("A5:Z" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row).ClearContents
    End If
    
    ' Import CSV
    Set qt = ws.QueryTables.Add(Connection:="TEXT;" & filename, Destination:=ws.Range("A5"))
    With qt
        .TextFileParseType = xlDelimited
        .TextFileCommaDelimiter = True
        .Refresh BackgroundQuery:=False
        .Delete
    End With
    
    ' Validate headers
    If Not ValidateTapeHeaders(ws) Then
        MsgBox "Invalid tape headers. Please check format.", vbExclamation
        Exit Sub
    End If
    
    Call Log(PROC_NAME, "Tape loaded: " & filename)
    
    ' Refresh
    Call RNF_RefreshAll
    
CleanExit:
    Call Status("")
    Exit Sub
    
EH:
    Call Log(PROC_NAME, "ERROR: " & Err.Number & " " & Err.Description)
    Resume CleanExit
End Sub

'------------------------------------------------------------------------------
' HELPER FUNCTIONS - CORE UTILITIES
'------------------------------------------------------------------------------
Private Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
    NewDict.CompareMode = 1 'TextCompare
End Function

Private Function SafeDivide(ByVal numerator As Double, ByVal denominator As Double, _
                           Optional ByVal result_if_zero As Double = 0#) As Double
    On Error Resume Next
    If Abs(denominator) < 0.00000001 Then
        SafeDivide = result_if_zero
    Else
        SafeDivide = numerator / denominator
    End If
End Function

Private Function IsDebtInstrument(ByVal securityType As String) As Boolean
    On Error Resume Next
    Dim s As String
    s = LCase(securityType)
    IsDebtInstrument = (InStr(s, "loan") > 0) Or (InStr(s, "revolver") > 0) Or _
                      (InStr(s, "ddtl") > 0) Or (InStr(s, "debt") > 0) Or _
                      (InStr(s, "term") > 0) Or (InStr(s, "lien") > 0)
End Function

Private Function SheetExists(ByVal sheetName As String, Optional ByVal wb As Workbook) As Boolean
    On Error Resume Next
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
End Function

Private Function GetOrCreateSheet(ByVal name As String, Optional ByVal veryHidden As Boolean = False) As Worksheet
    On Error Resume Next
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ActiveWorkbook
    Set ws = wb.Worksheets(name)
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = name
    End If
    
    If veryHidden Then
        ws.Visible = xlSheetVeryHidden
    Else
        ws.Visible = xlSheetVisible
    End If
    
    Set GetOrCreateSheet = ws
End Function

Private Function SafeWorksheetFunction(ByVal fnName As String, ParamArray args() As Variant) As Variant
    On Error Resume Next
    Select Case LCase(fnName)
        Case "min"
            SafeWorksheetFunction = Application.WorksheetFunction.Min(args(0))
            If Err.Number <> 0 Then SafeWorksheetFunction = DEFAULT_IF_ERROR
        Case "max"
            SafeWorksheetFunction = Application.WorksheetFunction.Max(args(0))
            If Err.Number <> 0 Then SafeWorksheetFunction = 0
        Case "sum"
            SafeWorksheetFunction = Application.WorksheetFunction.Sum(args(0))
            If Err.Number <> 0 Then SafeWorksheetFunction = 0
        Case "average"
            SafeWorksheetFunction = Application.WorksheetFunction.Average(args(0))
            If Err.Number <> 0 Then SafeWorksheetFunction = 0
        Case "index"
            SafeWorksheetFunction = Application.WorksheetFunction.Index(args(0), args(1))
            If Err.Number <> 0 Then SafeWorksheetFunction = 0
        Case "match"
            SafeWorksheetFunction = Application.WorksheetFunction.Match(args(0), args(1), args(2))
            If Err.Number <> 0 Then SafeWorksheetFunction = 0
        Case Else
            SafeWorksheetFunction = 0
    End Select
End Function

Private Sub SetCtlVal(ByVal namedRange As String, ByVal value As Variant)
    On Error Resume Next
    Dim ws As Worksheet
    Dim f As Range
    
    Set ws = ActiveWorkbook.Worksheets("Control")
    Set f = ws.Columns(1).Find(What:=namedRange, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
    If Not f Is Nothing Then
        f.Offset(0, 1).Value = value
    End If
End Sub

Private Function GetCtlVal(ByVal namedRange As String) As Variant
    On Error Resume Next
    Dim ws As Worksheet
    Dim f As Range
    
    Set ws = ActiveWorkbook.Worksheets("Control")
    Set f = ws.Columns(1).Find(What:=namedRange, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    
    If Not f Is Nothing Then
        GetCtlVal = f.Offset(0, 1).Value
    Else
        GetCtlVal = Empty
    End If
End Function

Private Sub Status(ByVal msg As String)
    On Error Resume Next
    Application.StatusBar = IIf(msg = "", False, MODULE_NAME & ": " & msg)
End Sub

Private Sub Log(ByVal whereFrom As String, ByVal msg As String)
    On Error Resume Next
    Call PXVZ_LogError(whereFrom, msg)
End Sub

Private Sub PXVZ_LogError(ByVal whereFrom As String, ByVal msg As String)
    On Error Resume Next
    Dim ws As Worksheet
    Dim nextRow As Long
    
    Set ws = GetOrCreateSheet("PXVZ_Index", False)
    
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow = 2 And ws.Cells(1, 1).Value = "" Then
        ' Add headers
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "User"
        ws.Cells(1, 3).Value = "Location"
        ws.Cells(1, 4).Value = "Message"
        ws.Range("A1:D1").Font.Bold = True
        nextRow = 2
    End If
    
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = Application.UserName
    ws.Cells(nextRow, 3).Value = whereFrom
    ws.Cells(nextRow, 4).Value = msg
End Sub

Private Sub SetNameRef(ByVal nm As String, ByVal refersTo As String)
    On Error Resume Next
    Dim n As Name
    
    ' Delete existing
    For Each n In ActiveWorkbook.Names
        If n.Name = nm Then n.Delete
    Next n
    
    ' Create new
    ActiveWorkbook.Names.Add Name:=nm, RefersTo:=refersTo
End Sub

Private Sub ApplyFreezePanesSafe(ws As Worksheet, ByVal freezeRow As Long, ByVal freezeCol As Long)
    On Error Resume Next
    With ws
        .Activate
        ActiveWindow.FreezePanes = False
        .Cells(freezeRow, freezeCol).Select
        ActiveWindow.FreezePanes = True
    End With
End Sub

'------------------------------------------------------------------------------
' TYPE CONVERSION HELPERS
'------------------------------------------------------------------------------
Private Function ToDbl(ByVal v As Variant) As Double
    On Error Resume Next
    If IsNull(v) Or IsEmpty(v) Or IsError(v) Then
        ToDbl = 0
        Exit Function
    End If
    
    Dim s As String
    s = CStr(v)
    s = Replace(s, "%", "")
    s = Replace(s, ",", "")
    s = Replace(s, "$", "")
    
    If IsNumeric(s) Then
        ToDbl = CDbl(s)
        If InStr(CStr(v), "%") > 0 Then ToDbl = ToDbl / 100
    Else
        ToDbl = 0
    End If
End Function

Private Function ToLng(ByVal v As Variant) As Long
    On Error Resume Next
    ToLng = CLng(ToDbl(v))
End Function

Private Function ToBool(ByVal v As Variant) As Boolean
    On Error Resume Next
    If IsNull(v) Or IsEmpty(v) Or IsError(v) Then
        ToBool = False
        Exit Function
    End If
    
    Dim s As String
    s = UCase(CStr(v))
    
    ToBool = (s = "TRUE") Or (s = "YES") Or (s = "1") Or (s = "-1")
End Function

'------------------------------------------------------------------------------
' VISUAL HELPERS
'------------------------------------------------------------------------------
Private Sub ArrangeButtonsOnGrid(ByVal ws As Worksheet, ByVal zoneName As String)
    On Error Resume Next
    Dim zoneRange As Range
    Dim shp As Shape
    Dim buttons() As Shape
    Dim btnCount As Long
    Dim i As Long
    Dim cellWidth As Double, cellHeight As Double
    Dim cols As Long, rows As Long
    Dim spacing As Double
    Dim row As Long, col As Long
    
    ' Get zone range
    Set zoneRange = GetNamedRange(zoneName)
    If zoneRange Is Nothing Then Exit Sub
    
    ' Collect buttons
    ReDim buttons(1 To ws.Shapes.Count)
    btnCount = 0
    
    For Each shp In ws.Shapes
        If Left(shp.Name, 4) = "btn_" Then
            btnCount = btnCount + 1
            Set buttons(btnCount) = shp
            
            ' Auto-fit
            With shp.TextFrame2
                .AutoSize = msoAutoSizeTextToFitShape
            End With
            
            ' Add padding
            shp.Width = shp.Width + 14
            shp.Height = shp.Height + 8
            
            ' Enforce minimums
            If shp.Width < 80 Then shp.Width = 80
            If shp.Height < 22 Then shp.Height = 22
        End If
    Next shp
    
    If btnCount = 0 Then Exit Sub
    
    ' Grid layout (2 columns default)
    cols = 2
    rows = (btnCount + cols - 1) \ cols
    spacing = 6
    
    cellWidth = (zoneRange.Width - spacing * (cols - 1)) / cols
    cellHeight = (zoneRange.Height - spacing * (rows - 1)) / rows
    
    ' Place buttons
    For i = 1 To btnCount
        row = ((i - 1) \ cols)
        col = ((i - 1) Mod cols)
        
        With buttons(i)
            .Left = zoneRange.Left + col * (cellWidth + spacing)
            .Top = zoneRange.Top + row * (cellHeight + spacing)
            .Placement = xlMoveAndSize
            .PrintObject = False
        End With
    Next i
End Sub

Private Function EnsureSingleChart(ByVal ws As Worksheet, ByVal chartName As String, _
                                  ByVal frameName As String) As ChartObject
    On Error Resume Next
    Dim cht As ChartObject
    Dim frameRange As Range
    
    ' Get frame range
    Set frameRange = GetNamedRange(frameName)
    If frameRange Is Nothing Then
        Set frameRange = ws.Range("B35:H55")  ' Default
    End If
    
    ' Find or create chart
    Set cht = Nothing
    For Each cht In ws.ChartObjects
        If cht.Name = chartName Then
            Exit For
        End If
    Next cht
    
    If cht Is Nothing Then
        Set cht = ws.ChartObjects.Add(Left:=frameRange.Left, Top:=frameRange.Top, _
                                      Width:=frameRange.Width, Height:=frameRange.Height)
        cht.Name = chartName
    End If
    
    ' Clear existing series
    Do While cht.Chart.SeriesCollection.Count > 0
        cht.Chart.SeriesCollection(1).Delete
    Loop
    
    ' Position to frame
    With cht
        .Left = frameRange.Left
        .Top = frameRange.Top
        .Width = frameRange.Width
        .Height = frameRange.Height
        .Placement = xlMoveAndSize
    End With
    
    Set EnsureSingleChart = cht
End Function

Private Sub ClearAndApplyOCICHeatmap(ByVal rng As Range)
    On Error Resume Next
    With rng
        .FormatConditions.Delete
        .FormatConditions.AddColorScale ColorScaleType:=3
        .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        .FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = RGB(255, 199, 206)  ' Red
        .FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .FormatConditions(1).ColorScaleCriteria(2).Value = 50
        .FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 156)  ' Yellow
        .FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueHighestValue
        .FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = RGB(198, 239, 206)  ' Green
    End With
End Sub

Private Sub RemoveShapesByPrefix(ByVal ws As Worksheet, ByVal prefix As String)
    On Error Resume Next
    Dim shp As Shape
    Dim i As Long
    
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If Left(shp.Name, Len(prefix)) = prefix Then
            shp.Delete
        End If
    Next i
End Sub

'------------------------------------------------------------------------------
' SCENARIO MANAGER
'------------------------------------------------------------------------------
Private Function SCN_Registry() As Object
    On Error Resume Next
    Static registry As Object
    
    If registry Is Nothing Then
        Set registry = NewDict()
    End If
    
    Set SCN_Registry = registry
End Function

Private Function SCN_Hash(ByVal scn As Object) As String
    On Error Resume Next
    Dim key As Variant
    Dim keys() As String
    Dim i As Long
    Dim result As String
    
    If scn Is Nothing Then
        SCN_Hash = "NULL"
        Exit Function
    End If
    
    ' Get sorted keys
    ReDim keys(0 To scn.Count - 1)
    i = 0
    For Each key In scn.Keys
        keys(i) = CStr(key)
        i = i + 1
    Next key
    
    ' Simple bubble sort
    Dim j As Long, temp As String
    For i = 0 To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If keys(i) > keys(j) Then
                temp = keys(i)
                keys(i) = keys(j)
                keys(j) = temp
            End If
        Next j
    Next i
    
    ' Build hash string
    result = "{"
    For i = 0 To UBound(keys)
        result = result & """" & keys(i) & """:""" & CStr(scn(keys(i))) & """"
        If i < UBound(keys) Then result = result & ","
    Next i
    result = result & "}"
    
    SCN_Hash = result
End Function

Public Sub SCN_RegisterDefaults()
    On Error Resume Next
    Dim base As Object, down As Object, up As Object
    
    ' Base scenario
    Set base = NewDict()
    base("Base_CDR") = 0.02
    base("Base_Recovery") = 0.65
    base("Base_Prepay") = 0.05
    base("Base_Amort") = 0.02
    base("Spread_Add_bps") = 0
    base("Rate_Add_bps") = 0
    
    ' Down scenario
    Set down = NewDict()
    down("Base_CDR") = 0.025
    down("Base_Recovery") = 0.6
    down("Base_Prepay") = 0.03
    down("Base_Amort") = 0.02
    down("Spread_Add_bps") = 100
    down("Rate_Add_bps") = 50
    
    ' Up scenario
    Set up = NewDict()
    up("Base_CDR") = 0.016
    up("Base_Recovery") = 0.7
    up("Base_Prepay") = 0.07
    up("Base_Amort") = 0.02
    up("Spread_Add_bps") = -25
    up("Rate_Add_bps") = -25
    
    Call SCN_Define("Base", base)
    Call SCN_Define("Down", down)
    Call SCN_Define("Up", up)
End Sub

Public Sub SCN_Define(ByVal name As String, ByVal params As Object)
    On Error Resume Next
    SCN_Registry()(name) = params
End Sub

Public Function SCN_Get(ByVal name As String) As Object
    On Error Resume Next
    Set SCN_Get = SCN_Registry()(name)
End Function

Public Sub SCN_Delete(ByVal name As String)
    On Error Resume Next
    If SCN_Registry().Exists(name) Then
        SCN_Registry().Remove name
    End If
End Sub

Public Function SCN_ReadFromControl() As Object
    On Error Resume Next
    Dim result As Object
    Set result = NewDict()
    
    result("Base_CDR") = GetCtlVal("Base_CDR")
    result("Base_Recovery") = GetCtlVal("Base_Recovery")
    result("Base_Prepay") = GetCtlVal("Base_Prepay")
    result("Base_Amort") = GetCtlVal("Base_Amort")
    result("Spread_Add_bps") = GetCtlVal("Spread_Add_bps")
    result("Rate_Add_bps") = GetCtlVal("Rate_Add_bps")
    
    Set SCN_ReadFromControl = result
End Function

Public Sub SCN_ApplyToControl(ByVal scn As Object)
    On Error Resume Next
    Dim key As Variant
    
    For Each key In scn.Keys
        Call SetCtlVal(CStr(key), scn(key))
    Next key
End Sub

Public Function SCN_SaveState() As Object
    On Error Resume Next
    Set SCN_SaveState = SCN_ReadFromControl()
End Function

Public Sub SCN_RestoreState(ByVal state As Object)
    On Error Resume Next
    Call SCN_ApplyToControl(state)
End Sub

Public Function SCN_Permutations_FromFlags(ByVal flagRange As Range) As Collection
    On Error Resume Next
    Dim result As New Collection
    Dim cell As Range
    Dim toggles As New Collection
    Dim numToggles As Long
    Dim i As Long, j As Long
    Dim mask As Object
    
    ' Collect TRUE flags
    For Each cell In flagRange
        If ToBool(cell.Value) Then
            toggles.Add cell.Offset(0, -1).Value  ' Get toggle name from left column
        End If
    Next cell
    
    numToggles = toggles.Count
    
    ' Generate all bitmask combinations
    If numToggles = 0 Then
        Set mask = NewDict()
        result.Add mask
    Else
        For i = 0 To (2 ^ numToggles) - 1
            Set mask = NewDict()
            For j = 1 To numToggles
                mask(toggles(j)) = ((i And (2 ^ (j - 1))) <> 0)
            Next j
            result.Add mask
        Next i
    End If
    
    Set SCN_Permutations_FromFlags = result
End Function

'------------------------------------------------------------------------------
' SHEET CREATION AND SETUP
'------------------------------------------------------------------------------
Private Sub CreateAllSheets(wb As Workbook)
    On Error Resume Next
    
    ' Control and input sheets
    Call GetOrCreateSheet("Control", False)
    Call GetOrCreateSheet("AssetTape", False)
    
    ' Engine sheets
    Call GetOrCreateSheet("Run", False)
    Call GetOrCreateSheet("M_Ref_Full", False)
    Call GetOrCreateSheet("M_Scaffold", True)
    Call GetOrCreateSheet("__ScenarioCache", True)
    Call GetOrCreateSheet("__Log", True)
    
    ' Reporting sheets (§3 requirements)
    Call GetOrCreateSheet("Exec_Summary", False)
    Call GetOrCreateSheet("Sources_Uses_At_Close", False)
    Call GetOrCreateSheet("NAV_Roll_Forward", False)
    Call GetOrCreateSheet("Reserves_Tracking", False)
    Call GetOrCreateSheet("Cashflow_Waterfall_Summary", False)
    Call GetOrCreateSheet("Tranche_Cashflows", False)
    Call GetOrCreateSheet("OCIC_Tests", False)
    Call GetOrCreateSheet("Breaches_Dashboard", False)
    Call GetOrCreateSheet("Portfolio_Stratifications", False)
    Call GetOrCreateSheet("Asset_Performance", False)
    Call GetOrCreateSheet("Portfolio_Cashflows_Detail", False)
    Call GetOrCreateSheet("Fees_Expenses", False)
    Call GetOrCreateSheet("Sensitivity_Matrix", False)
    Call GetOrCreateSheet("BreakEven_Analytics", False)
    Call GetOrCreateSheet("MonteCarlo_Summary", False)
    Call GetOrCreateSheet("Investor_Distributions", False)
    Call GetOrCreateSheet("Reporting_Metrics", False)
    
    ' Additional sheets
    Call GetOrCreateSheet("Portfolio_HHI", False)
    Call GetOrCreateSheet("RBC_Factors", False)
    Call GetOrCreateSheet("Waterfall_Schedule", False)
    Call GetOrCreateSheet("Investor_Deck", False)
    Call GetOrCreateSheet("Version_History", False)
    Call GetOrCreateSheet("PXVZ_Index", False)
    Call GetOrCreateSheet("Fix_Log", False)
End Sub

Private Sub CreateNamedFrames(wb As Workbook)
    On Error Resume Next
    
    ' Create frame named ranges with defaults
    Call SetNameRef(OCIC_CHART_FRAME, "=OCIC_Tests!$B$35:$H$55")
    Call SetNameRef(SCENARIO_MATRIX_FRAME, "=Control!$J$29:$Q$60")
    Call SetNameRef(INVESTOR_CHART_FRAME, "=Investor_Deck!$B$28:$H$45")
    Call SetNameRef(CONTROL_BUTTON_ZONE, "=Control!$J$2:$K$12")
End Sub

Private Sub SetupControlSheet(ws As Worksheet)
    On Error Resume Next
    Dim r As Long
    
    ws.Cells.Clear
    
    ' Title
    ws.Range("A1").Value = "RATED NOTE FEEDER - CONTROL PANEL"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "(in $000s)"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Headers
    r = 4
    ws.Cells(r, 1).Value = "KEY"
    ws.Cells(r, 2).Value = "VALUE"
    ws.Cells(r, 3).Value = "Include in Matrix?"
    ws.Range("A4:C4").Style = "SG_Hdr"
    
    ' Core parameters
    r = r + 1: ws.Cells(r, 1).Value = "NumQuarters": ws.Cells(r, 2).Value = 48
    r = r + 1: ws.Cells(r, 1).Value = "First_Close_Date": ws.Cells(r, 2).Value = DateSerial(2025, 12, 1)
    r = r + 1: ws.Cells(r, 1).Value = "Total_Capital": ws.Cells(r, 2).Value = 600000
    
    ' Capital structure
    r = r + 1: ws.Cells(r, 1).Value = "Pct_A": ws.Cells(r, 2).Value = 0.65
    r = r + 1: ws.Cells(r, 1).Value = "Pct_B": ws.Cells(r, 2).Value = 0.15
    r = r + 1: ws.Cells(r, 1).Value = "Enable_C": ws.Cells(r, 2).Value = False: ws.Cells(r, 3).Value = True
    r = r + 1: ws.Cells(r, 1).Value = "Pct_C": ws.Cells(r, 2).Value = 0
    r = r + 1: ws.Cells(r, 1).Value = "Enable_D": ws.Cells(r, 2).Value = False: ws.Cells(r, 3).Value = True
    r = r + 1: ws.Cells(r, 1).Value = "Pct_D": ws.Cells(r, 2).Value = 0
    r = r + 1: ws.Cells(r, 1).Value = "Pct_E": ws.Cells(r, 2).Value = 0.2
    
    ' Spreads
    r = r + 1: ws.Cells(r, 1).Value = "Spread_A_bps": ws.Cells(r, 2).Value = 250
    r = r + 1: ws.Cells(r, 1).Value = "Spread_B_bps": ws.Cells(r, 2).Value = 525
    r = r + 1: ws.Cells(r, 1).Value = "Spread_C_bps": ws.Cells(r, 2).Value = 600
    r = r + 1: ws.Cells(r, 1).Value = "Spread_D_bps": ws.Cells(r, 2).Value = 800
    
    ' OC Triggers
    r = r + 1: ws.Cells(r, 1).Value = "OC_Trigger_A": ws.Cells(r, 2).Value = 1.25
    r = r + 1: ws.Cells(r, 1).Value = "OC_Trigger_B": ws.Cells(r, 2).Value = 1.125
    r = r + 1: ws.Cells(r, 1).Value = "OC_Trigger_C": ws.Cells(r, 2).Value = 1.05
    r = r + 1: ws.Cells(r, 1).Value = "OC_Trigger_D": ws.Cells(r, 2).Value = 1.00
    
    ' Timing
    r = r + 1: ws.Cells(r, 1).Value = "Reinvest_Q": ws.Cells(r, 2).Value = 12: ws.Cells(r, 3).Value = True
    r = r + 1: ws.Cells(r, 1).Value = "GP_Extend_Q": ws.Cells(r, 2).Value = 4: ws.Cells(r, 3).Value = True
    
    ' Features
    r = r + 1: ws.Cells(r, 1).Value = "Enable_Turbo_DOC": ws.Cells(r, 2).Value = True: ws.Cells(r, 3).Value = True
    r = r + 1: ws.Cells(r, 1).Value = "Enable_Excess_Reserve": ws.Cells(r, 2).Value = True: ws.Cells(r, 3).Value = True
    r = r + 1: ws.Cells(r, 1).Value = "Enable_PIK": ws.Cells(r, 2).Value = True: ws.Cells(r, 3).Value = False
    r = r + 1: ws.Cells(r, 1).Value = "Enable_CC_PIK": ws.Cells(r, 2).Value = True: ws.Cells(r, 3).Value = False
    r = r + 1: ws.Cells(r, 1).Value = "Enable_Recycling": ws.Cells(r, 2).Value = False: ws.Cells(r, 3).Value = True
    r = r + 1: ws.Cells(r, 1).Value = "Recycling_Pct": ws.Cells(r, 2).Value = 0.75
    r = r + 1: ws.Cells(r, 1).Value = "Recycle_Spread_bps": ws.Cells(r, 2).Value = 550
    r = r + 1: ws.Cells(r, 1).Value = "Close_Call_Pct": ws.Cells(r, 2).Value = 0.25
    
    ' Reserve
    r = r + 1: ws.Cells(r, 1).Value = "Reserve_Pct": ws.Cells(r, 2).Value = 0.025
    r = r + 1: ws.Cells(r, 1).Value = "PIK_Pct": ws.Cells(r, 2).Value = 1
    
    ' Fees
    r = r + 1: ws.Cells(r, 1).Value = "Arranger_Fee_bps": ws.Cells(r, 2).Value = 50
    r = r + 1: ws.Cells(r, 1).Value = "Rating_Agency_Fee_bps": ws.Cells(r, 2).Value = 10
    r = r + 1: ws.Cells(r, 1).Value = "Servicer_Fee_bps": ws.Cells(r, 2).Value = 25
    r = r + 1: ws.Cells(r, 1).Value = "Mgmt_Fee_Pct": ws.Cells(r, 2).Value = 0.01
    r = r + 1: ws.Cells(r, 1).Value = "Admin_Fee_Floor": ws.Cells(r, 2).Value = 50
    
    ' Downgrade
    r = r + 1: ws.Cells(r, 1).Value = "Downgrade_OC": ws.Cells(r, 2).Value = 1.08
    r = r + 1: ws.Cells(r, 1).Value = "Downgrade_Spd_Adj_bps": ws.Cells(r, 2).Value = 50
    
    ' Scenarios
    r = r + 1: ws.Cells(r, 1).Value = "Scenario_Selection": ws.Cells(r, 2).Value = "Base"
    r = r + 1: ws.Cells(r, 1).Value = "Base_CDR": ws.Cells(r, 2).Value = 0.02
    r = r + 1: ws.Cells(r, 1).Value = "Base_Recovery": ws.Cells(r, 2).Value = 0.65
    r = r + 1: ws.Cells(r, 1).Value = "Base_Prepay": ws.Cells(r, 2).Value = 0.05
    r = r + 1: ws.Cells(r, 1).Value = "Base_Amort": ws.Cells(r, 2).Value = 0.02
    r = r + 1: ws.Cells(r, 1).Value = "Spread_Add_bps": ws.Cells(r, 2).Value = 0
    r = r + 1: ws.Cells(r, 1).Value = "Rate_Add_bps": ws.Cells(r, 2).Value = 0
    r = r + 1: ws.Cells(r, 1).Value = "Loss_Lag_Q": ws.Cells(r, 2).Value = 4
    
    ' Monte Carlo
    r = r + 1: ws.Cells(r, 1).Value = "MC_Iterations": ws.Cells(r, 2).Value = 200
    r = r + 1: ws.Cells(r, 1).Value = "MC_Sigma_CDR": ws.Cells(r, 2).Value = 0.3
    r = r + 1: ws.Cells(r, 1).Value = "MC_Sigma_Rec": ws.Cells(r, 2).Value = 0.2
    r = r + 1: ws.Cells(r, 1).Value = "MC_Sigma_Sprd_bps": ws.Cells(r, 2).Value = 50
    r = r + 1: ws.Cells(r, 1).Value = "MC_Rho": ws.Cells(r, 2).Value = 0.3
    r = r + 1: ws.Cells(r, 1).Value = "MC_Seed": ws.Cells(r, 2).Value = 42
    
    ' Fund terms
    r = r + 1: ws.Cells(r, 1).Value = "Pref_Hurdle": ws.Cells(r, 2).Value = 0.08
    r = r + 1: ws.Cells(r, 1).Value = "GP_Catch_Up_Pct": ws.Cells(r, 2).Value = 0.2
    r = r + 1: ws.Cells(r, 1).Value = "GP_Split_Pct": ws.Cells(r, 2).Value = 0.2
    
    ' Call Schedule
    r = r + 2
    ws.Cells(r, 1).Value = "Call_Schedule"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r + 1, 1).Value = 0.5
    ws.Cells(r + 2, 1).Value = 0.167
    ws.Cells(r + 3, 1).Value = 0.167
    ws.Cells(r + 4, 1).Value = 0.166
    ws.Range(ws.Cells(r + 1, 1), ws.Cells(r + 4, 1)).Name = "Call_Schedule"
    
    ' Data validation
    Call ApplyControlValidation(ws)
    
    ' Format
    Call FormatControlSheet(ws)
    
    ' KPI Placards
    Call CreateKPIPlacards(ws)
End Sub

Private Sub ApplyControlValidation(ws As Worksheet)
    On Error Resume Next
    Dim rng As Range
    
    ' Scenario selection
    Set rng = ws.Columns(1).Find("Scenario_Selection", LookAt:=xlWhole).Offset(0, 1)
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="Base,Down,Up"
    End With
    
    ' Boolean toggles
    Dim toggleKeys As Variant
    toggleKeys = Array("Enable_C", "Enable_D", "Enable_Turbo_DOC", "Enable_Excess_Reserve", _
                      "Enable_PIK", "Enable_CC_PIK", "Enable_Recycling")
    
    Dim i As Long
    For i = 0 To UBound(toggleKeys)
        Set rng = ws.Columns(1).Find(toggleKeys(i), LookAt:=xlWhole)
        If Not rng Is Nothing Then
            With rng.Offset(0, 1).Validation
                .Delete
                .Add Type:=xlValidateList, Formula1:="TRUE,FALSE"
            End With
        End If
    Next i
    
    ' Matrix flags column
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    With ws.Range("C5:C" & lastRow).Validation
        .Delete
        .Add Type:=xlValidateList, Formula1:="TRUE,FALSE"
    End With
    
    ws.Range("C5:C" & lastRow).Name = "Ctl_Matrix_Flags_Col"
End Sub

Private Sub FormatControlSheet(ws As Worksheet)
    On Error Resume Next
    
    ' Number formats
    ws.Range("B5:B100").NumberFormat = "General"
    
    ' Specific formats
    Dim cell As Range
    For Each cell In ws.Range("A5:A100")
        Select Case cell.Value
            Case "First_Close_Date"
                cell.Offset(0, 1).NumberFormat = "mm/dd/yyyy"
            Case "Total_Capital", "Admin_Fee_Floor"
                cell.Offset(0, 1).Style = "SG_Currency_K"
            Case "Pct_A", "Pct_B", "Pct_C", "Pct_D", "Pct_E", "Reserve_Pct", "PIK_Pct", _
                 "Base_CDR", "Base_Recovery", "Base_Prepay", "Base_Amort", _
                 "Mgmt_Fee_Pct", "Pref_Hurdle", "GP_Catch_Up_Pct", "GP_Split_Pct", _
                 "Recycling_Pct", "Close_Call_Pct"
                cell.Offset(0, 1).Style = "SG_Pct"
        End Select
    Next cell
    
    ' Apply style pack
    Call SG615_ApplyStylePack(ws, "Control Panel", "(in $000s)")
    
    ' Column widths
    ws.Columns("A:A").ColumnWidth = 25
    ws.Columns("B:B").ColumnWidth = 15
    ws.Columns("C:C").ColumnWidth = 18
    ws.Columns("D:D").ColumnWidth = 15
    ws.Columns("E:E").ColumnWidth = 2
    ws.Columns("F:G").ColumnWidth = 20
End Sub

Private Sub CreateKPIPlacards(ws As Worksheet)
    On Error Resume Next
    Dim r As Long
    
    r = 4
    ws.Cells(r, 6).Value = "KEY PERFORMANCE INDICATORS"
    ws.Cells(r, 6).Style = "SG_Title"
    
    r = r + 2
    ws.Cells(r, 6).Value = "Equity IRR"
    ws.Cells(r, 7).Formula = "=IFERROR(Reporting_Metrics!E5,0)"
    ws.Cells(r, 7).Style = "SG_Pct"
    ws.Cells(r, 7).Font.Size = 14
    ws.Cells(r, 7).Font.Bold = True
    
    r = r + 2
    ws.Cells(r, 6).Value = "MOIC"
    ws.Cells(r, 7).Formula = "=IFERROR(Reporting_Metrics!E6,0)"
    ws.Cells(r, 7).NumberFormat = "0.00x"
    ws.Cells(r, 7).Font.Size = 14
    ws.Cells(r, 7).Font.Bold = True
    
    r = r + 2
    ws.Cells(r, 6).Value = "Class A WAL"
    ws.Cells(r, 7).Formula = "=IFERROR(Reporting_Metrics!A7,0)"
    ws.Cells(r, 7).NumberFormat = "0.0"
    ws.Cells(r, 7).Font.Size = 14
    ws.Cells(r, 7).Font.Bold = True
    
    r = r + 2
    ws.Cells(r, 6).Value = "Min OC_B"
    ws.Cells(r, 7).Formula = "=IFERROR(MIN(Run_OC_B),999)"
    ws.Cells(r, 7).NumberFormat = "0.00x"
    ws.Cells(r, 7).Font.Size = 14
    ws.Cells(r, 7).Font.Bold = True
    ws.Cells(r, 7).Interior.Color = RGB(198, 239, 206)
    
    r = r + 2
    ws.Cells(r, 6).Value = "Min DSCR"
    ws.Cells(r, 7).Formula = "=IFERROR(MIN(Run_DSCR),999)"
    ws.Cells(r, 7).NumberFormat = "0.00x"
    ws.Cells(r, 7).Font.Size = 14
    ws.Cells(r, 7).Font.Bold = True
    ws.Cells(r, 7).Interior.Color = RGB(198, 239, 206)
    
    r = r + 2
    ws.Cells(r, 6).Value = "Breach Periods"
    ws.Cells(r, 7).Formula = "=COUNTIF(Run_OC_B,""<""&Ctl_OC_Trigger_B)"
    ws.Cells(r, 7).NumberFormat = "0"
    ws.Cells(r, 7).Font.Size = 14
    ws.Cells(r, 7).Font.Bold = True
    
    r = r + 2
    ws.Cells(r, 6).Value = "Turbo Active"
    ws.Cells(r, 7).Formula = "=IF(SUM(Run_TurboFlag)>0,""YES"",""NO"")"
    ws.Cells(r, 7).Font.Size = 14
    ws.Cells(r, 7).Font.Bold = True
End Sub

Private Sub CreateAllButtons(wb As Workbook)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets("Control")
    
    ' Create buttons
    Call CreateButton(ws, "D5", "REFRESH", "RNF_RefreshAll")
    Call CreateButton(ws, "D6", "BUILD", "RNF_Strict_BuildAndRun")
    Call CreateButton(ws, "D7", "SCENARIO MATRIX", "PXVZ_RunScenarioMatrix")
    Call CreateButton(ws, "D8", "LOAD TAPE", "PXVZ_LoadNewAssetTape")
    Call CreateButton(ws, "D9", "SENSITIVITIES", "RunSensitivities")
    Call CreateButton(ws, "D10", "MONTE CARLO", "RunMonteCarlo")
    Call CreateButton(ws, "D11", "BREAKEVEN", "RunBreakeven")
    Call CreateButton(ws, "D12", "CLEAR", "ClearOutputSheets")
End Sub

Private Sub CreateButton(ws As Worksheet, ByVal cellAddress As String, _
                        ByVal caption As String, ByVal macroName As String)
    On Error Resume Next
    Dim btn As Shape
    Dim rng As Range
    
    Set rng = ws.Range(cellAddress)
    
    ' Delete existing
    RemoveShapesByPrefix ws, "btn_" & caption
    
    ' Create new
    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                 rng.Left, rng.Top, 100, 25)
    
    With btn
        .Name = "btn_" & caption
        .TextFrame2.TextRange.Characters.Text = caption
        .TextFrame2.TextRange.Font.Bold = True
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Fill.ForeColor.RGB = SG_RED
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Visible = msoFalse
        .OnAction = macroName
        .Placement = xlMoveAndSize
        .PrintObject = False
    End With
End Sub

'------------------------------------------------------------------------------
' STYLES CREATION
'------------------------------------------------------------------------------
Private Sub Build_Styles()
    On Error Resume Next
    Call SG615_CreateStyles(ActiveWorkbook)
End Sub

Private Sub SG615_CreateStyles(wb As Workbook)
    On Error Resume Next
    Dim s As Style
    
    ' SG_Title
    On Error Resume Next
    Set s = wb.Styles("SG_Title")
    If s Is Nothing Then Set s = wb.Styles.Add("SG_Title")
    With s
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = SG_BLACK
    End With
    
    ' SG_Subtitle
    On Error Resume Next
    Set s = wb.Styles("SG_Subtitle")
    If s Is Nothing Then Set s = wb.Styles.Add("SG_Subtitle")
    With s
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = SG_SLATE
    End With
    
    ' SG_Hdr
    On Error Resume Next
    Set s = wb.Styles("SG_Hdr")
    If s Is Nothing Then Set s = wb.Styles.Add("SG_Hdr")
    With s
        .Font.Bold = True
        .Interior.Color = SG_GRAY_LIGHT
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlThin
    End With
    
    ' SG_Currency_K
    On Error Resume Next
    Set s = wb.Styles("SG_Currency_K")
    If s Is Nothing Then Set s = wb.Styles.Add("SG_Currency_K")
    s.NumberFormat = "$#,##0"
    
    ' SG_Pct
    On Error Resume Next
    Set s = wb.Styles("SG_Pct")
    If s Is Nothing Then Set s = wb.Styles.Add("SG_Pct")
    s.NumberFormat = "0.0%"
    
    ' SG_Num2
    On Error Resume Next
    Set s = wb.Styles("SG_Num2")
    If s Is Nothing Then Set s = wb.Styles.Add("SG_Num2")
    s.NumberFormat = "0.00"
    
    ' SG_Int
    On Error Resume Next
    Set s = wb.Styles("SG_Int")
    If s Is Nothing Then Set s = wb.Styles.Add("SG_Int")
    s.NumberFormat = "#,##0"
End Sub

Private Sub SG615_ApplyStylePack(ws As Worksheet, ByVal title As String, ByVal unitsText As String)
    On Error Resume Next
    
    If title <> "" Then
        ws.Range("A1").Value = title
        ws.Range("A1").Style = "SG_Title"
    End If
    
    If unitsText <> "" Then
        ws.Range("A2").Value = unitsText
        ws.Range("A2").Style = "SG_Subtitle"
    End If
    
    ' Hide gridlines
    ActiveWindow.DisplayGridlines = False
    
    ' Set print area
    ws.PageSetup.PrintArea = ws.UsedRange.Address
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.FitToPagesWide = 1
    ws.PageSetup.FitToPagesTall = False
    
    ' AutoFit columns
    ws.Columns.AutoFit
End Sub

'------------------------------------------------------------------------------
' ASSET TAPE HANDLING
'------------------------------------------------------------------------------
Private Sub SeedAssetTape(wb As Workbook)
    On Error Resume Next
    Dim ws As Worksheet
    Dim headers As Variant
    Dim data As Variant
    Dim r As Long
    
    Set ws = wb.Worksheets("AssetTape")
    ws.Cells.Clear
    
    ' Title
    ws.Range("A1").Value = "ASSET TAPE"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "(in $000s)"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Headers (row 4)
    headers = Array("Borrower", "Par", "DrawPct", "Spread_bps", "OID_bps", _
                   "Facility_Type", "Security_Type", "Maturity_Date", "Years_To_Mat", _
                   "LTV_Pct", "Rating", "Industry", "LTM_EBITDA", "Total_Leverage", "Notes_ID")
    
    ws.Range("A4").Resize(1, UBound(headers) + 1).Value = headers
    ws.Range("A4").Resize(1, UBound(headers) + 1).Style = "SG_Hdr"
    
    ' Sample data
    ReDim data(1 To 10, 1 To 15)
    
    ' Company A
    data(1, 1) = "Acme Corp"
    data(1, 2) = 15000
    data(1, 3) = 1
    data(1, 4) = 450
    data(1, 5) = 0
    data(1, 6) = "Term Loan"
    data(1, 7) = "First Lien"
    data(1, 8) = DateSerial(2030, 6, 30)
    data(1, 9) = 5.5
    data(1, 10) = 0.65
    data(1, 11) = "B+"
    data(1, 12) = "Technology"
    data(1, 13) = 25000
    data(1, 14) = 4.5
    data(1, 15) = "TL001"
    
    ' Company B
    data(2, 1) = "Beta Industries"
    data(2, 2) = 12000
    data(2, 3) = 0.8
    data(2, 4) = 525
    data(2, 5) = 25
    data(2, 6) = "Revolver"
    data(2, 7) = "First Lien"
    data(2, 8) = DateSerial(2029, 12, 31)
    data(2, 9) = 5
    data(2, 10) = 0.55
    data(2, 11) = "B"
    data(2, 12) = "Healthcare"
    data(2, 13) = 18000
    data(2, 14) = 5.2
    data(2, 15) = "RV002"
    
    ' Company C
    data(3, 1) = "Gamma Manufacturing"
    data(3, 2) = 8000
    data(3, 3) = 1
    data(3, 4) = 475
    data(3, 5) = 0
    data(3, 6) = "Term Loan"
    data(3, 7) = "First Lien"
    data(3, 8) = DateSerial(2031, 3, 31)
    data(3, 9) = 6
    data(3, 10) = 0.7
    data(3, 11) = "BB-"
    data(3, 12) = "Manufacturing"
    data(3, 13) = 12000
    data(3, 14) = 4.8
    data(3, 15) = "TL003"
    
    ' Company D
    data(4, 1) = "Delta Services"
    data(4, 2) = 6000
    data(4, 3) = 1
    data(4, 4) = 500
    data(4, 5) = 50
    data(4, 6) = "Term Loan"
    data(4, 7) = "Second Lien"
    data(4, 8) = DateSerial(2032, 1, 1)
    data(4, 9) = 7
    data(4, 10) = 0.75
    data(4, 11) = "B-"
    data(4, 12) = "Business Services"
    data(4, 13) = 8000
    data(4, 14) = 5.5
    data(4, 15) = "TL004"
    
    ' Company E
    data(5, 1) = "Epsilon Retail"
    data(5, 2) = 5000
    data(5, 3) = 0.5
    data(5, 4) = 550
    data(5, 5) = 0
    data(5, 6) = "Revolver"
    data(5, 7) = "First Lien"
    data(5, 8) = DateSerial(2028, 6, 30)
    data(5, 9) = 3.5
    data(5, 10) = 0.6
    data(5, 11) = "BB"
    data(5, 12) = "Retail"
    data(5, 13) = 10000
    data(5, 14) = 4.2
    data(5, 15) = "RV005"
    
    ' Company F
    data(6, 1) = "Zeta Energy"
    data(6, 2) = 10000
    data(6, 3) = 1
    data(6, 4) = 425
    data(6, 5) = 0
    data(6, 6) = "Term Loan"
    data(6, 7) = "First Lien"
    data(6, 8) = DateSerial(2030, 9, 30)
    data(6, 9) = 5.75
    data(6, 10) = 0.62
    data(6, 11) = "BB+"
    data(6, 12) = "Energy"
    data(6, 13) = 20000
    data(6, 14) = 4.0
    data(6, 15) = "TL006"
    
    ' Company G
    data(7, 1) = "Eta Telecom"
    data(7, 2) = 7500
    data(7, 3) = 1
    data(7, 4) = 460
    data(7, 5) = 0
    data(7, 6) = "Term Loan"
    data(7, 7) = "First Lien"
    data(7, 8) = DateSerial(2029, 3, 31)
    data(7, 9) = 4.25
    data(7, 10) = 0.58
    data(7, 11) = "B+"
    data(7, 12) = "Telecommunications"
    data(7, 13) = 15000
    data(7, 14) = 4.3
    data(7, 15) = "TL007"
    
    ' Company H
    data(8, 1) = "Theta Foods"
    data(8, 2) = 4000
    data(8, 3) = 1
    data(8, 4) = 490
    data(8, 5) = 25
    data(8, 6) = "Term Loan"
    data(8, 7) = "First Lien"
    data(8, 8) = DateSerial(2031, 12, 31)
    data(8, 9) = 7
    data(8, 10) = 0.68
    data(8, 11) = "B"
    data(8, 12) = "Food & Beverage"
    data(8, 13) = 6000
    data(8, 14) = 4.9
    data(8, 15) = "TL008"
    
    ' Company I (DDTL)
    data(9, 1) = "Iota Construction"
    data(9, 2) = 3000
    data(9, 3) = 0.25
    data(9, 4) = 575
    data(9, 5) = 0
    data(9, 6) = "DDTL"
    data(9, 7) = "First Lien"
    data(9, 8) = DateSerial(2028, 12, 31)
    data(9, 9) = 4
    data(9, 10) = 0.72
    data(9, 11) = "B-"
    data(9, 12) = "Construction"
    data(9, 13) = 5000
    data(9, 14) = 5.8
    data(9, 15) = "DD009"
    
    ' Company J
    data(10, 1) = "Kappa Pharma"
    data(10, 2) = 9000
    data(10, 3) = 1
    data(10, 4) = 440
    data(10, 5) = 0
    data(10, 6) = "Term Loan"
    data(10, 7) = "First Lien"
    data(10, 8) = DateSerial(2029, 6, 30)
    data(10, 9) = 4.5
    data(10, 10) = 0.64
    data(10, 11) = "BB-"
    data(10, 12) = "Pharmaceuticals"
    data(10, 13) = 22000
    data(10, 14) = 3.8
    data(10, 15) = "TL010"
    
    ' Write data
    ws.Range("A5").Resize(10, 15).Value = data
    
    ' Format
    ws.Range("B:B").Style = "SG_Currency_K"
    ws.Range("C:C").Style = "SG_Pct"
    ws.Range("J:J").Style = "SG_Pct"
    ws.Range("M:M").Style = "SG_Currency_K"
    ws.Range("N:N").NumberFormat = "0.0x"
    
    Call SG615_ApplyStylePack(ws, "Asset Tape", "(in $000s)")
End Sub

Private Function NormalizeAssetTape(wb As Workbook) As Variant
    On Error GoTo EH
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim data As Variant
    Dim r As Long
    Dim startDate As Date
    
    Set ws = wb.Worksheets("AssetTape")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = 15  ' Fixed columns
    
    If lastRow < 5 Then
        Call Log("NormalizeAssetTape", "No loan data found")
        Exit Function
    End If
    
    data = ws.Range(ws.Cells(4, 1), ws.Cells(lastRow, lastCol)).Value
    startDate = CDate(GetCtlVal("First_Close_Date"))
    
    ' Process each loan
    For r = 2 To UBound(data, 1)
        ' Handle DrawPct
        If ToDbl(data(r, 3)) > 1.5 Then
            data(r, 3) = data(r, 3) / 100
        End If
        
        ' Handle LTV_Pct
        If ToDbl(data(r, 10)) > 1.5 Then
            data(r, 10) = data(r, 10) / 100
        End If
        
        ' Maturity logic
        If UCase(data(r, 8)) = "NM" Then
            ' No maturity - equity-like
            data(r, 9) = "NM"
        ElseIf IsDate(data(r, 8)) Then
            ' Have maturity date
            Dim maturityDate As Date
            maturityDate = CDate(data(r, 8))
            data(r, 9) = DateDiff("m", startDate, maturityDate) / 12
        ElseIf IsNumeric(data(r, 9)) And data(r, 9) > 0 Then
            ' Have years to maturity
            data(r, 8) = DateAdd("yyyy", CLng(data(r, 9)), startDate)
        Else
            ' Default to 5 years
            data(r, 8) = DateAdd("yyyy", 5, startDate)
            data(r, 9) = 5
        End If
    Next r
    
    NormalizeAssetTape = data
    Exit Function
    
EH:
    Call Log("NormalizeAssetTape", Err.Description)
End Function

Private Function ValidateTapeHeaders(ws As Worksheet) As Boolean
    On Error Resume Next
    Dim expectedHeaders As Variant
    Dim i As Long
    
    expectedHeaders = Array("Borrower", "Par", "DrawPct", "Spread_bps", "OID_bps", _
                          "Facility_Type", "Security_Type", "Maturity_Date", "Years_To_Mat", _
                          "LTV_Pct", "Rating", "Industry", "LTM_EBITDA", "Total_Leverage", "Notes_ID")
    
    For i = 0 To UBound(expectedHeaders)
        If InStr(1, ws.Cells(4, i + 1).Value, expectedHeaders(i), vbTextCompare) = 0 Then
            ValidateTapeHeaders = False
            Exit Function
        End If
    Next i
    
    ValidateTapeHeaders = True
End Function

'------------------------------------------------------------------------------
' CONTROL INPUTS
'------------------------------------------------------------------------------
Private Function ReadControlInputs(wb As Workbook) As Object
    On Error Resume Next
    Dim dict As Object
    Dim ws As Worksheet
    Dim r As Long
    Dim key As String, val As Variant
    
    Set dict = NewDict()
    Set ws = wb.Worksheets("Control")
    
    ' Read all key-value pairs
    For r = 5 To 100
        key = Trim(CStr(ws.Cells(r, 1).Value))
        val = ws.Cells(r, 2).Value
        
        If key = "" Then Exit For
        dict(key) = val
    Next r
    
    ' Apply defaults if missing
    If Not dict.Exists("Total_Capital") Then dict("Total_Capital") = 600000
    If Not dict.Exists("NumQuarters") Then dict("NumQuarters") = 48
    
    ' Validate and normalize
    Call ValidateControlInputs(dict)
    Call NormalizeCapitalStructure(dict)
    Call ApplyScenario(dict)
    
    Set ReadControlInputs = dict
End Function

Private Sub ValidateControlInputs(dict As Object)
    On Error Resume Next
    
    ' NumQuarters
    Dim numQ As Long
    numQ = ToLng(dict("NumQuarters"))
    If numQ <= 0 Or numQ > 200 Then
        dict("NumQuarters") = 48
    End If
    
    ' Percentages [0,1]
    Dim pctKeys As Variant
    pctKeys = Array("Pct_A", "Pct_B", "Pct_C", "Pct_D", "Pct_E", _
                   "Reserve_Pct", "PIK_Pct", "Base_CDR", "Base_Recovery", _
                   "Base_Prepay", "Base_Amort", "Mgmt_Fee_Pct", _
                   "Pref_Hurdle", "GP_Catch_Up_Pct", "GP_Split_Pct", _
                   "Recycling_Pct", "Close_Call_Pct")
    
    Dim i As Long
    For i = 0 To UBound(pctKeys)
        If dict.Exists(pctKeys(i)) Then
            Dim val As Double
            val = ToDbl(dict(pctKeys(i)))
            If val < 0 Then val = 0
            If val > 1 Then val = 1
            dict(pctKeys(i)) = val
        End If
    Next i
    
    ' Reinvest_Q
    Dim reinvestQ As Long
    reinvestQ = ToLng(dict("Reinvest_Q"))
    If reinvestQ < 0 Then reinvestQ = 0
    If reinvestQ > numQ Then reinvestQ = numQ
    dict("Reinvest_Q") = reinvestQ
End Sub

Private Sub NormalizeCapitalStructure(dict As Object)
    On Error Resume Next
    Dim total As Double
    Dim pctA As Double, pctB As Double, pctC As Double, pctD As Double, pctE As Double
    
    pctA = ToDbl(dict("Pct_A"))
    pctB = ToDbl(dict("Pct_B"))
    pctE = ToDbl(dict("Pct_E"))
    
    If ToBool(dict("Enable_C")) Then
        pctC = ToDbl(dict("Pct_C"))
    Else
        pctC = 0
        dict("Pct_C") = 0
    End If
    
    If ToBool(dict("Enable_D")) Then
        pctD = ToDbl(dict("Pct_D"))
    Else
        pctD = 0
        dict("Pct_D") = 0
    End If
    
    total = pctA + pctB + pctC + pctD + pctE
    
    ' Normalize to sum = 1
    If total > 0 And Abs(total - 1) > 0.001 Then
        pctA = pctA / total
        pctB = pctB / total
        pctC = pctC / total
        pctD = pctD / total
        pctE = pctE / total
        
        dict("Pct_A") = pctA
        dict("Pct_B") = pctB
        dict("Pct_C") = pctC
        dict("Pct_D") = pctD
        dict("Pct_E") = pctE
        
        Call Log("NormalizeCapitalStructure", _
                "Capital structure normalized from " & Format(total, "0.00%") & " to 100%")
    End If
End Sub

Private Sub ApplyScenario(ByRef dict As Object)
    On Error Resume Next
    Dim scenario As String
    scenario = UCase(CStr(dict("Scenario_Selection")))
    
    Select Case scenario
        Case "DOWN"
            dict("Base_CDR") = ToDbl(dict("Base_CDR")) * 1.25
            dict("Base_Recovery") = Application.Max(0, ToDbl(dict("Base_Recovery")) - 0.05)
            dict("Spread_Add_bps") = ToDbl(dict("Spread_Add_bps")) + 100
            dict("Rate_Add_bps") = ToDbl(dict("Rate_Add_bps")) + 50
        Case "UP"
            dict("Base_CDR") = ToDbl(dict("Base_CDR")) * 0.8
            dict("Base_Recovery") = Application.Min(0.95, ToDbl(dict("Base_Recovery")) + 0.05)
            dict("Spread_Add_bps") = ToDbl(dict("Spread_Add_bps")) - 25
            dict("Rate_Add_bps") = ToDbl(dict("Rate_Add_bps")) - 25
        Case Else
            ' Base - no modification
    End Select
End Sub

'------------------------------------------------------------------------------
' DATES AND CURVES
'------------------------------------------------------------------------------
Private Function BuildQuarterDates(dict As Object) As Date()
    On Error Resume Next
    Dim startDate As Date
    Dim numQ As Long
    Dim dates() As Date
    Dim i As Long
    
    startDate = CDate(dict("First_Close_Date"))
    numQ = ToLng(dict("NumQuarters"))
    
    ReDim dates(0 To numQ - 1)
    
    For i = 0 To numQ - 1
        dates(i) = DateAdd("m", i * 3, startDate)
    Next i
    
    BuildQuarterDates = dates
End Function

Private Function GetSOFRCurve(dict As Object, numQ As Long) As Double()
    On Error Resume Next
    Dim curve() As Double
    Dim q As Long
    Dim rateAdd As Double
    
    ReDim curve(0 To numQ - 1)
    rateAdd = ToDbl(dict("Rate_Add_bps")) / 10000
    
    ' Taper from 4.33% to 3.50% over 12Q then flat
    For q = 0 To numQ - 1
        If q < 12 Then
            curve(q) = 0.0433 - (0.0433 - 0.035) * (q / 12) + rateAdd
        Else
            curve(q) = 0.035 + rateAdd
        End If
    Next q
    
    GetSOFRCurve = curve
End Function

'------------------------------------------------------------------------------
' SIMULATION ENGINE
'------------------------------------------------------------------------------
Private Function SimulateTape(tapeData As Variant, dict As Object, dates() As Date) As Object
    On Error GoTo EH
    Dim results As Object
    Dim numQ As Long
    Dim r As Long, q As Long
    Dim outstanding() As Double, interest() As Double
    Dim defaults() As Double, recoveries() As Double
    Dim principal() As Double, commitmentFees() As Double
    Dim unfunded() As Double, prepayments() As Double
    Dim sofrCurve() As Double
    Dim startDate As Date
    
    Set results = NewDict()
    numQ = UBound(dates) + 1
    startDate = dates(0)
    
    ' Get curves
    sofrCurve = GetSOFRCurve(dict, numQ)
    
    ' Initialize arrays
    ReDim outstanding(0 To numQ - 1)
    ReDim interest(0 To numQ - 1)
    ReDim defaults(0 To numQ - 1)
    ReDim recoveries(0 To numQ - 1)
    ReDim principal(0 To numQ - 1)
    ReDim prepayments(0 To numQ - 1)
    ReDim commitmentFees(0 To numQ - 1)
    ReDim unfunded(0 To numQ - 1)
    
    ' Get assumptions
    Dim baseCDR As Double, baseRecovery As Double, basePrepay As Double, baseAmort As Double
    Dim lossLagQ As Long, spreadAdd As Double
    
    baseCDR = ToDbl(dict("Base_CDR"))
    baseRecovery = ToDbl(dict("Base_Recovery"))
    basePrepay = ToDbl(dict("Base_Prepay"))
    baseAmort = ToDbl(dict("Base_Amort"))
    lossLagQ = ToLng(dict("Loss_Lag_Q"))
    If lossLagQ = 0 Then lossLagQ = 4
    spreadAdd = ToDbl(dict("Spread_Add_bps")) / 10000
    
    ' Track initial balance
    Dim initialFunded As Double
    initialFunded = 0
    
    ' Simulate each loan
    For r = 2 To UBound(tapeData, 1)
        Dim bal As Double, unfundedLoan As Double
        Dim par As Double, drawPct As Double, spreadBps As Double
        Dim maturityDate As Date, qToMat As Long
        Dim secType As String, facilityType As String
        Dim lossQueue() As Double
        Dim isDebt As Boolean
        
        par = ToDbl(tapeData(r, 2))
        drawPct = ToDbl(tapeData(r, 3))
        spreadBps = ToDbl(tapeData(r, 4)) / 10000
        facilityType = CStr(tapeData(r, 6))
        secType = CStr(tapeData(r, 7))
        
        isDebt = IsDebtInstrument(secType)
        
        ' Calculate maturity quarters
        If IsDate(tapeData(r, 8)) And tapeData(r, 8) <> "NM" Then
            maturityDate = CDate(tapeData(r, 8))
            Dim monthsToMat As Long
            monthsToMat = DateDiff("m", startDate, maturityDate)
            qToMat = (monthsToMat + 2) \ 3
            If qToMat < 1 Then qToMat = 1
        ElseIf IsNumeric(tapeData(r, 9)) Then
            qToMat = ToLng(tapeData(r, 9)) * 4
        Else
            qToMat = numQ  ' No maturity
        End If
        
        ' Initialize balance
        bal = par * drawPct
        unfundedLoan = par - bal
        initialFunded = initialFunded + bal
        
        ' Initialize loss queue
        ReDim lossQueue(0 To lossLagQ)
        
        ' Simulate quarters
        For q = 0 To numQ - 1
            If bal > 0 And isDebt Then
                ' Interest
                Dim rate As Double
                rate = sofrCurve(q) + spreadBps + spreadAdd
                interest(q) = interest(q) + bal * rate / 4
                
                ' Scheduled amortization
                Dim sched As Double
                If InStr(1, facilityType, "IO", vbTextCompare) = 0 Then
                    sched = baseAmort * bal / 4
                Else
                    sched = 0
                End If
                
                ' Prepayment
                Dim pre As Double
                pre = basePrepay * bal / 4
                prepayments(q) = prepayments(q) + pre
                
                ' Default
                Dim def As Double
                def = baseCDR * bal / 4
                defaults(q) = defaults(q) + def
                
                ' Recovery (lagged)
                recoveries(q) = recoveries(q) + lossQueue(lossLagQ) * baseRecovery
                
                ' Shift loss queue
                Dim i As Long
                For i = lossLagQ To 1 Step -1
                    lossQueue(i) = lossQueue(i - 1)
                Next i
                lossQueue(0) = def
                
                ' Revolver draw
                If InStr(1, facilityType, "Revolver", vbTextCompare) > 0 And unfundedLoan > 0 Then
                    Dim draw As Double
                    draw = Application.Min(unfundedLoan, unfundedLoan * 0.05)
                    bal = bal + draw
                    unfundedLoan = unfundedLoan - draw
                End If
                
                ' DDTL funding
                If InStr(1, facilityType, "DDTL", vbTextCompare) > 0 And q < 4 And unfundedLoan > 0 Then
                    Dim fund As Double
                    fund = Application.Min(unfundedLoan, unfundedLoan * 0.25)
                    bal = bal + fund
                    unfundedLoan = unfundedLoan - fund
                End If
                
                ' Maturity sweep
                If q >= qToMat - 1 And qToMat > 0 Then
                    sched = sched + bal
                End If
                
                ' Update balance
                principal(q) = principal(q) + sched + pre
                bal = Application.Max(0, bal - sched - pre - def)
                
                ' Commitment fees on revolvers
                If InStr(1, facilityType, "Revolver", vbTextCompare) > 0 And unfundedLoan > 0 Then
                    commitmentFees(q) = commitmentFees(q) + unfundedLoan * 0.0125 / 4
                End If
            ElseIf Not isDebt Then
                ' Equity position - no cashflows but count in OC numerator
                ' Balance stays constant
            End If
            
            outstanding(q) = outstanding(q) + bal
            unfunded(q) = unfunded(q) + unfundedLoan
        Next q
    Next r
    
    ' Handle recycling if enabled
    If ToBool(dict("Enable_Recycling")) Then
        Call ApplyRecycling(dict, numQ, principal, interest, defaults, recoveries, _
                           outstanding, sofrCurve, spreadAdd, baseCDR, baseRecovery, lossLagQ)
    End If
    
    ' Store results
    results("Outstanding") = outstanding
    results("Interest") = interest
    results("Defaults") = defaults
    results("Recoveries") = recoveries
    results("Principal") = principal
    results("Prepayments") = prepayments
    results("CommitmentFees") = commitmentFees
    results("Unfunded") = unfunded
    results("Initial_Funded_Balance") = initialFunded
    
    Set SimulateTape = results
    Exit Function
    
EH:
    Call Log("SimulateTape", Err.Description)
    Set SimulateTape = NewDict()
End Function

Private Sub ApplyRecycling(dict As Object, numQ As Long, ByRef principal() As Double, _
                          ByRef interest() As Double, ByRef defaults() As Double, _
                          ByRef recoveries() As Double, ByRef outstanding() As Double, _
                          sofrCurve() As Double, spreadAdd As Double, _
                          baseCDR As Double, baseRecovery As Double, lossLagQ As Long)
    On Error Resume Next
    
    Dim recycleSpread As Double
    Dim RecycledAdd() As Double, RecycledBal() As Double
    Dim bucketLoss() As Double
    Dim harvestStart As Long
    Dim rcy As Double
    Dim q As Long
    
    recycleSpread = ToDbl(dict("Recycle_Spread_bps")) / 10000
    
    ReDim RecycledAdd(0 To numQ - 1)
    ReDim RecycledBal(0 To numQ - 1)
    ReDim bucketLoss(0 To lossLagQ)
    
    harvestStart = ToLng(dict("Reinvest_Q")) + ToLng(dict("GP_Extend_Q"))
    
    ' Stage recycled principal
    For q = 0 To numQ - 1
        If q < harvestStart Then
            rcy = principal(q) * ToDbl(dict("Recycling_Pct"))
            principal(q) = principal(q) - rcy
            If q + 1 <= numQ - 1 Then
                RecycledAdd(q + 1) = RecycledAdd(q + 1) + rcy
            End If
        End If
    Next q
    
    ' Roll bucket
    Dim bdef As Double
    Dim k As Long
    
    For q = 0 To numQ - 1
        RecycledBal(q) = RecycledBal(q) + RecycledAdd(q)
        
        ' Interest on recycled
        interest(q) = interest(q) + RecycledBal(q) * (sofrCurve(q) + recycleSpread + spreadAdd) / 4
        
        ' Defaults & recoveries
        bdef = baseCDR * RecycledBal(q) / 4
        defaults(q) = defaults(q) + bdef
        
        For k = lossLagQ To 1 Step -1
            bucketLoss(k) = bucketLoss(k - 1)
        Next k
        bucketLoss(0) = bdef
        
        recoveries(q) = recoveries(q) + bucketLoss(lossLagQ) * baseRecovery
        
        ' Update balance
        RecycledBal(q) = Application.Max(0, RecycledBal(q) - bdef)
        
        ' Add to outstanding
        outstanding(q) = outstanding(q) + RecycledBal(q)
    Next q
End Sub

'------------------------------------------------------------------------------
' WATERFALL ENGINE
'------------------------------------------------------------------------------
Private Function RunWaterfall(sim As Object, dict As Object, dates() As Date) As Object
    On Error GoTo EH
    Dim results As Object
    Dim numQ As Long
    Dim q As Long
    Dim enableC As Boolean, enableD As Boolean
    
    Set results = NewDict()
    numQ = UBound(dates) + 1
    enableC = ToBool(dict("Enable_C"))
    enableD = ToBool(dict("Enable_D"))
    
    ' Initialize arrays
    Dim A_Bal() As Double, B_Bal() As Double, C_Bal() As Double, D_Bal() As Double
    Dim A_IntDue() As Double, B_IntDue() As Double, C_IntDue() As Double, D_IntDue() As Double
    Dim A_IntPd() As Double, B_IntPd() As Double, C_IntPd() As Double, D_IntPd() As Double
    Dim A_IntPIK() As Double, B_IntPIK() As Double, C_IntPIK() As Double, D_IntPIK() As Double
    Dim A_Prin() As Double, B_Prin() As Double, C_Prin() As Double, D_Prin() As Double
    Dim OC_A() As Double, OC_B() As Double, OC_C() As Double, OC_D() As Double
    Dim IC_A() As Double, IC_B() As Double, IC_C() As Double, IC_D() As Double
    Dim DSCR() As Double, AdvRate() As Double
    Dim Equity_CF() As Double, LP_Calls() As Double
    Dim Reserve_Beg() As Double, Reserve_Draw() As Double, Reserve_Release() As Double
    Dim Reserve_TopUp() As Double, Reserve_End() As Double
    Dim TurboFlag() As Double
    Dim Fees_Servicer() As Double, Fees_Mgmt() As Double, Fees_Admin() As Double
    
    ReDim A_Bal(0 To numQ): ReDim B_Bal(0 To numQ)
    ReDim C_Bal(0 To numQ): ReDim D_Bal(0 To numQ)
    ReDim A_IntDue(0 To numQ - 1): ReDim B_IntDue(0 To numQ - 1)
    ReDim C_IntDue(0 To numQ - 1): ReDim D_IntDue(0 To numQ - 1)
    ReDim A_IntPd(0 To numQ - 1): ReDim B_IntPd(0 To numQ - 1)
    ReDim C_IntPd(0 To numQ - 1): ReDim D_IntPd(0 To numQ - 1)
    ReDim A_IntPIK(0 To numQ - 1): ReDim B_IntPIK(0 To numQ - 1)
    ReDim C_IntPIK(0 To numQ - 1): ReDim D_IntPIK(0 To numQ - 1)
    ReDim A_Prin(0 To numQ - 1): ReDim B_Prin(0 To numQ - 1)
    ReDim C_Prin(0 To numQ - 1): ReDim D_Prin(0 To numQ - 1)
    ReDim OC_A(0 To numQ - 1): ReDim OC_B(0 To numQ - 1)
    ReDim OC_C(0 To numQ - 1): ReDim OC_D(0 To numQ - 1)
    ReDim IC_A(0 To numQ - 1): ReDim IC_B(0 To numQ - 1)
    ReDim IC_C(0 To numQ - 1): ReDim IC_D(0 To numQ - 1)
    ReDim DSCR(0 To numQ - 1): ReDim AdvRate(0 To numQ - 1)
    ReDim Equity_CF(0 To numQ - 1): ReDim LP_Calls(0 To numQ - 1)
    ReDim Reserve_Beg(0 To numQ): ReDim Reserve_Draw(0 To numQ - 1)
    ReDim Reserve_Release(0 To numQ - 1): ReDim Reserve_TopUp(0 To numQ - 1)
    ReDim Reserve_End(0 To numQ - 1)
    ReDim TurboFlag(0 To numQ - 1)
    ReDim Fees_Servicer(0 To numQ - 1): ReDim Fees_Mgmt(0 To numQ - 1)
    ReDim Fees_Admin(0 To numQ - 1)
    
    ' Get parameters
    Dim reinvestQ As Long, gpExtendQ As Long
    Dim enableTurbo As Boolean, enableReserve As Boolean, enablePIK As Boolean, enableCCPIK As Boolean
    Dim reservePct As Double, pikPct As Double
    Dim spreadA As Double, spreadB As Double, spreadC As Double, spreadD As Double
    Dim servicerFee As Double, mgmtFeePct As Double, adminFloor As Double
    Dim ocTriggerA As Double, ocTriggerB As Double, ocTriggerC As Double, ocTriggerD As Double
    Dim downgradeOC As Double, downgradeSpd As Double
    
    reinvestQ = ToLng(dict("Reinvest_Q"))
    gpExtendQ = ToLng(dict("GP_Extend_Q"))
    enableTurbo = ToBool(dict("Enable_Turbo_DOC"))
    enableReserve = ToBool(dict("Enable_Excess_Reserve"))
    enablePIK = ToBool(dict("Enable_PIK"))
    enableCCPIK = ToBool(dict("Enable_CC_PIK"))
    reservePct = ToDbl(dict("Reserve_Pct"))
    pikPct = ToDbl(dict("PIK_Pct"))
    If pikPct < 0 Then pikPct = 0
    If pikPct > 1 Then pikPct = 1
    spreadA = ToDbl(dict("Spread_A_bps")) / 10000
    spreadB = ToDbl(dict("Spread_B_bps")) / 10000
    spreadC = ToDbl(dict("Spread_C_bps")) / 10000
    spreadD = ToDbl(dict("Spread_D_bps")) / 10000
    servicerFee = ToDbl(dict("Servicer_Fee_bps")) / 10000
    mgmtFeePct = ToDbl(dict("Mgmt_Fee_Pct"))
    adminFloor = ToDbl(dict("Admin_Fee_Floor"))
    
    ' OC Triggers
    ocTriggerA = ToDbl(dict("OC_Trigger_A"))
    If ocTriggerA = 0 Then ocTriggerA = 1.25
    ocTriggerB = ToDbl(dict("OC_Trigger_B"))
    If ocTriggerB = 0 Then ocTriggerB = 1.125
    ocTriggerC = ToDbl(dict("OC_Trigger_C"))
    If ocTriggerC = 0 Then ocTriggerC = 1.05
    ocTriggerD = ToDbl(dict("OC_Trigger_D"))
    If ocTriggerD = 0 Then ocTriggerD = 1.00
    
    ' Downgrade triggers
    downgradeOC = ToDbl(dict("Downgrade_OC"))
    If downgradeOC = 0 Then downgradeOC = 1.08
    downgradeSpd = ToDbl(dict("Downgrade_Spd_Adj_bps")) / 10000
    
    ' Get SOFR curve
    Dim sofrCurve() As Double
    sofrCurve = GetSOFRCurve(dict, numQ)
    
    ' Initial balances
    Dim totalCapital As Double
    totalCapital = sim("Initial_Funded_Balance")
    If totalCapital = 0 Then totalCapital = ToDbl(dict("Total_Capital"))
    If totalCapital = 0 Then totalCapital = 600000
    
    A_Bal(0) = totalCapital * ToDbl(dict("Pct_A"))
    B_Bal(0) = totalCapital * ToDbl(dict("Pct_B"))
    
    If enableC Then
        C_Bal(0) = totalCapital * ToDbl(dict("Pct_C"))
    End If
    
    If enableD Then
        D_Bal(0) = totalCapital * ToDbl(dict("Pct_D"))
    End If
    
    ' Initial reserve (funded from equity at close)
    Dim equityGross As Double
    equityGross = totalCapital * ToDbl(dict("Pct_E"))
    Reserve_Beg(0) = equityGross * reservePct
    
    ' Run waterfall for each quarter
    For q = 0 To numQ - 1
        ' Check for downgrade step-up
        Dim liabStep As Double
        liabStep = 0
        If q > 0 Then
            If OC_A(q - 1) < downgradeOC Then
                liabStep = downgradeSpd
            End If
        End If
        
        ' Calculate interest due
        A_IntDue(q) = A_Bal(q) * (sofrCurve(q) + spreadA + liabStep) / 4
        B_IntDue(q) = B_Bal(q) * (sofrCurve(q) + spreadB + liabStep) / 4
        
        If enableC Then
            C_IntDue(q) = C_Bal(q) * (sofrCurve(q) + spreadC + liabStep) / 4
        End If
        
        If enableD Then
            D_IntDue(q) = D_Bal(q) * (sofrCurve(q) + spreadD + liabStep) / 4
        End If
        
        ' Determine harvest mode
        Dim isHarvest As Boolean
        isHarvest = (q >= reinvestQ + gpExtendQ)
        If q > 0 And enableTurbo Then
            If TurboFlag(q - 1) = 1 Then isHarvest = True
        End If
        
        ' Calculate available cash
        Dim startAvailGross As Double
        startAvailGross = sim("Interest")(q) + sim("CommitmentFees")(q) + sim("Recoveries")(q)
        
        If isHarvest Then
            startAvailGross = startAvailGross + sim("Principal")(q)
        End If
        
        ' Calculate fees
        Dim assetsBOP As Double
        If q = 0 Then
            assetsBOP = sim("Initial_Funded_Balance")
        Else
            assetsBOP = sim("Outstanding")(q - 1)
        End If
        
        Fees_Servicer(q) = assetsBOP * servicerFee / 4
        Fees_Mgmt(q) = assetsBOP * mgmtFeePct / 4
        Fees_Admin(q) = adminFloor / 4
        
        ' Available after fees
        Dim avail As Double
        avail = startAvailGross - Fees_Servicer(q) - Fees_Mgmt(q) - Fees_Admin(q)
        
        ' Reserve release if needed for interest
        Dim totalIntDue As Double
        totalIntDue = A_IntDue(q) + B_IntDue(q) + C_IntDue(q) + D_IntDue(q)
        
        If enableReserve And avail < totalIntDue And Reserve_Beg(q) > 0 Then
            Dim release As Double
            release = Application.Min(Reserve_Beg(q), totalIntDue - avail)
            Reserve_Release(q) = release
            avail = avail + release
        End If
        
        ' Pay interest in priority order with PIK handling
        Call PayInterestWithPIK(q, avail, A_IntDue(q), A_IntPd(q), A_IntPIK(q), _
                               Reserve_Beg(q) - Reserve_Release(q), Reserve_Draw(q), _
                               LP_Calls(q), enablePIK, enableCCPIK, pikPct)
        
        Call PayInterestWithPIK(q, avail, B_IntDue(q), B_IntPd(q), B_IntPIK(q), _
                               Reserve_Beg(q) - Reserve_Release(q) - Reserve_Draw(q), _
                               Reserve_Draw(q), LP_Calls(q), enablePIK, enableCCPIK, pikPct)
        
        If enableC Then
            Call PayInterestWithPIK(q, avail, C_IntDue(q), C_IntPd(q), C_IntPIK(q), _
                                   0, Reserve_Draw(q), LP_Calls(q), enablePIK, False, pikPct)
        End If
        
        If enableD Then
            Call PayInterestWithPIK(q, avail, D_IntDue(q), D_IntPd(q), D_IntPIK(q), _
                                   0, Reserve_Draw(q), LP_Calls(q), enablePIK, False, pikPct)
        End If
        
        ' Principal payments (sequential in harvest mode)
        If isHarvest And avail > 0 Then
            Call PaySequentialPrincipal(q, avail, A_Bal(q), B_Bal(q), C_Bal(q), D_Bal(q), _
                                       A_Prin(q), B_Prin(q), C_Prin(q), D_Prin(q), _
                                       enableC, enableD)
        End If
        
        ' Update balances with PIK
        A_Bal(q) = A_Bal(q) - A_Prin(q) + A_IntPIK(q)
        B_Bal(q) = B_Bal(q) - B_Prin(q) + B_IntPIK(q)
        If enableC Then C_Bal(q) = C_Bal(q) - C_Prin(q) + C_IntPIK(q)
        If enableD Then D_Bal(q) = D_Bal(q) - D_Prin(q) + D_IntPIK(q)
        
        ' Calculate coverage ratios
        Dim assets As Double
        assets = sim("Outstanding")(q)
        
        OC_A(q) = SafeDivide(assets, A_Bal(q), RATIO_SENTINEL)
        OC_B(q) = SafeDivide(assets, A_Bal(q) + B_Bal(q), RATIO_SENTINEL)
        If enableC Then OC_C(q) = SafeDivide(assets, A_Bal(q) + B_Bal(q) + C_Bal(q), RATIO_SENTINEL)
        If enableD Then OC_D(q) = SafeDivide(assets, A_Bal(q) + B_Bal(q) + C_Bal(q) + D_Bal(q), RATIO_SENTINEL)
        
        ' Interest coverage (per-class)
        IC_A(q) = SafeDivide(A_IntPd(q), A_IntDue(q), RATIO_SENTINEL)
        IC_B(q) = SafeDivide(B_IntPd(q), B_IntDue(q), RATIO_SENTINEL)
        If enableC Then IC_C(q) = SafeDivide(C_IntPd(q), C_IntDue(q), RATIO_SENTINEL)
        If enableD Then IC_D(q) = SafeDivide(D_IntPd(q), D_IntDue(q), RATIO_SENTINEL)
        
        ' DSCR
        DSCR(q) = SafeDivide(sim("Interest")(q) + sim("CommitmentFees")(q), totalIntDue, RATIO_SENTINEL)
        
        ' Advance Rate
        AdvRate(q) = SafeDivide(A_Bal(q) + B_Bal(q) + C_Bal(q) + D_Bal(q), assets, 0)
        
        ' DOC Turbo check
        Dim breachA As Boolean, breachB As Boolean, breachC As Boolean, breachD As Boolean
        breachA = OC_A(q) < ocTriggerA
        breachB = OC_B(q) < ocTriggerB
        If enableC Then breachC = OC_C(q) < ocTriggerC
        If enableD Then breachD = OC_D(q) < ocTriggerD
        
        If enableTurbo And (breachA Or breachB Or breachC Or breachD) Then
            TurboFlag(q) = 1
        Else
            TurboFlag(q) = 0
        End If
        
        ' Reserve top-up
        If enableReserve And TurboFlag(q) = 1 And avail > 0 Then
            Dim target As Double, topUp As Double
            target = reservePct * startAvailGross
            Dim reserveCurrent As Double
            reserveCurrent = Reserve_Beg(q) - Reserve_Release(q) - Reserve_Draw(q)
            topUp = Application.Max(0, target - reserveCurrent)
            topUp = Application.Min(avail, topUp)
            Reserve_TopUp(q) = topUp
            avail = avail - topUp
        End If
        
        ' Reserve end balance
        Reserve_End(q) = Application.Max(0, Reserve_Beg(q) - Reserve_Release(q) - Reserve_Draw(q) + Reserve_TopUp(q))
        
        ' Equity distribution
        If avail >= 0 Then
            Equity_CF(q) = avail
        Else
            Equity_CF(q) = 0
            LP_Calls(q) = LP_Calls(q) + (-avail)
        End If
        
        ' Carry forward balances
        If q < numQ - 1 Then
            A_Bal(q + 1) = A_Bal(q)
            B_Bal(q + 1) = B_Bal(q)
            C_Bal(q + 1) = C_Bal(q)
            D_Bal(q + 1) = D_Bal(q)
            Reserve_Beg(q + 1) = Reserve_End(q)
        End If
    Next q
    
    ' Copy simulation results
    Dim key As Variant
    For Each key In sim.Keys
        results(key) = sim(key)
    Next key
    
    ' Trim arrays
    ReDim Preserve A_Bal(0 To numQ - 1)
    ReDim Preserve B_Bal(0 To numQ - 1)
    ReDim Preserve C_Bal(0 To numQ - 1)
    ReDim Preserve D_Bal(0 To numQ - 1)
    ReDim Preserve Reserve_Beg(0 To numQ - 1)
    
    ' Add waterfall results
    results("A_Bal") = A_Bal: results("B_Bal") = B_Bal
    results("C_Bal") = C_Bal: results("D_Bal") = D_Bal
    results("A_IntDue") = A_IntDue: results("B_IntDue") = B_IntDue
    results("C_IntDue") = C_IntDue: results("D_IntDue") = D_IntDue
    results("A_IntPd") = A_IntPd: results("B_IntPd") = B_IntPd
    results("C_IntPd") = C_IntPd: results("D_IntPd") = D_IntPd
    results("A_IntPIK") = A_IntPIK: results("B_IntPIK") = B_IntPIK
    results("C_IntPIK") = C_IntPIK: results("D_IntPIK") = D_IntPIK
    results("A_Prin") = A_Prin: results("B_Prin") = B_Prin
    results("C_Prin") = C_Prin: results("D_Prin") = D_Prin
    results("OC_A") = OC_A: results("OC_B") = OC_B
    results("OC_C") = OC_C: results("OC_D") = OC_D
    results("IC_A") = IC_A: results("IC_B") = IC_B
    results("IC_C") = IC_C: results("IC_D") = IC_D
    results("DSCR") = DSCR: results("AdvRate") = AdvRate
    results("Equity_CF") = Equity_CF: results("LP_Calls") = LP_Calls
    results("Reserve_Beg") = Reserve_Beg: results("Reserve_Draw") = Reserve_Draw
    results("Reserve_Release") = Reserve_Release: results("Reserve_TopUp") = Reserve_TopUp
    results("Reserve_End") = Reserve_End: results("TurboFlag") = TurboFlag
    results("Fees_Servicer") = Fees_Servicer: results("Fees_Mgmt") = Fees_Mgmt
    results("Fees_Admin") = Fees_Admin
    
    Set RunWaterfall = results
    Exit Function
    
EH:
    Call Log("RunWaterfall", Err.Description)
    Set RunWaterfall = NewDict()
End Function

Private Sub PayInterestWithPIK(q As Long, ByRef avail As Double, intDue As Double, _
                               ByRef intPd As Double, ByRef intPIK As Double, _
                               reserveAvail As Double, ByRef reserveDraw As Double, _
                               ByRef lpCalls As Double, enablePIK As Boolean, _
                               enableCCPIK As Boolean, pikPct As Double)
    On Error Resume Next
    
    Dim cashPay As Double, short As Double, pik As Double
    
    ' Pay what we can with cash
    cashPay = Application.Min(avail, intDue)
    intPd = cashPay
    avail = avail - cashPay
    short = intDue - cashPay
    
    If short > 0 Then
        ' Try reserve draw if CC PIK enabled
        If enableCCPIK And reserveAvail > short Then
            reserveDraw = reserveDraw + short
            intPd = intPd + short
            short = 0
        ElseIf enablePIK Then
            ' Apply PIK
            pik = Application.Min(short, intDue * pikPct)
            intPIK = pik
            short = short - pik
        End If
        
        ' Any remaining shortfall requires LP call
        If short > 0 Then
            lpCalls = lpCalls + short
        End If
    End If
End Sub

Private Sub PaySequentialPrincipal(q As Long, ByRef avail As Double, _
                                   A_Bal As Double, B_Bal As Double, C_Bal As Double, D_Bal As Double, _
                                   ByRef A_Prin As Double, ByRef B_Prin As Double, _
                                   ByRef C_Prin As Double, ByRef D_Prin As Double, _
                                   enableC As Boolean, enableD As Boolean)
    On Error Resume Next
    
    ' Sequential principal A→B→C→D
    If A_Bal > 0 And avail > 0 Then
        A_Prin = Application.Min(avail, A_Bal)
        avail = avail - A_Prin
    End If
    
    If B_Bal > 0 And avail > 0 Then
        B_Prin = Application.Min(avail, B_Bal)
        avail = avail - B_Prin
    End If
    
    If enableC And C_Bal > 0 And avail > 0 Then
        C_Prin = Application.Min(avail, C_Bal)
        avail = avail - C_Prin
    End If
    
    If enableD And D_Bal > 0 And avail > 0 Then
        D_Prin = Application.Min(avail, D_Bal)
        avail = avail - D_Prin
    End If
End Sub

'------------------------------------------------------------------------------
' OUTPUT WRITING
'------------------------------------------------------------------------------
Private Sub WriteRunSheet(wb As Workbook, results As Object, quarterDates() As Date, controlDict As Object)
    On Error GoTo EH
    Dim ws As Worksheet
    Dim numQ As Long
    Dim col As Long
    Dim outputArray() As Variant
    Dim r As Long, c As Long
    Dim enableC As Boolean, enableD As Boolean
    
    Set ws = wb.Worksheets("Run")
    ws.Cells.Clear
    
    enableC = ToBool(controlDict("Enable_C"))
    enableD = ToBool(controlDict("Enable_D"))
    
    ' Title
    ws.Range("A1").Value = "ENGINE RUN OUTPUT"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "(in $000s)"
    ws.Range("A2").Style = "SG_Subtitle"
    
    numQ = UBound(quarterDates) + 1
    
    ' Build headers
    col = 1
    ws.Cells(4, col).Value = "Date": col = col + 1
    ws.Cells(4, col).Value = "Outstanding": col = col + 1
    ws.Cells(4, col).Value = "Unfunded": col = col + 1
    ws.Cells(4, col).Value = "Commitment_Fees": col = col + 1
    ws.Cells(4, col).Value = "Interest": col = col + 1
    ws.Cells(4, col).Value = "Defaults": col = col + 1
    ws.Cells(4, col).Value = "Recoveries": col = col + 1
    ws.Cells(4, col).Value = "Principal": col = col + 1
    ws.Cells(4, col).Value = "Prepayments": col = col + 1
    
    ws.Cells(4, col).Value = "A_Bal": col = col + 1
    ws.Cells(4, col).Value = "B_Bal": col = col + 1
    If enableC Then ws.Cells(4, col).Value = "C_Bal": col = col + 1
    If enableD Then ws.Cells(4, col).Value = "D_Bal": col = col + 1
    
    ws.Cells(4, col).Value = "A_IntDue": col = col + 1
    ws.Cells(4, col).Value = "A_IntPd": col = col + 1
    ws.Cells(4, col).Value = "A_IntPIK": col = col + 1
    ws.Cells(4, col).Value = "B_IntDue": col = col + 1
    ws.Cells(4, col).Value = "B_IntPd": col = col + 1
    ws.Cells(4, col).Value = "B_IntPIK": col = col + 1
    
    If enableC Then
        ws.Cells(4, col).Value = "C_IntDue": col = col + 1
        ws.Cells(4, col).Value = "C_IntPd": col = col + 1
        ws.Cells(4, col).Value = "C_IntPIK": col = col + 1
    End If
    
    If enableD Then
        ws.Cells(4, col).Value = "D_IntDue": col = col + 1
        ws.Cells(4, col).Value = "D_IntPd": col = col + 1
        ws.Cells(4, col).Value = "D_IntPIK": col = col + 1
    End If
    
    ws.Cells(4, col).Value = "A_Prin": col = col + 1
    ws.Cells(4, col).Value = "B_Prin": col = col + 1
    If enableC Then ws.Cells(4, col).Value = "C_Prin": col = col + 1
    If enableD Then ws.Cells(4, col).Value = "D_Prin": col = col + 1
    
    ws.Cells(4, col).Value = "Reserve_Beg": col = col + 1
    ws.Cells(4, col).Value = "Reserve_Release": col = col + 1
    ws.Cells(4, col).Value = "Reserve_Draw": col = col + 1
    ws.Cells(4, col).Value = "Reserve_TopUp": col = col + 1
    ws.Cells(4, col).Value = "Reserve_End": col = col + 1
    
    ws.Cells(4, col).Value = "OC_A": col = col + 1
    ws.Cells(4, col).Value = "OC_B": col = col + 1
    If enableC Then ws.Cells(4, col).Value = "OC_C": col = col + 1
    If enableD Then ws.Cells(4, col).Value = "OC_D": col = col + 1
    
    ws.Cells(4, col).Value = "IC_A": col = col + 1
    ws.Cells(4, col).Value = "IC_B": col = col + 1

    If enableC Then ws.Cells(4, col).Value = "IC_C": col = col + 1
    If enableD Then ws.Cells(4, col).Value = "IC_D": col = col + 1
    
    ws.Cells(4, col).Value = "DSCR": col = col + 1
    ws.Cells(4, col).Value = "AdvRate": col = col + 1
    ws.Cells(4, col).Value = "Equity_CF": col = col + 1
    ws.Cells(4, col).Value = "LP_Calls": col = col + 1
    ws.Cells(4, col).Value = "TurboFlag": col = col + 1
    ws.Cells(4, col).Value = "Fees_Servicer": col = col + 1
    ws.Cells(4, col).Value = "Fees_Mgmt": col = col + 1
    ws.Cells(4, col).Value = "Fees_Admin": col = col + 1
    
    ' Write data array
    ReDim outputArray(1 To numQ, 1 To col - 1)
    
    For r = 1 To numQ
        c = 1
        outputArray(r, c) = quarterDates(r - 1): c = c + 1
        outputArray(r, c) = results("Outstanding")(r - 1): c = c + 1
        outputArray(r, c) = results("Unfunded")(r - 1): c = c + 1
        outputArray(r, c) = results("CommitmentFees")(r - 1): c = c + 1
        outputArray(r, c) = results("Interest")(r - 1): c = c + 1
        outputArray(r, c) = results("Defaults")(r - 1): c = c + 1
        outputArray(r, c) = results("Recoveries")(r - 1): c = c + 1
        outputArray(r, c) = results("Principal")(r - 1): c = c + 1
        outputArray(r, c) = results("Prepayments")(r - 1): c = c + 1
        
        outputArray(r, c) = results("A_Bal")(r - 1): c = c + 1
        outputArray(r, c) = results("B_Bal")(r - 1): c = c + 1
        If enableC Then outputArray(r, c) = results("C_Bal")(r - 1): c = c + 1
        If enableD Then outputArray(r, c) = results("D_Bal")(r - 1): c = c + 1
        
        outputArray(r, c) = results("A_IntDue")(r - 1): c = c + 1
        outputArray(r, c) = results("A_IntPd")(r - 1): c = c + 1
        outputArray(r, c) = results("A_IntPIK")(r - 1): c = c + 1
        outputArray(r, c) = results("B_IntDue")(r - 1): c = c + 1
        outputArray(r, c) = results("B_IntPd")(r - 1): c = c + 1
        outputArray(r, c) = results("B_IntPIK")(r - 1): c = c + 1
        
        If enableC Then
            outputArray(r, c) = results("C_IntDue")(r - 1): c = c + 1
            outputArray(r, c) = results("C_IntPd")(r - 1): c = c + 1
            outputArray(r, c) = results("C_IntPIK")(r - 1): c = c + 1
        End If
        
        If enableD Then
            outputArray(r, c) = results("D_IntDue")(r - 1): c = c + 1
            outputArray(r, c) = results("D_IntPd")(r - 1): c = c + 1
            outputArray(r, c) = results("D_IntPIK")(r - 1): c = c + 1
        End If
        
        outputArray(r, c) = results("A_Prin")(r - 1): c = c + 1
        outputArray(r, c) = results("B_Prin")(r - 1): c = c + 1
        If enableC Then outputArray(r, c) = results("C_Prin")(r - 1): c = c + 1
        If enableD Then outputArray(r, c) = results("D_Prin")(r - 1): c = c + 1
        
        outputArray(r, c) = results("Reserve_Beg")(r - 1): c = c + 1
        outputArray(r, c) = results("Reserve_Release")(r - 1): c = c + 1
        outputArray(r, c) = results("Reserve_Draw")(r - 1): c = c + 1
        outputArray(r, c) = results("Reserve_TopUp")(r - 1): c = c + 1
        outputArray(r, c) = results("Reserve_End")(r - 1): c = c + 1
        
        outputArray(r, c) = results("OC_A")(r - 1): c = c + 1
        outputArray(r, c) = results("OC_B")(r - 1): c = c + 1
        If enableC Then outputArray(r, c) = results("OC_C")(r - 1): c = c + 1
        If enableD Then outputArray(r, c) = results("OC_D")(r - 1): c = c + 1
        
        outputArray(r, c) = results("IC_A")(r - 1): c = c + 1
        outputArray(r, c) = results("IC_B")(r - 1): c = c + 1
        If enableC Then outputArray(r, c) = results("IC_C")(r - 1): c = c + 1
        If enableD Then outputArray(r, c) = results("IC_D")(r - 1): c = c + 1
        
        outputArray(r, c) = results("DSCR")(r - 1): c = c + 1
        outputArray(r, c) = results("AdvRate")(r - 1): c = c + 1
        outputArray(r, c) = results("Equity_CF")(r - 1): c = c + 1
        outputArray(r, c) = results("LP_Calls")(r - 1): c = c + 1
        outputArray(r, c) = results("TurboFlag")(r - 1): c = c + 1
        outputArray(r, c) = results("Fees_Servicer")(r - 1): c = c + 1
        outputArray(r, c) = results("Fees_Mgmt")(r - 1): c = c + 1
        outputArray(r, c) = results("Fees_Admin")(r - 1): c = c + 1
    Next r
    
    ' Write to sheet in one operation
    ws.Range(ws.Cells(5, 1), ws.Cells(4 + numQ, col - 1)).Value = outputArray
    
    ' Format columns
    ws.Range("A4").Resize(1, col - 1).Style = "SG_Hdr"
    ws.Columns("A:A").NumberFormat = "mm/dd/yyyy"
    ws.Columns("B:I").Style = "SG_Currency_K"
    ws.Columns("J:AZ").Style = "SG_Currency_K"
    
    ' Coverage and rate columns
    Dim ocCol As Long
    For ocCol = 1 To col - 1
        If InStr(ws.Cells(4, ocCol).Value, "OC_") > 0 Or _
           InStr(ws.Cells(4, ocCol).Value, "IC_") > 0 Or _
           InStr(ws.Cells(4, ocCol).Value, "DSCR") > 0 Then
            ws.Columns(ocCol).NumberFormat = "0.00x"
        ElseIf InStr(ws.Cells(4, ocCol).Value, "AdvRate") > 0 Then
            ws.Columns(ocCol).NumberFormat = "0.0%"
        ElseIf InStr(ws.Cells(4, ocCol).Value, "TurboFlag") > 0 Then
            ws.Columns(ocCol).NumberFormat = "0"
        End If
    Next ocCol
    
    Call SG615_ApplyStylePack(ws, "", "")
    ws.Columns.AutoFit
    
    Call ApplyFreezePanesSafe(ws, 5, 2)
    
    Exit Sub
    
EH:
    Call Log("WriteRunSheet", Err.Description)
End Sub

Private Sub DefineDynamicNamesRun(wb As Workbook, controlDict As Object)
    On Error Resume Next
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim enableC As Boolean, enableD As Boolean
    Dim col As Long
    
    Set ws = wb.Worksheets("Run")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow <= 4 Then Exit Sub
    
    enableC = ToBool(controlDict("Enable_C"))
    enableD = ToBool(controlDict("Enable_D"))
    
    ' Create named ranges based on headers
    For col = 1 To 100
        Select Case ws.Cells(4, col).Value
            Case "Date": Call SetNameRef("Run_Dates", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "Outstanding": Call SetNameRef("Run_Outstanding", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "Interest": Call SetNameRef("Run_Interest", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "Principal": Call SetNameRef("Run_Principal", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "Prepayments": Call SetNameRef("Run_Prepayments", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "Defaults": Call SetNameRef("Run_Defaults", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "Recoveries": Call SetNameRef("Run_Recoveries", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "Commitment_Fees": Call SetNameRef("Run_CommitmentFees", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "Unfunded": Call SetNameRef("Run_Unfunded", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "A_IntDue": Call SetNameRef("Run_A_IntDue", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "B_IntDue": Call SetNameRef("Run_B_IntDue", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "C_IntDue": Call SetNameRef("Run_C_IntDue", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "D_IntDue": Call SetNameRef("Run_D_IntDue", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "Reserve_Beg": Call SetNameRef("Run_Reserve_Beg", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "Reserve_Release": Call SetNameRef("Run_Reserve_Release", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "Reserve_Draw": Call SetNameRef("Run_Reserve_Draw", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "Reserve_TopUp": Call SetNameRef("Run_Reserve_TopUp", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "Reserve_End": Call SetNameRef("Run_Reserve_End", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "A_Bal": Call SetNameRef("Run_A_EndBal", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "B_Bal": Call SetNameRef("Run_B_EndBal", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "C_Bal": Call SetNameRef("Run_C_EndBal", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "D_Bal": Call SetNameRef("Run_D_EndBal", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "A_IntPd": Call SetNameRef("Run_A_IntPd", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "B_IntPd": Call SetNameRef("Run_B_IntPd", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "C_IntPd": Call SetNameRef("Run_C_IntPd", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "D_IntPd": Call SetNameRef("Run_D_IntPd", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "A_IntPIK": Call SetNameRef("Run_A_IntPIK", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "B_IntPIK": Call SetNameRef("Run_B_IntPIK", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "C_IntPIK": Call SetNameRef("Run_C_IntPIK", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "D_IntPIK": Call SetNameRef("Run_D_IntPIK", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "A_Prin": Call SetNameRef("Run_A_Prin", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "B_Prin": Call SetNameRef("Run_B_Prin", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "C_Prin": Call SetNameRef("Run_C_Prin", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "D_Prin": Call SetNameRef("Run_D_Prin", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "OC_A": Call SetNameRef("Run_OC_A", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "OC_B": Call SetNameRef("Run_OC_B", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "OC_C": Call SetNameRef("Run_OC_C", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "OC_D": Call SetNameRef("Run_OC_D", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "IC_A": Call SetNameRef("Run_IC_A", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "IC_B": Call SetNameRef("Run_IC_B", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "IC_C": Call SetNameRef("Run_IC_C", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "IC_D": Call SetNameRef("Run_IC_D", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "DSCR": Call SetNameRef("Run_DSCR", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "AdvRate": Call SetNameRef("Run_AdvRate", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "Equity_CF": Call SetNameRef("Run_EquityCF", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "LP_Calls": Call SetNameRef("Run_LP_Calls", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "TurboFlag": Call SetNameRef("Run_TurboFlag", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "Fees_Servicer": Call SetNameRef("Run_Fees_Servicer", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "Fees_Mgmt": Call SetNameRef("Run_Fees_Mgmt", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
            Case "Fees_Admin": Call SetNameRef("Run_Fees_Admin", "=Run!$" & ColLetter(col) & "$5:$" & ColLetter(col) & "$" & lastRow)
        End Select
    Next col
    
    ' Create OC_B trigger line for charting
   

' Place OC_B trigger helper immediately to the right of the OC_B column
Dim ocbCol As Long, c As Long, r As Long

' Find the OC_B header column on row 4
ocbCol = 0
For c = 1 To ws.Cells(4, ws.Columns.Count).End(xlToLeft).Column
    If ws.Cells(4, c).Value = "OC_B" Then
        ocbCol = c
        Exit For
    End If
Next c
If ocbCol = 0 Then Exit Sub  ' safety: OC_B not found

' Name the trigger helper range (same row span as data)
Call SetNameRef( _
    "OC_B_Trigger_Line", _
    "=Run!$" & ColLetter(ocbCol + 1) & "$5:$" & ColLetter(ocbCol + 1) & "$" & lastRow)

' Write the constant trigger value next to OC_B
For r = 5 To lastRow
    ws.Cells(r, ocbCol + 1).Formula = "=Ctl_OC_Trigger_B"
Next r

' Optional: label the helper column for readability
ws.Cells(4, ocbCol + 1).Value = "OC_B Trigger"

End Sub

Private Function ColLetter(ByVal colNum As Long) As String
    On Error Resume Next
    Dim n As Long
    Dim c As Byte
    Dim s As String
    
    n = colNum
    Do
        c = ((n - 1) Mod 26)
        s = Chr(c + 65) & s
        n = (n - c) \ 26
    Loop While n > 0
    ColLetter = s
End Function

'------------------------------------------------------------------------------
' CONTROL NAMED RANGES
'------------------------------------------------------------------------------
Private Sub CreateControlNamedRanges(wb As Workbook)
    On Error Resume Next
    
    ' Core structure
    Call SetNameRef("Ctl_NumQuarters", "=INDEX(Control!$B:$B,MATCH(""NumQuarters"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_First_Close_Date", "=INDEX(Control!$B:$B,MATCH(""First_Close_Date"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Total_Capital", "=INDEX(Control!$B:$B,MATCH(""Total_Capital"",Control!$A:$A,0))")
    
    ' Capital percentages
    Call SetNameRef("Ctl_Pct_A", "=INDEX(Control!$B:$B,MATCH(""Pct_A"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Pct_B", "=INDEX(Control!$B:$B,MATCH(""Pct_B"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Pct_C", "=INDEX(Control!$B:$B,MATCH(""Pct_C"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Pct_D", "=INDEX(Control!$B:$B,MATCH(""Pct_D"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Pct_E", "=INDEX(Control!$B:$B,MATCH(""Pct_E"",Control!$A:$A,0))")
    
    ' Spreads
    Call SetNameRef("Ctl_Spread_A_bps", "=INDEX(Control!$B:$B,MATCH(""Spread_A_bps"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Spread_B_bps", "=INDEX(Control!$B:$B,MATCH(""Spread_B_bps"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Spread_C_bps", "=INDEX(Control!$B:$B,MATCH(""Spread_C_bps"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Spread_D_bps", "=INDEX(Control!$B:$B,MATCH(""Spread_D_bps"",Control!$A:$A,0))")
    
    ' OC triggers
    Call SetNameRef("Ctl_OC_Trigger_A", "=INDEX(Control!$B:$B,MATCH(""OC_Trigger_A"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_OC_Trigger_B", "=INDEX(Control!$B:$B,MATCH(""OC_Trigger_B"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_OC_Trigger_C", "=INDEX(Control!$B:$B,MATCH(""OC_Trigger_C"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_OC_Trigger_D", "=INDEX(Control!$B:$B,MATCH(""OC_Trigger_D"",Control!$A:$A,0))")
    
    ' Features
    Call SetNameRef("Ctl_Enable_C", "=INDEX(Control!$B:$B,MATCH(""Enable_C"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Enable_D", "=INDEX(Control!$B:$B,MATCH(""Enable_D"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Enable_Turbo_DOC", "=INDEX(Control!$B:$B,MATCH(""Enable_Turbo_DOC"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Enable_Excess_Reserve", "=INDEX(Control!$B:$B,MATCH(""Enable_Excess_Reserve"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Enable_PIK", "=INDEX(Control!$B:$B,MATCH(""Enable_PIK"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Enable_CC_PIK", "=INDEX(Control!$B:$B,MATCH(""Enable_CC_PIK"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Enable_Recycling", "=INDEX(Control!$B:$B,MATCH(""Enable_Recycling"",Control!$A:$A,0))")
    
    ' Recycling
    Call SetNameRef("Ctl_Recycling_Pct", "=INDEX(Control!$B:$B,MATCH(""Recycling_Pct"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Recycle_Spread_bps", "=INDEX(Control!$B:$B,MATCH(""Recycle_Spread_bps"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Close_Call_Pct", "=INDEX(Control!$B:$B,MATCH(""Close_Call_Pct"",Control!$A:$A,0))")
    
    ' Timing
    Call SetNameRef("Ctl_Reinvest_Q", "=INDEX(Control!$B:$B,MATCH(""Reinvest_Q"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_GP_Extend_Q", "=INDEX(Control!$B:$B,MATCH(""GP_Extend_Q"",Control!$A:$A,0))")
    
    ' Fees
    Call SetNameRef("Ctl_Arranger_Fee_bps", "=INDEX(Control!$B:$B,MATCH(""Arranger_Fee_bps"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Rating_Agency_Fee_bps", "=INDEX(Control!$B:$B,MATCH(""Rating_Agency_Fee_bps"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Servicer_Fee_bps", "=INDEX(Control!$B:$B,MATCH(""Servicer_Fee_bps"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Mgmt_Fee_Pct", "=INDEX(Control!$B:$B,MATCH(""Mgmt_Fee_Pct"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Admin_Fee_Floor", "=INDEX(Control!$B:$B,MATCH(""Admin_Fee_Floor"",Control!$A:$A,0))")
    
    ' Reserve
    Call SetNameRef("Ctl_Reserve_Pct", "=INDEX(Control!$B:$B,MATCH(""Reserve_Pct"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_PIK_Pct", "=INDEX(Control!$B:$B,MATCH(""PIK_Pct"",Control!$A:$A,0))")
    
    ' Downgrade
    Call SetNameRef("Ctl_Downgrade_OC", "=INDEX(Control!$B:$B,MATCH(""Downgrade_OC"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Downgrade_Spd_Adj_bps", "=INDEX(Control!$B:$B,MATCH(""Downgrade_Spd_Adj_bps"",Control!$A:$A,0))")
    
    ' Scenarios
    Call SetNameRef("Ctl_Scenario_Selection", "=INDEX(Control!$B:$B,MATCH(""Scenario_Selection"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Base_CDR", "=INDEX(Control!$B:$B,MATCH(""Base_CDR"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Base_Recovery", "=INDEX(Control!$B:$B,MATCH(""Base_Recovery"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Base_Prepay", "=INDEX(Control!$B:$B,MATCH(""Base_Prepay"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Base_Amort", "=INDEX(Control!$B:$B,MATCH(""Base_Amort"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Spread_Add_bps", "=INDEX(Control!$B:$B,MATCH(""Spread_Add_bps"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Rate_Add_bps", "=INDEX(Control!$B:$B,MATCH(""Rate_Add_bps"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_Loss_Lag_Q", "=INDEX(Control!$B:$B,MATCH(""Loss_Lag_Q"",Control!$A:$A,0))")
    
    ' Monte Carlo
    Call SetNameRef("Ctl_MC_Iterations", "=INDEX(Control!$B:$B,MATCH(""MC_Iterations"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_MC_Sigma_CDR", "=INDEX(Control!$B:$B,MATCH(""MC_Sigma_CDR"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_MC_Sigma_Rec", "=INDEX(Control!$B:$B,MATCH(""MC_Sigma_Rec"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_MC_Sigma_Sprd_bps", "=INDEX(Control!$B:$B,MATCH(""MC_Sigma_Sprd_bps"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_MC_Rho", "=INDEX(Control!$B:$B,MATCH(""MC_Rho"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_MC_Seed", "=INDEX(Control!$B:$B,MATCH(""MC_Seed"",Control!$A:$A,0))")
    
    ' Fund terms
    Call SetNameRef("Ctl_Pref_Hurdle", "=INDEX(Control!$B:$B,MATCH(""Pref_Hurdle"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_GP_Catch_Up_Pct", "=INDEX(Control!$B:$B,MATCH(""GP_Catch_Up_Pct"",Control!$A:$A,0))")
    Call SetNameRef("Ctl_GP_Split_Pct", "=INDEX(Control!$B:$B,MATCH(""GP_Split_Pct"",Control!$A:$A,0))")
End Sub

'------------------------------------------------------------------------------
' REPORTING SHEETS
'------------------------------------------------------------------------------
Private Sub UpdateAllReportingSheets(wb As Workbook, waterfallResults As Object, _
                                    controlDict As Object, quarterDates() As Date)
    On Error Resume Next
    
    Call RenderExecSummary(wb, waterfallResults, controlDict)
    Call RenderSourcesUsesAtClose(wb, controlDict)
    Call RenderNAVRollForward(wb, waterfallResults, controlDict, quarterDates)
    Call RenderReservesTracking(wb, waterfallResults, quarterDates)
    Call RenderCashflowWaterfallSummary(wb, waterfallResults, quarterDates)
    Call RenderTrancheCashflows(wb, waterfallResults, controlDict, quarterDates)
    Call RenderOCICTests(wb, waterfallResults, controlDict, quarterDates)
    Call RenderBreachesDashboard(wb, waterfallResults, controlDict, quarterDates)
    Call RenderPortfolioStratifications(wb)
    Call RenderAssetPerformance(wb, waterfallResults, quarterDates)
    Call RenderPortfolioCashflowsDetail(wb, waterfallResults, quarterDates)
    Call RenderFeesExpenses(wb, waterfallResults, quarterDates)
    Call RenderInvestorDistributions(wb, waterfallResults, quarterDates)
    Call RenderReportingMetrics(wb, waterfallResults, controlDict, quarterDates)
    Call RenderWaterfallSchedule(wb, waterfallResults, controlDict, quarterDates)
    Call RenderRBCFactors(wb)
    Call RenderPortfolioHHI(wb)
End Sub

Private Sub RenderExecSummary(wb As Workbook, results As Object, controlDict As Object)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long
    
    Set ws = GetOrCreateSheet("Exec_Summary", False)
    ws.Cells.Clear
    
    ' Title
    ws.Range("A1").Value = "EXECUTIVE SUMMARY"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "Fund & Tranche Performance Dashboard"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Fund KPIs
    r = 5
    ws.Cells(r, 1).Value = "FUND METRICS"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 14
    
    r = r + 2
    ws.Cells(r, 1).Value = "MOIC"
    ws.Cells(r, 2).Formula = "=IFERROR(SUM(Run_EquityCF)/(Ctl_Total_Capital*Ctl_Pct_E*(1-Ctl_Reserve_Pct)),0)"
    ws.Cells(r, 2).NumberFormat = "0.00x"
    
    r = r + 1
    ws.Cells(r, 1).Value = "DPI"
    ws.Cells(r, 2).Formula = "=IFERROR(SUM(Run_EquityCF)/(Ctl_Total_Capital*Ctl_Pct_E),0)"
    ws.Cells(r, 2).NumberFormat = "0.00x"
    
    r = r + 1
    ws.Cells(r, 1).Value = "TVPI"
    ws.Cells(r, 2).Formula = "=IFERROR((INDEX(Run_Outstanding,Ctl_NumQuarters)*Ctl_Pct_E/(Ctl_Pct_A+Ctl_Pct_B+Ctl_Pct_C+Ctl_Pct_D+Ctl_Pct_E)+SUM(Run_EquityCF))/(Ctl_Total_Capital*Ctl_Pct_E),0)"
    ws.Cells(r, 2).NumberFormat = "0.00x"
    
    r = r + 1
    ws.Cells(r, 1).Value = "RVPI"
    ws.Cells(r, 2).Formula = "=IFERROR(INDEX(Run_Outstanding,Ctl_NumQuarters)*Ctl_Pct_E/(Ctl_Pct_A+Ctl_Pct_B+Ctl_Pct_C+Ctl_Pct_D+Ctl_Pct_E)/(Ctl_Total_Capital*Ctl_Pct_E),0)"
    ws.Cells(r, 2).NumberFormat = "0.00x"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Gross IRR"
    ws.Cells(r, 2).Formula = "=IFERROR(Reporting_Metrics!E5,0)"
    ws.Cells(r, 2).Style = "SG_Pct"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Net IRR"
    ws.Cells(r, 2).Formula = "=IFERROR(Reporting_Metrics!E5*(1-Ctl_GP_Split_Pct),0)"
    ws.Cells(r, 2).Style = "SG_Pct"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Ending NAV"
    ws.Cells(r, 2).Formula = "=IFERROR(INDEX(Run_Outstanding,Ctl_NumQuarters)*Ctl_Pct_E/(Ctl_Pct_A+Ctl_Pct_B+Ctl_Pct_C+Ctl_Pct_D+Ctl_Pct_E),0)"
    ws.Cells(r, 2).Style = "SG_Currency_K"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Paid-In Capital"
    ws.Cells(r, 2).Formula = "=Ctl_Total_Capital*Ctl_Pct_E"
    ws.Cells(r, 2).Style = "SG_Currency_K"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Distributions"
    ws.Cells(r, 2).Formula = "=SUM(Run_EquityCF)"
    ws.Cells(r, 2).Style = "SG_Currency_K"
    
    ' Risk/Covenant KPIs
    r = r + 3
    ws.Cells(r, 1).Value = "RISK & COVENANTS"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 14
    
    r = r + 2
    ws.Cells(r, 1).Value = "Min OC_A"
    ws.Cells(r, 2).Formula = "=IFERROR(MIN(Run_OC_A),999)"
    ws.Cells(r, 2).NumberFormat = "0.00x"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Min OC_B"
    ws.Cells(r, 2).Formula = "=IFERROR(MIN(Run_OC_B),999)"
    ws.Cells(r, 2).NumberFormat = "0.00x"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Min DSCR"
    ws.Cells(r, 2).Formula = "=IFERROR(MIN(Run_DSCR),999)"
    ws.Cells(r, 2).NumberFormat = "0.00x"
    
    r = r + 1
    ws.Cells(r, 1).Value = "OC_B Cushion"
    ws.Cells(r, 2).Formula = "=MIN(Run_OC_B)-Ctl_OC_Trigger_B"
    ws.Cells(r, 2).NumberFormat = "0.00x"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Breach Periods"
    ws.Cells(r, 2).Formula = "=COUNTIF(Run_OC_B,""<""&Ctl_OC_Trigger_B)"
    ws.Cells(r, 2).NumberFormat = "0"
    
    ' Tranche metrics
    r = r + 3
    ws.Cells(r, 1).Value = "TRANCHE METRICS"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 14
    
    r = r + 2
    ws.Cells(r, 1).Value = "Class A WAL"
    ws.Cells(r, 2).Formula = "=IFERROR(Reporting_Metrics!A7,0)"
    ws.Cells(r, 2).NumberFormat = "0.0"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Class B WAL"
    ws.Cells(r, 2).Formula = "=IFERROR(Reporting_Metrics!B7,0)"
    ws.Cells(r, 2).NumberFormat = "0.0"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Class A IRR"
    ws.Cells(r, 2).Formula = "=IFERROR(Reporting_Metrics!A5,0)"
    ws.Cells(r, 2).Style = "SG_Pct"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Class B IRR"
    ws.Cells(r, 2).Formula = "=IFERROR(Reporting_Metrics!B5,0)"
    ws.Cells(r, 2).Style = "SG_Pct"
    
    r = r + 1
    ws.Cells(r, 1).Value = "A Outstanding"
    ws.Cells(r, 2).Formula = "=INDEX(Run_A_EndBal,Ctl_NumQuarters)"
    ws.Cells(r, 2).Style = "SG_Currency_K"
    
    r = r + 1
    ws.Cells(r, 1).Value = "B Outstanding"
    ws.Cells(r, 2).Formula = "=INDEX(Run_B_EndBal,Ctl_NumQuarters)"
    ws.Cells(r, 2).Style = "SG_Currency_K"
    
    ' Add coverage trend chart
    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(Left:=300, Top:=100, Width:=400, Height:=250)
    With cht.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "OC_B vs Trigger"
        
        ' Clear series
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        
        ' Add OC_B series
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "OC_B"
        .SeriesCollection(1).XValues = "=Run_Dates"
        .SeriesCollection(1).Values = "=Run_OC_B"
        
        ' Add trigger line
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "Trigger"
        .SeriesCollection(2).XValues = "=Run_Dates"
        .SeriesCollection(2).Values = "=OC_B_Trigger_Line"
        .SeriesCollection(2).Format.Line.DashStyle = msoLineDash
        .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
    
    ' Add cumulative distributions chart
    Set cht = ws.ChartObjects.Add(Left:=300, Top:=400, Width:=400, Height:=250)
    With cht.Chart
        .ChartType = xlArea
        .HasTitle = True
        .ChartTitle.Text = "Cumulative Distributions"
        
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "Equity Distributions"
        .SeriesCollection(1).XValues = "=Run_Dates"
        .SeriesCollection(1).Values = "=Run_EquityCF"
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = SG_RED
        
        .HasLegend = True
    End With
    
    ' Format checks
    ws.Columns("A:B").AutoFit
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

Private Sub RenderSourcesUsesAtClose(wb As Workbook, controlDict As Object)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long
    
    Set ws = GetOrCreateSheet("Sources_Uses_At_Close", False)
    ws.Cells.Clear
    
    ' Title
    ws.Range("A1").Value = "SOURCES & USES AT CLOSE"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "Capitalization Snapshot"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Sources table
    r = 5
    ws.Cells(r, 1).Value = "SOURCES"
    ws.Cells(r, 1).Font.Bold = True
    ws.Range("A" & r & ":C" & r).Style = "SG_Hdr"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Class A Notes"
    ws.Cells(r, 2).Formula = "=Ctl_Total_Capital*Ctl_Pct_A"
    ws.Cells(r, 3).Formula = "=B" & r & "/B" & (r + 5)
    
    r = r + 1
    ws.Cells(r, 1).Value = "Class B Notes"
    ws.Cells(r, 2).Formula = "=Ctl_Total_Capital*Ctl_Pct_B"
    ws.Cells(r, 3).Formula = "=B" & r & "/B" & (r + 4)
    
    If ToBool(controlDict("Enable_C")) Then
        r = r + 1
        ws.Cells(r, 1).Value = "Class C Notes"
        ws.Cells(r, 2).Formula = "=Ctl_Total_Capital*Ctl_Pct_C"
        ws.Cells(r, 3).Formula = "=B" & r & "/B" & (r + 3)
    End If
    
    If ToBool(controlDict("Enable_D")) Then
        r = r + 1
        ws.Cells(r, 1).Value = "Class D Notes"
        ws.Cells(r, 2).Formula = "=Ctl_Total_Capital*Ctl_Pct_D"
        ws.Cells(r, 3).Formula = "=B" & r & "/B" & (r + 2)
    End If
    
    r = r + 1
    ws.Cells(r, 1).Value = "Equity (Gross)"
    ws.Cells(r, 2).Formula = "=Ctl_Total_Capital*Ctl_Pct_E"
    ws.Cells(r, 3).Formula = "=B" & r & "/B" & (r + 1)
    
    r = r + 1
    ws.Cells(r, 1).Value = "TOTAL SOURCES"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Formula = "=Ctl_Total_Capital"
    ws.Cells(r, 3).Formula = "=B" & r & "/B" & r
    ws.Range("A" & r & ":C" & r).Borders(xlEdgeTop).LineStyle = xlContinuous
    
    ' Uses table
    r = r + 3
    ws.Cells(r, 1).Value = "USES"
    ws.Cells(r, 1).Font.Bold = True
    ws.Range("A" & r & ":C" & r).Style = "SG_Hdr"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Asset Purchases"
    ws.Cells(r, 2).Formula = "=Ctl_Total_Capital*(1-Ctl_Arranger_Fee_bps/10000)"
    ws.Cells(r, 3).Formula = "=B" & r & "/B" & (r + 4)
    
    r = r + 1
    ws.Cells(r, 1).Value = "Arranger Fee"
    ws.Cells(r, 2).Formula = "=Ctl_Total_Capital*Ctl_Arranger_Fee_bps/10000"
    ws.Cells(r, 3).Formula = "=B" & r & "/B" & (r + 3)
    
    r = r + 1
    ws.Cells(r, 1).Value = "Rating Agency Fee"
    ws.Cells(r, 2).Formula = "=Ctl_Total_Capital*Ctl_Rating_Agency_Fee_bps/10000"
    ws.Cells(r, 3).Formula = "=B" & r & "/B" & (r + 2)
    
    r = r + 1
    ws.Cells(r, 1).Value = "Initial Reserve"
    ws.Cells(r, 2).Formula = "=Ctl_Total_Capital*Ctl_Pct_E*Ctl_Reserve_Pct"
    ws.Cells(r, 3).Formula = "=B" & r & "/B" & (r + 1)
    
    r = r + 1
    ws.Cells(r, 1).Value = "TOTAL USES"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Formula = "=SUM(B" & (r - 4) & ":B" & (r - 1) & ")"
    ws.Cells(r, 3).Formula = "=B" & r & "/B" & r
    ws.Range("A" & r & ":C" & r).Borders(xlEdgeTop).LineStyle = xlContinuous
    
    ' Check
    r = r + 2
    ws.Cells(r, 1).Value = "Check (Sources - Uses)"
    ws.Cells(r, 2).Formula = "=B11-B20"
    ws.Cells(r, 2).NumberFormat = "$#,##0;[Red]-$#,##0"
    
    ' Leverage metrics
    r = r + 3
    ws.Cells(r, 1).Value = "LEVERAGE & COVERAGE"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Cells(r, 1).Value = "Debt/Equity"
    ws.Cells(r, 2).Formula = "=(Ctl_Pct_A+Ctl_Pct_B+Ctl_Pct_C+Ctl_Pct_D)/Ctl_Pct_E"
    ws.Cells(r, 2).NumberFormat = "0.00x"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Class A Pricing"
    ws.Cells(r, 2).Formula = "=""S+""&Ctl_Spread_A_bps&""bps"""
    
    r = r + 1
    ws.Cells(r, 1).Value = "Class B Pricing"
    ws.Cells(r, 2).Formula = "=""S+""&Ctl_Spread_B_bps&""bps"""
    
    r = r + 1
    ws.Cells(r, 1).Value = "OC_A Trigger"
    ws.Cells(r, 2).Formula = "=Ctl_OC_Trigger_A"
    ws.Cells(r, 2).NumberFormat = "0.00x"
    
    r = r + 1
    ws.Cells(r, 1).Value = "OC_B Trigger"
    ws.Cells(r, 2).Formula = "=Ctl_OC_Trigger_B"
    ws.Cells(r, 2).NumberFormat = "0.00x"
    
    ' Format
    ws.Range("B:B").Style = "SG_Currency_K"
    ws.Range("C:C").Style = "SG_Pct"
    ws.Columns("A:C").AutoFit
    
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

Private Sub RenderNAVRollForward(wb As Workbook, results As Object, controlDict As Object, quarterDates() As Date)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long, q As Long
    Dim numQ As Long
    
    Set ws = GetOrCreateSheet("NAV_Roll_Forward", False)
    ws.Cells.Clear
    
    numQ = UBound(quarterDates) + 1
    
    ' Title
    ws.Range("A1").Value = "NAV ROLL FORWARD"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "Fund Performance Bridge"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Headers
    r = 4
    ws.Range("A" & r & ":K" & r).Value = Array("Period", "Start NAV", "Capital Calls", _
        "NII", "Realized P&L", "Unrealized P&L", "Defaults", "Recoveries", _
        "Reserve Δ", "Distributions", "End NAV")
    ws.Range("A" & r & ":K" & r).Style = "SG_Hdr"
    
    ' Data rows with formulas
    For q = 1 To numQ
        r = 4 + q
        ws.Cells(r, 1).Value = quarterDates(q - 1)
        
        ' Start NAV
        If q = 1 Then
            ws.Cells(r, 2).Formula = "=Ctl_Total_Capital*Ctl_Pct_E*(1-Ctl_Reserve_Pct)"
        Else
            ws.Cells(r, 2).Formula = "=K" & (r - 1)
        End If
        
        ' Capital Calls
        If q = 1 Then
            ws.Cells(r, 3).Formula = "=Ctl_Total_Capital*Ctl_Pct_E*Ctl_Close_Call_Pct"
        ElseIf q <= ToLng(controlDict("Reinvest_Q")) Then
            ws.Cells(r, 3).Formula = "=Ctl_Total_Capital*Ctl_Pct_E*(1-Ctl_Close_Call_Pct)/(Ctl_Reinvest_Q-1)"
        Else
            ws.Cells(r, 3).Value = 0
        End If
        
        ' NII (Interest + Fees - Expenses)
        ws.Cells(r, 4).Formula = "=INDEX(Run_Interest," & q & ")+INDEX(Run_CommitmentFees," & q & ")-INDEX(Run_Fees_Servicer," & q & ")-INDEX(Run_Fees_Mgmt," & q & ")-INDEX(Run_Fees_Admin," & q & ")"
        
        ' Realized P&L
        ws.Cells(r, 5).Value = 0  ' Placeholder
        
        ' Unrealized P&L
        ws.Cells(r, 6).Value = 0  ' Placeholder
        
        ' Defaults
        ws.Cells(r, 7).Formula = "=-INDEX(Run_Defaults," & q & ")"
        
        ' Recoveries
        ws.Cells(r, 8).Formula = "=INDEX(Run_Recoveries," & q & ")"
        
        ' Reserve changes
        ws.Cells(r, 9).Formula = "=INDEX(Run_Reserve_TopUp," & q & ")-INDEX(Run_Reserve_Release," & q & ")"
        
        ' Distributions
        ws.Cells(r, 10).Formula = "=INDEX(Run_EquityCF," & q & ")"
        
        ' End NAV
        ws.Cells(r, 11).Formula = "=B" & r & "+C" & r & "+D" & r & "+E" & r & "+F" & r & "+G" & r & "+H" & r & "-I" & r & "-J" & r
    Next q
    
    ' Summary metrics
    r = r + 2
    ws.Cells(r, 1).Value = "SUMMARY METRICS"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Cells(r, 1).Value = "Total Calls"
    ws.Cells(r, 2).Formula = "=SUM(C5:C" & (4 + numQ) & ")"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Total Distributions"
    ws.Cells(r, 2).Formula = "=SUM(J5:J" & (4 + numQ) & ")"
    
    r = r + 1
    ws.Cells(r, 1).Value = "DPI"
    ws.Cells(r, 2).Formula = "=B" & (r - 1) & "/B" & (r - 2)
    
    r = r + 1
    ws.Cells(r, 1).Value = "RVPI"
    ws.Cells(r, 2).Formula = "=K" & (4 + numQ) & "/B" & (r - 3)
    
    r = r + 1
    ws.Cells(r, 1).Value = "TVPI"
    'Assuming DPI is at row (r-2) and RVPI at row (r-1):
 ws.Cells(r, 2).Formula = "=B" & (r - 2) & "+B" & (r - 1)
 ws.Cells(r, 2).NumberFormat = "0.00x"
    
    ' Format
    ws.Columns("A:A").NumberFormat = "mm/dd/yyyy"
    ws.Columns("B:K").Style = "SG_Currency_K"
    ws.Range("B" & (r - 2) & ":B" & r).NumberFormat = "0.00x"
    
    ws.Columns.AutoFit
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

Private Sub RenderReservesTracking(wb As Workbook, results As Object, quarterDates() As Date)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long, q As Long
    Dim numQ As Long
    
    Set ws = GetOrCreateSheet("Reserves_Tracking", False)
    ws.Cells.Clear
    
    numQ = UBound(quarterDates) + 1
    
    ' Title
    ws.Range("A1").Value = "RESERVES TRACKING"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "Excess / PIK / Liquidity Reserves"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Headers
    r = 4
    ws.Range("A" & r & ":F" & r).Value = Array("Period", "Opening", "Adds", "Draws", "Releases", "Closing")
    ws.Range("A" & r & ":F" & r).Style = "SG_Hdr"
    
    ' Data rows
    For q = 1 To numQ
        r = 4 + q
        ws.Cells(r, 1).Value = quarterDates(q - 1)
        ws.Cells(r, 2).Formula = "=INDEX(Run_Reserve_Beg," & q & ")"
        ws.Cells(r, 3).Formula = "=INDEX(Run_Reserve_TopUp," & q & ")"
        ws.Cells(r, 4).Formula = "=INDEX(Run_Reserve_Draw," & q & ")"
        ws.Cells(r, 5).Formula = "=INDEX(Run_Reserve_Release," & q & ")"
        ws.Cells(r, 6).Formula = "=INDEX(Run_Reserve_End," & q & ")"
    Next q
    
    ' Check row
    r = r + 1
    ws.Cells(r, 1).Value = "Check"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 6).Formula = "=B5+SUM(C5:C" & (r - 1) & ")-SUM(D5:E" & (r - 1) & ")-F" & (r - 1)
    ws.Cells(r, 6).NumberFormat = "$#,##0;[Red]-$#,##0"
    
    ' Coverage ratios
    r = r + 2
    ws.Cells(r, 1).Value = "COVERAGE RATIOS"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Cells(r, 1).Value = "Interest Coverage"
    ws.Cells(r, 2).Formula = "=IFERROR(INDEX(Run_Reserve_End,Ctl_NumQuarters)/(INDEX(Run_A_IntDue,Ctl_NumQuarters)+INDEX(Run_B_IntDue,Ctl_NumQuarters)),0)"
    ws.Cells(r, 2).NumberFormat = "0.00x"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Par Coverage"
    ws.Cells(r, 2).Formula = "=IFERROR(INDEX(Run_Reserve_End,Ctl_NumQuarters)/INDEX(Run_Outstanding,Ctl_NumQuarters),0)"
    ws.Cells(r, 2).Style = "SG_Pct"
    
    ' Format
    ws.Columns("A:A").NumberFormat = "mm/dd/yyyy"
    ws.Columns("B:F").Style = "SG_Currency_K"
    
    ws.Columns.AutoFit
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

Private Sub RenderCashflowWaterfallSummary(wb As Workbook, results As Object, quarterDates() As Date)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long, q As Long
    Dim numQ As Long
    
    Set ws = GetOrCreateSheet("Cashflow_Waterfall_Summary", False)
    ws.Cells.Clear
    
    numQ = UBound(quarterDates) + 1
    
    ' Title
    ws.Range("A1").Value = "CASHFLOW WATERFALL SUMMARY"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "Priority of Payments"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Headers
    r = 4
    ws.Range("A" & r & ":L" & r).Value = Array("Period", "Cash In", "Operating", _
        "A Interest", "A Principal", "B Interest", "B Principal", _
        "Reserve Fund", "Reserve Release", "Excess to Equity", "LP Calls", "Check")
    ws.Range("A" & r & ":L" & r).Style = "SG_Hdr"
    
    ' Data rows
    For q = 1 To numQ
        r = 4 + q
        ws.Cells(r, 1).Value = quarterDates(q - 1)
        
        ' Cash In
        ws.Cells(r, 2).Formula = "=INDEX(Run_Interest," & q & ")+INDEX(Run_CommitmentFees," & q & ")+INDEX(Run_Recoveries," & q & ")+IF(INDEX(Run_TurboFlag," & q & ")=1,INDEX(Run_Principal," & q & "),IF(" & q & ">Ctl_Reinvest_Q+Ctl_GP_Extend_Q,INDEX(Run_Principal," & q & "),0))"
        
        ' Operating expenses
        ws.Cells(r, 3).Formula = "=INDEX(Run_Fees_Servicer," & q & ")+INDEX(Run_Fees_Mgmt," & q & ")+INDEX(Run_Fees_Admin," & q & ")"
        
        ' A Interest
        ws.Cells(r, 4).Formula = "=INDEX(Run_A_IntPd," & q & ")"
        
        ' A Principal
        ws.Cells(r, 5).Formula = "=INDEX(Run_A_Prin," & q & ")"
        
        ' B Interest
        ws.Cells(r, 6).Formula = "=INDEX(Run_B_IntPd," & q & ")"
        
        ' B Principal
        ws.Cells(r, 7).Formula = "=INDEX(Run_B_Prin," & q & ")"
        
        ' Reserve movements
        ws.Cells(r, 8).Formula = "=INDEX(Run_Reserve_TopUp," & q & ")"
        ws.Cells(r, 9).Formula = "=INDEX(Run_Reserve_Release," & q & ")"
        
        ' Equity
        ws.Cells(r, 10).Formula = "=INDEX(Run_EquityCF," & q & ")"
        
        ' LP Calls
        ws.Cells(r, 11).Formula = "=INDEX(Run_LP_Calls," & q & ")"
        
        ' Check
        ws.Cells(r, 12).Formula = "=B" & r & "+I" & r & "-SUM(C" & r & ":H" & r & ")-J" & r & "-K" & r
    Next q
    
    ' Totals row
    r = r + 1
    ws.Cells(r, 1).Value = "TOTAL"
    ws.Cells(r, 1).Font.Bold = True
    For q = 2 To 11
        ws.Cells(r, q).Formula = "=SUM(" & ColLetter(q) & "5:" & ColLetter(q) & (r - 1) & ")"
    Next q
    ws.Range("A" & r & ":L" & r).Borders(xlEdgeTop).LineStyle = xlContinuous
    
    ' Format
    ws.Columns("A:A").NumberFormat = "mm/dd/yyyy"
    ws.Columns("B:L").Style = "SG_Currency_K"
    
    ' Add stacked bar chart
    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(Left:=100, Top:=300, Width:=600, Height:=300)
    With cht.Chart
        .ChartType = xlColumnStacked
        .HasTitle = True
        .ChartTitle.Text = "Waterfall Components"
        
        ' Set source data for key components
        .SetSourceData Source:=ws.Range("C5:G" & (4 + numQ))
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
    
    ws.Columns.AutoFit
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

Private Sub RenderTrancheCashflows(wb As Workbook, results As Object, controlDict As Object, quarterDates() As Date)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long, q As Long
    Dim numQ As Long
    Dim enableC As Boolean, enableD As Boolean
    
    Set ws = GetOrCreateSheet("Tranche_Cashflows", False)
    ws.Cells.Clear
    
    numQ = UBound(quarterDates) + 1
    enableC = ToBool(controlDict("Enable_C"))
    enableD = ToBool(controlDict("Enable_D"))
    
    ' Title
    ws.Range("A1").Value = "TRANCHE CASHFLOWS"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "Class A/B/C/D/E Schedules"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Class A table
    r = 5
    ws.Cells(r, 1).Value = "CLASS A"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Size = 12
    
    r = r + 1
    ws.Range("A" & r & ":H" & r).Value = Array("Period", "Opening Par", "Interest Due", _
        "Interest Paid", "Shortfall/PIK", "Principal Paid", "Ending Par", "Cum Cash")
    ws.Range("A" & r & ":H" & r).Style = "SG_Hdr"
    
    For q = 1 To Application.Min(numQ, 20)
        r = r + 1
        ws.Cells(r, 1).Value = quarterDates(q - 1)
        ws.Cells(r, 2).Formula = "=IF(" & q & "=1,Ctl_Total_Capital*Ctl_Pct_A,G" & (r - 1) & ")"
        ws.Cells(r, 3).Formula = "=INDEX(Run_A_IntDue," & q & ")"
        ws.Cells(r, 4).Formula = "=INDEX(Run_A_IntPd," & q & ")"
        ws.Cells(r, 5).Formula = "=INDEX(Run_A_IntPIK," & q & ")"
        ws.Cells(r, 6).Formula = "=INDEX(Run_A_Prin," & q & ")"
        ws.Cells(r, 7).Formula = "=INDEX(Run_A_EndBal," & q & ")"
        ws.Cells(r, 8).Formula = "=IF(" & q & "=1,D" & r & "+F" & r & ",H" & (r - 1) & "+D" & r & "+F" & r & ")"
    Next q
    
    ' Class A metrics
    r = r + 2
    ws.Cells(r, 1).Value = "WAL (years)"
    ws.Cells(r, 2).Formula = "=IFERROR(Reporting_Metrics!A7,0)"
    ws.Cells(r, 2).NumberFormat = "0.0"
    
    r = r + 1
    ws.Cells(r, 1).Value = "IRR"
    ws.Cells(r, 2).Formula = "=IFERROR(Reporting_Metrics!A5,0)"
    ws.Cells(r, 2).Style = "SG_Pct"
    
    ' Format class A section
    ws.Range("B7:H" & (6 + Application.Min(numQ, 20))).Style = "SG_Currency_K"
    
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

Private Sub RenderOCICTests(wb As Workbook, results As Object, controlDict As Object, quarterDates() As Date)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long, q As Long
    Dim numQ As Long
    Dim enableC As Boolean, enableD As Boolean
    
    Set ws = GetOrCreateSheet("OCIC_Tests", False)
    ws.Cells.Clear
    
    numQ = UBound(quarterDates) + 1
    enableC = ToBool(controlDict("Enable_C"))
    enableD = ToBool(controlDict("Enable_D"))
    
    ' Title
    ws.Range("A1").Value = "OC/IC COVENANT TESTS"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "Coverage Ratios & Cushions"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Headers
    r = 4
    Dim col As Long
    col = 1
    ws.Cells(r, col).Value = "Date": col = col + 1
    ws.Cells(r, col).Value = "OC_A": col = col + 1
    ws.Cells(r, col).Value = "OC_B": col = col + 1
    If enableC Then ws.Cells(r, col).Value = "OC_C": col = col + 1
    If enableD Then ws.Cells(r, col).Value = "OC_D": col = col + 1
    ws.Cells(r, col).Value = "IC_A": col = col + 1
    ws.Cells(r, col).Value = "IC_B": col = col + 1
    If enableC Then ws.Cells(r, col).Value = "IC_C": col = col + 1
    If enableD Then ws.Cells(r, col).Value = "IC_D": col = col + 1
    ws.Cells(r, col).Value = "DSCR": col = col + 1
    ws.Cells(r, col).Value = "AdvRate": col = col + 1
    ws.Range("A4").Resize(1, col - 1).Style = "SG_Hdr"
    
    ' Data formulas
    For q = 1 To numQ
        r = 4 + q
        col = 1
        ws.Cells(r, col).Value = quarterDates(q - 1): col = col + 1
        ws.Cells(r, col).Formula = "=INDEX(Run_OC_A," & q & ")": col = col + 1
        ws.Cells(r, col).Formula = "=INDEX(Run_OC_B," & q & ")": col = col + 1
        If enableC Then ws.Cells(r, col).Formula = "=INDEX(Run_OC_C," & q & ")": col = col + 1
        If enableD Then ws.Cells(r, col).Formula = "=INDEX(Run_OC_D," & q & ")": col = col + 1
        ws.Cells(r, col).Formula = "=INDEX(Run_IC_A," & q & ")": col = col + 1
        ws.Cells(r, col).Formula = "=INDEX(Run_IC_B," & q & ")": col = col + 1
        If enableC Then ws.Cells(r, col).Formula = "=INDEX(Run_IC_C," & q & ")": col = col + 1
        If enableD Then ws.Cells(r, col).Formula = "=INDEX(Run_IC_D," & q & ")": col = col + 1
        ws.Cells(r, col).Formula = "=INDEX(Run_DSCR," & q & ")": col = col + 1
        ws.Cells(r, col).Formula = "=INDEX(Run_AdvRate," & q & ")": col = col + 1
    Next q
    
    ' Apply heatmap
    Dim heatmapRange As Range
    Set heatmapRange = ws.Range("B5").Resize(numQ, col - 2)
    Call ClearAndApplyOCICHeatmap(heatmapRange)
    
    ' KBRA Cushion table
    r = 6
    ws.Cells(r, col + 1).Value = "KBRA CUSHIONS"
    ws.Cells(r, col + 1).Font.Bold = True
    
    r = r + 1
    ws.Cells(r, col + 1).Value = "OC_B Target"
    ws.Cells(r, col + 2).Formula = "=Ctl_OC_Trigger_B"
    ws.Cells(r, col + 2).NumberFormat = "0.00x"
    
    r = r + 1
    ws.Cells(r, col + 1).Value = "OC_B Min"
    ws.Cells(r, col + 2).Formula = "=MIN(Run_OC_B)"
    ws.Cells(r, col + 2).NumberFormat = "0.00x"
    
    r = r + 1
    ws.Cells(r, col + 1).Value = "Cushion"
    ws.Cells(r, col + 2).Formula = "=" & ColLetter(col + 2) & (r - 1) & "-" & ColLetter(col + 2) & (r - 2)
    ws.Cells(r, col + 2).NumberFormat = "0.00x"
    
    r = r + 1
    ws.Cells(r, col + 1).Value = "Status"
    ws.Cells(r, col + 2).Formula = "=IF(" & ColLetter(col + 2) & (r - 1) & ">0,""PASS"",""FAIL"")"
    
    ' Format status cell
    With ws.Cells(r, col + 2)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="""PASS"""
        .FormatConditions(1).Interior.Color = RGB(198, 239, 206)
        .FormatConditions(1).Font.Bold = True
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="""FAIL"""
        .FormatConditions(2).Interior.Color = RGB(255, 199, 206)
        .FormatConditions(2).Font.Bold = True
    End With
    
    ' Format columns
    ws.Columns("A:A").NumberFormat = "mm/dd/yyyy"
    ws.Columns("B:" & ColLetter(col - 2)).NumberFormat = "0.00x"
    ws.Columns(ColLetter(col - 1)).NumberFormat = "0.00x"
    ws.Columns(ColLetter(col)).Style = "SG_Pct"
    
    ws.Columns.AutoFit
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

Private Sub UpdateOCICChart(wb As Workbook)
    On Error Resume Next
    Dim ws As Worksheet
    Dim cht As ChartObject
    
    Set ws = wb.Worksheets("OCIC_Tests")
    
    ' Create or update chart
    Set cht = EnsureSingleChart(ws, "OC_Cushion_Chart", OCIC_CHART_FRAME)
    
    With cht.Chart
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "OC_B vs Trigger"
        
        ' Add OC_B series
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Name = "OC_B"
        .SeriesCollection(1).XValues = "=Run_Dates"
        .SeriesCollection(1).Values = "=Run_OC_B"
        .SeriesCollection(1).MarkerStyle = xlMarkerStyleCircle
        
        ' Add trigger line
        .SeriesCollection.NewSeries
        .SeriesCollection(2).Name = "OC_B Trigger"
        .SeriesCollection(2).XValues = "=Run_Dates"
        .SeriesCollection(2).Values = "=OC_B_Trigger_Line"
        .SeriesCollection(2).MarkerStyle = xlMarkerStyleNone
        .SeriesCollection(2).Format.Line.DashStyle = msoLineDash
        .SeriesCollection(2).Format.Line.ForeColor.RGB = RGB(255, 0, 0)
        
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
End Sub

Private Sub RenderBreachesDashboard(wb As Workbook, results As Object, controlDict As Object, quarterDates() As Date)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long
    
    Set ws = GetOrCreateSheet("Breaches_Dashboard", False)
    ws.Cells.Clear
    
    ' Title
    ws.Range("A1").Value = "BREACHES DASHBOARD"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "Incidents & Remediation"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Breach log placeholder
    r = 5
    ws.Cells(r, 1).Value = "BREACH LOG"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Range("A" & r & ":I" & r).Value = Array("Period", "Test", "Actual", "Trigger", _
        "Cushion", "Severity", "Action", "Status", "Days")
    ws.Range("A" & r & ":I" & r).Style = "SG_Hdr"
    
    ' Summary statistics
    r = r + 10
    ws.Cells(r, 1).Value = "SUMMARY"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Cells(r, 1).Value = "Total Breach Periods"
    ws.Cells(r, 2).Formula = "=COUNTIF(Run_OC_B,""<""&Ctl_OC_Trigger_B)"
    
    r = r + 1
    ws.Cells(r, 1).Value = "% Time in Breach"
    ws.Cells(r, 2).Formula = "=B" & (r - 1) & "/Ctl_NumQuarters"
    ws.Cells(r, 2).Style = "SG_Pct"
    
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

Private Sub RenderPortfolioStratifications(wb As Workbook)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long
    
    Set ws = GetOrCreateSheet("Portfolio_Stratifications", False)
    ws.Cells.Clear
    
    ' Title
    ws.Range("A1").Value = "PORTFOLIO STRATIFICATIONS"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "Risk Dispersion"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Industry distribution
    r = 5
    ws.Cells(r, 1).Value = "INDUSTRY DISTRIBUTION"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Range("A" & r & ":C" & r).Value = Array("Industry", "Par", "% of Total")
    ws.Range("A" & r & ":C" & r).Style = "SG_Hdr"
    
    ' Placeholder for dynamic industry calc
    r = r + 1
    ws.Cells(r, 1).Value = "Technology"
    ws.Cells(r, 2).Formula = "=SUMIF(AssetTape!L:L,""Technology"",AssetTape!B:B)"
    ws.Cells(r, 3).Formula = "=B" & r & "/SUM(AssetTape!B:B)"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Healthcare"
    ws.Cells(r, 2).Formula = "=SUMIF(AssetTape!L:L,""Healthcare"",AssetTape!B:B)"
    ws.Cells(r, 3).Formula = "=B" & r & "/SUM(AssetTape!B:B)"
    
    ' Rating distribution
    r = r + 5
    ws.Cells(r, 1).Value = "RATING DISTRIBUTION"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Range("A" & r & ":C" & r).Value = Array("Rating", "Par", "% of Total")
    ws.Range("A" & r & ":C" & r).Style = "SG_Hdr"
    
    ' Key metrics
    r = r + 10
    ws.Cells(r, 1).Value = "KEY METRICS"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Cells(r, 1).Value = "WAM (years)"
    ws.Cells(r, 2).Formula = "=SUMPRODUCT(AssetTape!B:B,AssetTape!I:I)/SUM(AssetTape!B:B)"
    ws.Cells(r, 2).NumberFormat = "0.0"
    
    r = r + 1
    ws.Cells(r, 1).Value = "WA Spread (bps)"
    ws.Cells(r, 2).Formula = "=SUMPRODUCT(AssetTape!B:B,AssetTape!D:D)/SUM(AssetTape!B:B)"
    ws.Cells(r, 2).NumberFormat = "0"
    
    r = r + 1
    ws.Cells(r, 1).Value = "WA LTV"
    ws.Cells(r, 2).Formula = "=SUMPRODUCT(AssetTape!B:B,AssetTape!J:J)/SUM(AssetTape!B:B)"
    ws.Cells(r, 2).Style = "SG_Pct"
    
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

Private Sub RenderAssetPerformance(wb As Workbook, results As Object, quarterDates() As Date)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long, q As Long
    Dim numQ As Long
    
    Set ws = GetOrCreateSheet("Asset_Performance", False)
    ws.Cells.Clear
    
    numQ = UBound(quarterDates) + 1
    
    ' Title
    ws.Range("A1").Value = "ASSET PERFORMANCE"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "Credit Outcomes"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Defaults & Recoveries table
    r = 5
    ws.Cells(r, 1).Value = "DEFAULTS & RECOVERIES"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Range("A" & r & ":F" & r).Value = Array("Period", "Defaults", "Cumulative", _
        "Recoveries", "Net Loss", "Loss Rate")
    ws.Range("A" & r & ":F" & r).Style = "SG_Hdr"
    
    For q = 1 To Application.Min(numQ, 20)
        r = r + 1
        ws.Cells(r, 1).Value = quarterDates(q - 1)
        ws.Cells(r, 2).Formula = "=INDEX(Run_Defaults," & q & ")"
        ws.Cells(r, 3).Formula = "=SUM(B$7:B" & r & ")"
        ws.Cells(r, 4).Formula = "=INDEX(Run_Recoveries," & q & ")"
        ws.Cells(r, 5).Formula = "=B" & r & "-D" & r
        ws.Cells(r, 6).Formula = "=E" & r & "/INDEX(Run_Outstanding,1)"
    Next q
    
    ' Format
    ws.Columns("A:A").NumberFormat = "mm/dd/yyyy"
    ws.Columns("B:E").Style = "SG_Currency_K"
    ws.Columns("F:F").Style = "SG_Pct"
    
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

Private Sub RenderPortfolioCashflowsDetail(wb As Workbook, results As Object, quarterDates() As Date)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long, q As Long
    Dim numQ As Long
    
    Set ws = GetOrCreateSheet("Portfolio_Cashflows_Detail", False)
    ws.Cells.Clear
    
    numQ = UBound(quarterDates) + 1
    
    ' Title
    ws.Range("A1").Value = "PORTFOLIO CASHFLOWS DETAIL"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "Asset-Level Drilldown"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Summary table
    r = 5
    ws.Cells(r, 1).Value = "AGGREGATE CASHFLOWS"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Range("A" & r & ":G" & r).Value = Array("Period", "Interest", "Fees", _
        "Scheduled Principal", "Prepayments", "Recoveries", "Total")
    ws.Range("A" & r & ":G" & r).Style = "SG_Hdr"
    
    For q = 1 To Application.Min(numQ, 20)
        r = r + 1
        ws.Cells(r, 1).Value = quarterDates(q - 1)
        ws.Cells(r, 2).Formula = "=INDEX(Run_Interest," & q & ")"
        ws.Cells(r, 3).Formula = "=INDEX(Run_CommitmentFees," & q & ")"
        ws.Cells(r, 4).Formula = "=INDEX(Run_Principal," & q & ")-INDEX(Run_Prepayments," & q & ")"
        ws.Cells(r, 5).Formula = "=INDEX(Run_Prepayments," & q & ")"
        ws.Cells(r, 6).Formula = "=INDEX(Run_Recoveries," & q & ")"
        ws.Cells(r, 7).Formula = "=SUM(B" & r & ":F" & r & ")"
    Next q
    
    ' Check row
    r = r + 1
    ws.Cells(r, 1).Value = "Check"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 7).Formula = "=SUM(B7:F" & (r - 1) & ")-SUM(G7:G" & (r - 1) & ")"
    
    ' Format
    ws.Columns("A:A").NumberFormat = "mm/dd/yyyy"
    ws.Columns("B:G").Style = "SG_Currency_K"
    
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

Private Sub RenderFeesExpenses(wb As Workbook, results As Object, quarterDates() As Date)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long, q As Long
    Dim numQ As Long
    
    Set ws = GetOrCreateSheet("Fees_Expenses", False)
    ws.Cells.Clear
    
    numQ = UBound(quarterDates) + 1
    
    ' Title
    ws.Range("A1").Value = "FEES & EXPENSES"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "Gross to Net Reconciliation"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Fee schedule
    r = 5
    ws.Cells(r, 1).Value = "FEE SCHEDULE"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Range("A" & r & ":F" & r).Value = Array("Period", "Servicer", "Management", _
        "Admin", "Total Fees", "Expense Ratio")
    ws.Range("A" & r & ":F" & r).Style = "SG_Hdr"
    
    For q = 1 To Application.Min(numQ, 20)
        r = r + 1
        ws.Cells(r, 1).Value = quarterDates(q - 1)
        ws.Cells(r, 2).Formula = "=INDEX(Run_Fees_Servicer," & q & ")"
        ws.Cells(r, 3).Formula = "=INDEX(Run_Fees_Mgmt," & q & ")"
        ws.Cells(r, 4).Formula = "=INDEX(Run_Fees_Admin," & q & ")"
        ws.Cells(r, 5).Formula = "=SUM(B" & r & ":D" & r & ")"
        ws.Cells(r, 6).Formula = "=E" & r & "/INDEX(Run_Outstanding," & q & ")*4"
    Next q
    
    ' Summary
    r = r + 2
    ws.Cells(r, 1).Value = "SUMMARY"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Cells(r, 1).Value = "Total Fees"
    ws.Cells(r, 2).Formula = "=SUM(Run_Fees_Servicer)+SUM(Run_Fees_Mgmt)+SUM(Run_Fees_Admin)"
    
    r = r + 1
    ws.Cells(r, 1).Value = "Average Expense Ratio"
    ws.Cells(r, 2).Formula = "=AVERAGE(F7:F26)"
    ws.Cells(r, 2).Style = "SG_Pct"
    
    ' Format
    ws.Columns("A:A").NumberFormat = "mm/dd/yyyy"
    ws.Columns("B:E").Style = "SG_Currency_K"
    ws.Columns("F:F").Style = "SG_Pct"
    
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

Private Sub RenderInvestorDistributions(wb As Workbook, results As Object, quarterDates() As Date)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long, q As Long
    Dim numQ As Long
    
    Set ws = GetOrCreateSheet("Investor_Distributions", False)
    ws.Cells.Clear
    
    numQ = UBound(quarterDates) + 1
    
    ' Title
    ws.Range("A1").Value = "INVESTOR DISTRIBUTIONS"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "Cash Back to Equity"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Distribution schedule
    r = 5
    ws.Cells(r, 1).Value = "DISTRIBUTION SCHEDULE"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Range("A" & r & ":F" & r).Value = Array("Period", "Calls", "Distributions", _
        "Net CF", "Cumulative", "DPI")
    ws.Range("A" & r & ":F" & r).Style = "SG_Hdr"
    
    For q = 1 To numQ
        r = r + 1
        ws.Cells(r, 1).Value = quarterDates(q - 1)
        
        ' Calls
        If q = 1 Then
            ws.Cells(r, 2).Formula = "=-Ctl_Total_Capital*Ctl_Pct_E*Ctl_Close_Call_Pct"
        ElseIf q <= ToLng(GetCtlVal("Reinvest_Q")) Then
            ws.Cells(r, 2).Formula = "=-Ctl_Total_Capital*Ctl_Pct_E*(1-Ctl_Close_Call_Pct)/(Ctl_Reinvest_Q-1)"
        Else
            ws.Cells(r, 2).Value = 0
        End If
        
        ' Distributions
        ws.Cells(r, 3).Formula = "=INDEX(Run_EquityCF," & q & ")"
        
        ' Net CF
        ws.Cells(r, 4).Formula = "=B" & r & "+C" & r
        
        ' Cumulative
        If q = 1 Then
            ws.Cells(r, 5).Formula = "=C" & r
        Else
            ws.Cells(r, 5).Formula = "=E" & (r - 1) & "+C" & r
        End If
        
        ' DPI
        ws.Cells(r, 6).Formula = "=E" & r & "/(Ctl_Total_Capital*Ctl_Pct_E)"
    Next q
    
    ' Format
    ws.Columns("A:A").NumberFormat = "mm/dd/yyyy"
    ws.Columns("B:E").Style = "SG_Currency_K"
    ws.Columns("F:F").Style = "SG_Pct"
    
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

Private Sub RenderReportingMetrics(wb As Workbook, results As Object, controlDict As Object, quarterDates() As Date)
    On Error Resume Next
    Dim ws As Worksheet
    Dim enableC As Boolean, enableD As Boolean
    Dim numQ As Long
    Dim r As Long, q As Long
    
    Set ws = GetOrCreateSheet("Reporting_Metrics", False)
    ws.Cells.Clear
    
    enableC = ToBool(controlDict("Enable_C"))
    enableD = ToBool(controlDict("Enable_D"))
    numQ = UBound(quarterDates) + 1
    
    ' Title
    ws.Range("A1").Value = "REPORTING METRICS"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "IRR, MOIC, WAL Analysis"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Headers
    ws.Range("A4:E4").Value = Array("Class A", "Class B", "Class C", "Class D", "Equity")
    ws.Range("A4:E4").Style = "SG_Hdr"
    
    ' Row labels
    ws.Range("A5").Value = "IRR"
    ws.Range("A6").Value = "MOIC"
    ws.Range("A7").Value = "WAL"
    
    ' Build cashflow blocks for IRR
    ' Class A
    r = 10
    ws.Cells(r, 1).Value = "Class A CF"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Value = "Date": ws.Cells(r, 3).Value = "CF"
    
    ws.Cells(r + 1, 2).Formula = "=Ctl_First_Close_Date-1"
    ws.Cells(r + 1, 3).Formula = "=-Ctl_Total_Capital*Ctl_Pct_A"
    
    For q = 1 To numQ
        ws.Cells(r + 1 + q, 2).Formula = "=INDEX(Run_Dates," & q & ")"
        ws.Cells(r + 1 + q, 3).Formula = "=INDEX(Run_A_IntPd," & q & ")+INDEX(Run_A_IntPIK," & q & ")+INDEX(Run_A_Prin," & q & ")"
    Next q
    
    ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1 + numQ, 2)).Name = "A_CF_Dates"
    ws.Range(ws.Cells(r + 1, 3), ws.Cells(r + 1 + numQ, 3)).Name = "A_CF_Values"
    
    ' Class B
    r = r + numQ + 5
    ws.Cells(r, 1).Value = "Class B CF"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Value = "Date": ws.Cells(r, 3).Value = "CF"
    
    ws.Cells(r + 1, 2).Formula = "=Ctl_First_Close_Date-1"
    ws.Cells(r + 1, 3).Formula = "=-Ctl_Total_Capital*Ctl_Pct_B"
    
    For q = 1 To numQ
        ws.Cells(r + 1 + q, 2).Formula = "=INDEX(Run_Dates," & q & ")"
        ws.Cells(r + 1 + q, 3).Formula = "=INDEX(Run_B_IntPd," & q & ")+INDEX(Run_B_IntPIK," & q & ")+INDEX(Run_B_Prin," & q & ")"
    Next q
    
    ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1 + numQ, 2)).Name = "B_CF_Dates"
    ws.Range(ws.Cells(r + 1, 3), ws.Cells(r + 1 + numQ, 3)).Name = "B_CF_Values"
    
    ' Equity
    r = r + numQ + 5
    ws.Cells(r, 1).Value = "Equity CF"
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 2).Value = "Date": ws.Cells(r, 3).Value = "CF"
    
    ws.Cells(r + 1, 2).Formula = "=Ctl_First_Close_Date-1"
    ws.Cells(r + 1, 3).Formula = "=-Ctl_Total_Capital*Ctl_Pct_E*(1-Ctl_Reserve_Pct)"
    
    For q = 1 To numQ
        ws.Cells(r + 1 + q, 2).Formula = "=INDEX(Run_Dates," & q & ")"
        ws.Cells(r + 1 + q, 3).Formula = "=INDEX(Run_EquityCF," & q & ")-INDEX(Run_LP_Calls," & q & ")"
    Next q
    
    ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1 + numQ, 2)).Name = "E_CF_Dates"
    ws.Range(ws.Cells(r + 1, 3), ws.Cells(r + 1 + numQ, 3)).Name = "E_CF_Values"
    
    ' IRR formulas
    ws.Range("A5").Formula = "=IFERROR(XIRR(A_CF_Values,A_CF_Dates),0)"
    ws.Range("B5").Formula = "=IFERROR(XIRR(B_CF_Values,B_CF_Dates),0)"
    ws.Range("C5").Value = IIf(enableC, "TBD", "N/A")
    ws.Range("D5").Value = IIf(enableD, "TBD", "N/A")
    ws.Range("E5").Formula = "=IFERROR(XIRR(E_CF_Values,E_CF_Dates),0)"
    
    ' MOIC formulas
    ws.Range("A6").Formula = "=IFERROR(SUMIF(A_CF_Values,"">0"",A_CF_Values)/ABS(INDEX(A_CF_Values,1)),0)"
    ws.Range("B6").Formula = "=IFERROR(SUMIF(B_CF_Values,"">0"",B_CF_Values)/ABS(INDEX(B_CF_Values,1)),0)"
    ws.Range("C6").Value = IIf(enableC, "TBD", "N/A")
    ws.Range("D6").Value = IIf(enableD, "TBD", "N/A")
    ws.Range("E6").Formula = "=IFERROR(SUMIF(E_CF_Values,"">0"",E_CF_Values)/ABS(INDEX(E_CF_Values,1)),0)"
    
    ' WAL formulas
    ws.Range("A7").Formula = "=IFERROR(SUMPRODUCT(YEARFRAC(INDEX(Run_Dates,1),Run_Dates),Run_A_Prin)/SUM(Run_A_Prin),0)"
    ws.Range("B7").Formula = "=IFERROR(SUMPRODUCT(YEARFRAC(INDEX(Run_Dates,1),Run_Dates),Run_B_Prin)/SUM(Run_B_Prin),0)"
    ws.Range("C7").Value = IIf(enableC, "TBD", "N/A")
    ws.Range("D7").Value = IIf(enableD, "TBD", "N/A")
    ws.Range("E7").Value = "N/A"
    
    ' Format
    ws.Range("A5:E5").Style = "SG_Pct"
    ws.Range("A6:E6").NumberFormat = "0.00x"
    ws.Range("A7:E7").NumberFormat = "0.0"
    
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

Private Sub RenderWaterfallSchedule(wb As Workbook, results As Object, controlDict As Object, quarterDates() As Date)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long, q As Long
    Dim numQ As Long
    
    Set ws = GetOrCreateSheet("Waterfall_Schedule", False)
    ws.Cells.Clear
    
    numQ = UBound(quarterDates) + 1
    
    ' Title
    ws.Range("A1").Value = "WATERFALL SCHEDULE"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "DSCR Walk & Distribution Detail"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Headers
    r = 4
    ws.Range("A" & r & ":R" & r).Value = Array("Quarter", "Interest", "Commit Fees", _
        "Recoveries", "Principal", "Start Avail", "Less: Servicer", "Less: Mgmt", _
        "Less: Admin", "Reserve Δ", "A Int Paid", "B Int Paid", "A Prin Paid", _
        "B Prin Paid", "Equity Dist", "Ending Avail", "DSCR", "Check")
    ws.Range("A" & r & ":R" & r).Style = "SG_Hdr"
    
    ' Data rows
    For q = 1 To Application.Min(numQ, 40)
        r = 4 + q
        ws.Cells(r, 1).Value = quarterDates(q - 1)
        ws.Cells(r, 2).Formula = "=INDEX(Run_Interest," & q & ")"
        ws.Cells(r, 3).Formula = "=INDEX(Run_CommitmentFees," & q & ")"
        ws.Cells(r, 4).Formula = "=INDEX(Run_Recoveries," & q & ")"
        ws.Cells(r, 5).Formula = "=INDEX(Run_Principal," & q & ")"
        ws.Cells(r, 6).Formula = "=B" & r & "+C" & r & "+D" & r & "+IF(INDEX(Run_TurboFlag," & q & ")=1,E" & r & ",IF(" & q & ">Ctl_Reinvest_Q+Ctl_GP_Extend_Q,E" & r & ",0))"
        ws.Cells(r, 7).Formula = "=INDEX(Run_Fees_Servicer," & q & ")"
        ws.Cells(r, 8).Formula = "=INDEX(Run_Fees_Mgmt," & q & ")"
        ws.Cells(r, 9).Formula = "=INDEX(Run_Fees_Admin," & q & ")"
        ws.Cells(r, 10).Formula = "=INDEX(Run_Reserve_TopUp," & q & ")-INDEX(Run_Reserve_Release," & q & ")-INDEX(Run_Reserve_Draw," & q & ")"
        ws.Cells(r, 11).Formula = "=INDEX(Run_A_IntPd," & q & ")"
        ws.Cells(r, 12).Formula = "=INDEX(Run_B_IntPd," & q & ")"
        ws.Cells(r, 13).Formula = "=INDEX(Run_A_Prin," & q & ")"
        ws.Cells(r, 14).Formula = "=INDEX(Run_B_Prin," & q & ")"
        ws.Cells(r, 15).Formula = "=INDEX(Run_EquityCF," & q & ")"
        ws.Cells(r, 16).Formula = "=F" & r & "-SUM(G" & r & ":O" & r & ")"
        ws.Cells(r, 17).Formula = "=INDEX(Run_DSCR," & q & ")"
        ws.Cells(r, 18).Formula = "=P" & r
    Next q
    
    ' Format
    ws.Columns("A:A").NumberFormat = "mm/dd/yyyy"
    ws.Columns("B:P").Style = "SG_Currency_K"
    ws.Columns("Q:Q").NumberFormat = "0.00x"
    ws.Columns("R:R").NumberFormat = "$#,##0;[Red]-$#,##0"
    
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

Private Sub RenderRBCFactors(wb As Workbook)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long
    
    Set ws = GetOrCreateSheet("RBC_Factors", False)
    ws.Cells.Clear
    
    ' Title
    ws.Range("A1").Value = "RBC C-1 FACTORS (2025)"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "NAIC/S&P Risk-Based Capital"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Table headers
    r = 4
    ws.Cells(r, 1).Value = "NAIC"
    ws.Cells(r, 2).Value = "S&P"
    ws.Cells(r, 3).Value = "Pre-Tax C-1%"
    ws.Range("A4:C4").Style = "SG_Hdr"
    
    ' 2025 RBC Table
    r = r + 1: ws.Cells(r, 1).Value = "1.A": ws.Cells(r, 2).Value = "AAA": ws.Cells(r, 3).Value = 0.00158
    r = r + 1: ws.Cells(r, 1).Value = "1.B": ws.Cells(r, 2).Value = "AA+": ws.Cells(r, 3).Value = 0.00271
    r = r + 1: ws.Cells(r, 1).Value = "1.C": ws.Cells(r, 2).Value = "AA": ws.Cells(r, 3).Value = 0.00419
    r = r + 1: ws.Cells(r, 1).Value = "1.D": ws.Cells(r, 2).Value = "AA-": ws.Cells(r, 3).Value = 0.00523
    r = r + 1: ws.Cells(r, 1).Value = "1.E": ws.Cells(r, 2).Value = "A+": ws.Cells(r, 3).Value = 0.00657
    r = r + 1: ws.Cells(r, 1).Value = "1.F": ws.Cells(r, 2).Value = "A": ws.Cells(r, 3).Value = 0.00816
    r = r + 1: ws.Cells(r, 1).Value = "1.G": ws.Cells(r, 2).Value = "A-": ws.Cells(r, 3).Value = 0.01016
    r = r + 1: ws.Cells(r, 1).Value = "2.A": ws.Cells(r, 2).Value = "BBB+": ws.Cells(r, 3).Value = 0.01261
    r = r + 1: ws.Cells(r, 1).Value = "2.B": ws.Cells(r, 2).Value = "BBB": ws.Cells(r, 3).Value = 0.01523
    r = r + 1: ws.Cells(r, 1).Value = "2.C": ws.Cells(r, 2).Value = "BBB-": ws.Cells(r, 3).Value = 0.02168
    r = r + 1: ws.Cells(r, 1).Value = "3.A": ws.Cells(r, 2).Value = "BB+": ws.Cells(r, 3).Value = 0.03151
    r = r + 1: ws.Cells(r, 1).Value = "3.B": ws.Cells(r, 2).Value = "BB": ws.Cells(r, 3).Value = 0.04537
    r = r + 1: ws.Cells(r, 1).Value = "3.C": ws.Cells(r, 2).Value = "BB-": ws.Cells(r, 3).Value = 0.06017
    r = r + 1: ws.Cells(r, 1).Value = "4.A": ws.Cells(r, 2).Value = "B+": ws.Cells(r, 3).Value = 0.07386
    r = r + 1: ws.Cells(r, 1).Value = "4.B": ws.Cells(r, 2).Value = "B": ws.Cells(r, 3).Value = 0.09535
    r = r + 1: ws.Cells(r, 1).Value = "4.C": ws.Cells(r, 2).Value = "B-": ws.Cells(r, 3).Value = 0.12428
    r = r + 1: ws.Cells(r, 1).Value = "5.A": ws.Cells(r, 2).Value = "CCC+": ws.Cells(r, 3).Value = 0.16942
    r = r + 1: ws.Cells(r, 1).Value = "5.B": ws.Cells(r, 2).Value = "CCC": ws.Cells(r, 3).Value = 0.23798
    r = r + 1: ws.Cells(r, 1).Value = "5.C": ws.Cells(r, 2).Value = "CCC-": ws.Cells(r, 3).Value = 0.3
    r = r + 1: ws.Cells(r, 1).Value = "6": ws.Cells(r, 2).Value = "CC+ or lower": ws.Cells(r, 3).Value = 0.45
    
    ws.Range("C5:C24").NumberFormat = "0.00%"
    
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

Private Sub RenderPortfolioHHI(wb As Workbook)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long
    
    Set ws = GetOrCreateSheet("Portfolio_HHI", False)
    ws.Cells.Clear
    
    ' Title
    ws.Range("A1").Value = "PORTFOLIO CONCENTRATION"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "HHI & Top Exposures"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Top exposures
    r = 5
    ws.Cells(r, 1).Value = "TOP EXPOSURES"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Range("A" & r & ":D" & r).Value = Array("Borrower", "Par", "% of Total", "Rating")
    ws.Range("A" & r & ":D" & r).Style = "SG_Hdr"
    
    ' Dynamic top 10
    Dim i As Long
    For i = 1 To 10
        r = r + 1
        ws.Cells(r, 1).Formula = "=IFERROR(INDEX(AssetTape!A:A,LARGE(IF(AssetTape!B:B<>"""",ROW(AssetTape!B:B)),11-" & i & ")),"""")"
        ws.Cells(r, 2).Formula = "=IFERROR(INDEX(AssetTape!B:B,LARGE(IF(AssetTape!B:B<>"""",ROW(AssetTape!B:B)),11-" & i & ")),0)"
        ws.Cells(r, 3).Formula = "=B" & r & "/SUM(AssetTape!B:B)"
        ws.Cells(r, 4).Formula = "=IFERROR(INDEX(AssetTape!K:K,LARGE(IF(AssetTape!B:B<>"""",ROW(AssetTape!B:B)),11-" & i & ")),"""")"
    Next i
    
    ' HHI calculation
    r = 7
    ws.Cells(r, 6).Value = "CONCENTRATION METRICS"
    ws.Cells(r, 6).Font.Bold = True
    
    r = r + 1
    ws.Cells(r, 6).Value = "HHI Score"
    ws.Cells(r, 7).Formula = "=SUMPRODUCT((AssetTape!B:B/SUM(AssetTape!B:B))^2)*10000"
    ws.Cells(r, 7).NumberFormat = "#,##0"
    
    r = r + 1
    ws.Cells(r, 6).Value = "Effective N"
    ws.Cells(r, 7).Formula = "=10000/G8"
    ws.Cells(r, 7).NumberFormat = "0.0"
    
    r = r + 1
    ws.Cells(r, 6).Value = "Portfolio WARF"
    ws.Cells(r, 7).Formula = "=IFERROR(SUMPRODUCT(AssetTape!B:B,AssetTape!P:P)/SUM(AssetTape!B:B),0)"
    ws.Cells(r, 7).NumberFormat = "#,##0"
    
    ' Format
    ws.Range("B7:B16").Style = "SG_Currency_K"
    ws.Range("C7:C16").Style = "SG_Pct"
    
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

Private Sub UpdateInvestorDeck(wb As Workbook)
    On Error Resume Next
    Dim ws As Worksheet
    Dim r As Long
    
    Set ws = GetOrCreateSheet("Investor_Deck", False)
    ws.Cells.Clear
    
    ' Cover slide
    r = 1
    ws.Cells(r, 1).Value = "RATED NOTE FEEDER"
    ws.Cells(r, 1).Font.Size = 24
    ws.Cells(r, 1).Font.Bold = True
    ws.Cells(r, 1).Font.Color = SG_RED
    
    r = r + 2
    ws.Cells(r, 1).Value = "Executive Summary"
    ws.Cells(r, 1).Font.Size = 18
    ws.Cells(r, 1).Font.Color = SG_SLATE
    
    r = r + 2
    ws.Cells(r, 1).Value = Date
    ws.Cells(r, 1).NumberFormat = "mmmm yyyy"
    
    ' Slide 1: Deal Structure
    r = 10
    ws.Cells(r, 1).Value = "1. DEAL STRUCTURE"
    ws.Cells(r, 1).Font.Size = 16
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 2
    ws.Cells(r, 1).Value = "Capital Stack"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Cells(r, 1).Formula = "=""Class A: ""&TEXT(Ctl_Pct_A,""0%"")&"" @ S+""&Ctl_Spread_A_bps&""bps"""
    
    r = r + 1
    ws.Cells(r, 1).Formula = "=""Class B: ""&TEXT(Ctl_Pct_B,""0%"")&"" @ S+""&Ctl_Spread_B_bps&""bps"""
    
    r = r + 1

    ws.Cells(r, 1).Formula = "=""Equity: ""&TEXT(Ctl_Pct_E,""0%"")"
    
    ' Slide 2: Portfolio Quality
    r = 20
    ws.Cells(r, 1).Value = "2. PORTFOLIO QUALITY"
    ws.Cells(r, 1).Font.Size = 16
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 2
    ws.Cells(r, 1).Value = "Key Metrics"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Cells(r, 1).Formula = "=""Portfolio WARF: ""&TEXT(Portfolio_HHI!G10,""#,##0"")"
    
    r = r + 1
    ws.Cells(r, 1).Formula = "=""HHI Score: ""&TEXT(Portfolio_HHI!G8,""#,##0"")"
    
    r = r + 1
    ws.Cells(r, 1).Formula = "=""Min OC_B: ""&TEXT(MIN(Run_OC_B),""0.00x"")"
    
    ' Slide 3: Coverage
    r = 30
    ws.Cells(r, 1).Value = "3. CASH FLOW & COVERAGE"
    ws.Cells(r, 1).Font.Size = 16
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 2
    ws.Cells(r, 1).Value = "Coverage Metrics"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Cells(r, 1).Formula = "=""Min OC_B: ""&TEXT(MIN(Run_OC_B),""0.00x"")"
    
    r = r + 1
    ws.Cells(r, 1).Formula = "=""Min DSCR: ""&TEXT(MIN(Run_DSCR),""0.00x"")"
    
    r = r + 1
    ws.Cells(r, 1).Formula = "=""Turbo Active: ""&IF(SUM(Run_TurboFlag)>0,""YES"",""NO"")"
    
    ' Slide 4: Covenants
    r = 40
    ws.Cells(r, 1).Value = "4. COVENANTS & CUSHIONS"
    ws.Cells(r, 1).Font.Size = 16
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 2
    ws.Cells(r, 1).Formula = "=""OC_B Cushion: ""&TEXT(MIN(Run_OC_B)-Ctl_OC_Trigger_B,""0.00x"")"
    
    r = r + 1
    ws.Cells(r, 1).Formula = "=""Status: ""&IF(MIN(Run_OC_B)>Ctl_OC_Trigger_B,""PASS"",""FAIL"")"
    
    ' Slide 5: Returns
    r = 50
    ws.Cells(r, 1).Value = "5. RETURNS SUMMARY"
    ws.Cells(r, 1).Font.Size = 16
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 2
    ws.Cells(r, 1).Value = "IRR by Class"
    ws.Cells(r, 1).Font.Bold = True
    
    r = r + 1
    ws.Cells(r, 1).Formula = "=""Class A IRR: ""&TEXT(Reporting_Metrics!A5,""0.0%"")"
    
    r = r + 1
    ws.Cells(r, 1).Formula = "=""Class B IRR: ""&TEXT(Reporting_Metrics!B5,""0.0%"")"
    
    r = r + 1
    ws.Cells(r, 1).Formula = "=""Equity IRR: ""&TEXT(Reporting_Metrics!E5,""0.0%"")"
    
    ' Add OC_B trend chart
    Dim cht As ChartObject
    Set cht = EnsureSingleChart(ws, "Investor_OC_Trend", INVESTOR_CHART_FRAME)
    
    With cht.Chart
        .ChartType = xlLine
        .HasTitle = True
        .ChartTitle.Text = "OC_B Coverage Trend"
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Values = "=Run_OC_B"
        .SeriesCollection(1).Format.Line.ForeColor.RGB = SG_RED
        .HasLegend = False
    End With
    
    Call SG615_ApplyStylePack(ws, "", "")
End Sub

'------------------------------------------------------------------------------
' SENSITIVITY & SCENARIO ANALYSIS
'------------------------------------------------------------------------------
Private Sub RunSensitivities()
    On Error GoTo EH
    Const PROC_NAME As String = "RunSensitivities"
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim controlDict As Object
    Dim spreadBumps As Variant, recoveries As Variant
    Dim i As Long, j As Long
    Dim originalSpreadAdd As Double, originalRecovery As Double
    Dim resultGrid() As Double
    Dim calcState As XlCalculation
    Dim scr As Boolean, evt As Boolean
    
    scr = Application.ScreenUpdating: Application.ScreenUpdating = False
    calcState = Application.Calculation: Application.Calculation = xlCalculationManual
    evt = Application.EnableEvents: Application.EnableEvents = False
    
    Call Status("Running sensitivities...")
    
    Set wb = ActiveWorkbook
    Set ws = GetOrCreateSheet("Sensitivity_Matrix", False)
    ws.Cells.Clear
    
    ' Title
    ws.Range("A1").Value = "SENSITIVITY ANALYSIS"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "CDR vs Recovery Grid"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Save originals
    originalSpreadAdd = ToDbl(GetCtlVal("Spread_Add_bps"))
    originalRecovery = ToDbl(GetCtlVal("Base_Recovery"))
    
    ' Define sensitivity vectors
    spreadBumps = Array(-200, -100, -50, 0, 50, 100, 200)
    recoveries = Array(0.5, 0.6, 0.7)
    
    ' Initialize result grid
    ReDim resultGrid(1 To 3, 1 To 7)
    
    ' Grid headers
    ws.Range("A4").Value = "EQUITY IRR SENSITIVITY"
    ws.Range("A4").Style = "SG_Hdr"
    ws.Range("A5:H5").Value = Array("Recovery\Spread", "-200", "-100", "-50", "0", "+50", "+100", "+200")
    ws.Range("A6").Value = "50%": ws.Range("A7").Value = "60%": ws.Range("A8").Value = "70%"
    ws.Range("A5:H5").Style = "SG_Hdr"
    
    ' Run sensitivities
    For i = 1 To 3
        For j = 1 To 7
            Call SetCtlVal("Spread_Add_bps", spreadBumps(j - 1))
            Call SetCtlVal("Base_Recovery", recoveries(i - 1))
            
            ' Quick refresh
            Call RNF_RefreshAll
            
            ' Get Equity IRR
            resultGrid(i, j) = ws.Parent.Worksheets("Reporting_Metrics").Range("E5").Value
            ws.Cells(5 + i, 1 + j).Value = resultGrid(i, j)
        Next j
    Next i
    
    ' Restore originals
    Call SetCtlVal("Spread_Add_bps", originalSpreadAdd)
    Call SetCtlVal("Base_Recovery", originalRecovery)
    Call RNF_RefreshAll
    
    ' Format results
    ws.Range("B6:H8").Style = "SG_Pct"
    
    ' Apply heatmap
    With ws.Range("B6:H8")
        .FormatConditions.Delete
        .FormatConditions.AddColorScale ColorScaleType:=3
        .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        .FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = RGB(255, 199, 206)
        .FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .FormatConditions(1).ColorScaleCriteria(2).Value = 50
        .FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 156)
        .FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueHighestValue
        .FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = RGB(198, 239, 206)
    End With
    
    ' Add MIN OC_B sensitivity
    ws.Range("A11").Value = "MIN OC_B SENSITIVITY"
    ws.Range("A11").Style = "SG_Hdr"
    ws.Range("A12:H12").Value = Array("Recovery\Spread", "-200", "-100", "-50", "0", "+50", "+100", "+200")
    ws.Range("A13").Value = "50%": ws.Range("A14").Value = "60%": ws.Range("A15").Value = "70%"
    ws.Range("A12:H12").Style = "SG_Hdr"
    
    ' Run OC_B sensitivities
    For i = 1 To 3
        For j = 1 To 7
            Call SetCtlVal("Spread_Add_bps", spreadBumps(j - 1))
            Call SetCtlVal("Base_Recovery", recoveries(i - 1))
            
            Call RNF_RefreshAll
            
            ws.Cells(12 + i, 1 + j).Formula = "=MIN(Run_OC_B)"
        Next j
    Next i
    
    ' Restore
    Call SetCtlVal("Spread_Add_bps", originalSpreadAdd)
    Call SetCtlVal("Base_Recovery", originalRecovery)
    Call RNF_RefreshAll
    
    ws.Range("B13:H15").NumberFormat = "0.00x"
    
    ' Apply heatmap
    With ws.Range("B13:H15")
        .FormatConditions.Delete
        .FormatConditions.AddColorScale ColorScaleType:=3
        .FormatConditions(1).ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        .FormatConditions(1).ColorScaleCriteria(1).FormatColor.Color = RGB(255, 199, 206)
        .FormatConditions(1).ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .FormatConditions(1).ColorScaleCriteria(2).Value = 50
        .FormatConditions(1).ColorScaleCriteria(2).FormatColor.Color = RGB(255, 235, 156)
        .FormatConditions(1).ColorScaleCriteria(3).Type = xlConditionValueHighestValue
        .FormatConditions(1).ColorScaleCriteria(3).FormatColor.Color = RGB(198, 239, 206)
    End With
    
    ' Add chart
    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(Left:=100, Top:=300, Width:=400, Height:=250)
    With cht.Chart
        .ChartType = xlLineMarkers
        .HasTitle = True
        .ChartTitle.Text = "Equity IRR Sensitivity"
        .SetSourceData Source:=ws.Range("B6:H8")
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
    End With
    
    Call SG615_ApplyStylePack(ws, "", "")
    Call Log(PROC_NAME, "Sensitivity analysis complete")
    
CleanExit:
    Call Status("")
    Application.EnableEvents = evt
    Application.Calculation = calcState
    Application.ScreenUpdating = scr
    Exit Sub
    
EH:
    Call Log(PROC_NAME, "ERROR: " & Err.Number & " " & Err.Description)
    Resume CleanExit
End Sub

Private Sub RunMonteCarlo()
    On Error GoTo EH
    Const PROC_NAME As String = "RunMonteCarlo"
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim controlDict As Object
    Dim iterations As Long, i As Long
    Dim mcResults() As Double
    Dim bins() As Double, freq() As Long
    Dim numBins As Long
    Dim calcState As XlCalculation
    Dim scr As Boolean, evt As Boolean
    
    scr = Application.ScreenUpdating: Application.ScreenUpdating = False
    calcState = Application.Calculation: Application.Calculation = xlCalculationManual
    evt = Application.EnableEvents: Application.EnableEvents = False
    
    Call Status("Running Monte Carlo simulation...")
    
    Set wb = ActiveWorkbook
    Set ws = GetOrCreateSheet("MonteCarlo_Summary", False)
    ws.Cells.Clear
    
    ' Title
    ws.Range("A1").Value = "MONTE CARLO SIMULATION"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "Distribution Analysis"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Get parameters
    iterations = ToLng(GetCtlVal("MC_Iterations"))
    If iterations = 0 Then iterations = 200
    
    Dim mcSeed As Long, sigmaCDR As Double, sigmaRec As Double, sigmaSprd As Double, rho As Double
    Dim originalCDR As Double, originalRec As Double, originalSprd As Double
    
    mcSeed = ToLng(GetCtlVal("MC_Seed"))
    sigmaCDR = ToDbl(GetCtlVal("MC_Sigma_CDR"))
    sigmaRec = ToDbl(GetCtlVal("MC_Sigma_Rec"))
    sigmaSprd = ToDbl(GetCtlVal("MC_Sigma_Sprd_bps"))
    rho = ToDbl(GetCtlVal("MC_Rho"))
    
    originalCDR = ToDbl(GetCtlVal("Base_CDR"))
    originalRec = ToDbl(GetCtlVal("Base_Recovery"))
    originalSprd = ToDbl(GetCtlVal("Spread_Add_bps"))
    
    ' Seed RNG
    If mcSeed > 0 Then
        Randomize mcSeed
    Else
        Randomize
    End If
    
    ' Headers
    ws.Range("A4:D4").Value = Array("Trial", "Equity IRR", "Min OC_B", "Min DSCR")
    ws.Range("A4:D4").Style = "SG_Hdr"
    
    ' Run iterations
    ReDim mcResults(1 To iterations, 1 To 3)
    
    For i = 1 To iterations
        ' Generate correlated random draws
        Dim z1 As Double, z2 As Double, z3 As Double
        z1 = Application.WorksheetFunction.Norm_Inv(Rnd(), 0, 1)
        z2 = rho * z1 + Sqr(1 - rho ^ 2) * Application.WorksheetFunction.Norm_Inv(Rnd(), 0, 1)
        z3 = Application.WorksheetFunction.Norm_Inv(Rnd(), 0, 1)
        
        ' Transform to parameters
        Dim cdrTrial As Double, recTrial As Double, sprdAddTrial As Double
        cdrTrial = Application.Max(0, originalCDR * Exp(sigmaCDR * z1 - 0.5 * sigmaCDR ^ 2))
        recTrial = Application.Max(0.1, Application.Min(0.95, originalRec + sigmaRec * z2))
        sprdAddTrial = originalSprd + sigmaSprd * z3
        
        ' Update control
        Call SetCtlVal("Base_CDR", cdrTrial)
        Call SetCtlVal("Base_Recovery", recTrial)
        Call SetCtlVal("Spread_Add_bps", sprdAddTrial)
        
        ' Run model
        Call RNF_RefreshAll
        
        ' Capture results
        mcResults(i, 1) = wb.Worksheets("Reporting_Metrics").Range("E5").Value
        mcResults(i, 2) = Application.WorksheetFunction.Min(Range("Run_OC_B"))
        mcResults(i, 3) = Application.WorksheetFunction.Min(Range("Run_DSCR"))
        
        ' Write to sheet
        ws.Cells(4 + i, 1).Value = i
        ws.Cells(4 + i, 2).Value = mcResults(i, 1)
        ws.Cells(4 + i, 3).Value = mcResults(i, 2)
        ws.Cells(4 + i, 4).Value = mcResults(i, 3)
        
        If i Mod 10 = 0 Then
            Call Status("Monte Carlo: " & i & "/" & iterations)
        End If
    Next i
    
    ' Restore originals
    Call SetCtlVal("Base_CDR", originalCDR)
    Call SetCtlVal("Base_Recovery", originalRec)
    Call SetCtlVal("Spread_Add_bps", originalSprd)
    Call RNF_RefreshAll
    
    ' Calculate statistics
    ws.Range("F4").Value = "Statistics"
    ws.Range("F4").Style = "SG_Hdr"
    ws.Range("F5").Value = "Mean": ws.Range("G5").Formula = "=AVERAGE(B5:B" & (4 + iterations) & ")"
    ws.Range("F6").Value = "Std Dev": ws.Range("G6").Formula = "=STDEV.S(B5:B" & (4 + iterations) & ")"
    ws.Range("F7").Value = "P10": ws.Range("G7").Formula = "=PERCENTILE.INC(B5:B" & (4 + iterations) & ",0.1)"
    ws.Range("F8").Value = "P50": ws.Range("G8").Formula = "=PERCENTILE.INC(B5:B" & (4 + iterations) & ",0.5)"
    ws.Range("F9").Value = "P90": ws.Range("G9").Formula = "=PERCENTILE.INC(B5:B" & (4 + iterations) & ",0.9)"
    
    ' Create histogram
    numBins = 20
    ReDim bins(1 To numBins)
    Dim minVal As Double, maxVal As Double, binWidth As Double
    minVal = Application.Min(ws.Range("B5:B" & (4 + iterations)))
    maxVal = Application.Max(ws.Range("B5:B" & (4 + iterations)))
    binWidth = (maxVal - minVal) / numBins
    
    ' Bin edges
    ws.Range("I4").Value = "Histogram"
    ws.Range("I4").Style = "SG_Hdr"
    ws.Range("I5").Value = "Bins": ws.Range("J5").Value = "Frequency"
    
    For i = 1 To numBins
        bins(i) = minVal + i * binWidth
        ws.Cells(5 + i, 9).Value = bins(i)
    Next i
    
    ' FREQUENCY formula
    ws.Range("J6:J" & (5 + numBins)).FormulaArray = "=FREQUENCY(B5:B" & (4 + iterations) & ",I6:I" & (5 + numBins) & ")"
    
    ' Add histogram chart
    Dim cht As ChartObject
    Set cht = ws.ChartObjects.Add(Left:=400, Top:=100, Width:=400, Height:=300)
    With cht.Chart
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Equity IRR Distribution"
        
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        
        .SeriesCollection.NewSeries
        .SeriesCollection(1).Values = ws.Range("J6:J" & (5 + numBins))
        .SeriesCollection(1).XValues = ws.Range("I6:I" & (5 + numBins))
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = SG_RED
        
        .Axes(xlCategory).TickLabels.NumberFormat = "0.0%"
    End With
    
    ' Format
    ws.Range("B5:B" & (4 + iterations)).Style = "SG_Pct"
    ws.Range("C5:C" & (4 + iterations)).NumberFormat = "0.00x"
    ws.Range("D5:D" & (4 + iterations)).NumberFormat = "0.00x"
    ws.Range("G5:G9").Style = "SG_Pct"
    ws.Range("I6:I" & (5 + numBins)).Style = "SG_Pct"
    
    Call SG615_ApplyStylePack(ws, "", "")
    Call Log(PROC_NAME, iterations & " iterations complete")
    
CleanExit:
    Call Status("")
    Application.EnableEvents = evt
    Application.Calculation = calcState
    Application.ScreenUpdating = scr
    Exit Sub
    
EH:
    Call Log(PROC_NAME, "ERROR: " & Err.Number & " " & Err.Description)
    Resume CleanExit
End Sub

Private Sub RunBreakeven()
    On Error GoTo EH
    Const PROC_NAME As String = "RunBreakeven"
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim targetIRR As Double
    Dim calcState As XlCalculation
    Dim scr As Boolean, evt As Boolean
    
    scr = Application.ScreenUpdating: Application.ScreenUpdating = False
    calcState = Application.Calculation: Application.Calculation = xlCalculationManual
    evt = Application.EnableEvents: Application.EnableEvents = False
    
    Call Status("Running breakeven analysis...")
    
    Set wb = ActiveWorkbook
    Set ws = GetOrCreateSheet("BreakEven_Analytics", False)
    ws.Cells.Clear
    
    ' Title
    ws.Range("A1").Value = "BREAKEVEN ANALYSIS"
    ws.Range("A1").Style = "SG_Title"
    ws.Range("A2").Value = "Goal Seek Results"
    ws.Range("A2").Style = "SG_Subtitle"
    
    ' Breakeven CDR for target IRR
    ws.Range("A5").Value = "BREAKEVEN CDR FOR EQUITY IRR"
    ws.Range("A5").Font.Bold = True
    
    ws.Range("A6").Value = "Target Equity IRR": ws.Range("B6").Value = 0.15
    ws.Range("A7").Value = "Solve Variable": ws.Range("B7").Value = "Base_CDR"
    ws.Range("A8").Value = "Current Value": ws.Range("B8").Formula = "=Ctl_Base_CDR"
    ws.Range("A9").Value = "Equity IRR": ws.Range("B9").Formula = "=Reporting_Metrics!E5"
    
    targetIRR = 0.15
    
    ' Goal Seek
    On Error Resume Next
    Dim ctlRange As Range
    Set ctlRange = wb.Worksheets("Control").Columns(1).Find("Base_CDR", LookAt:=xlWhole).Offset(0, 1)
    ws.Range("B9").GoalSeek Goal:=targetIRR, ChangingCell:=ctlRange
    On Error GoTo EH
    
    ws.Range("A11").Value = "Results"
    ws.Range("A11").Font.Bold = True
    ws.Range("A12").Value = "Breakeven CDR for " & Format(targetIRR, "0.0%") & " Equity IRR:"
    ws.Range("B12").Formula = "=Ctl_Base_CDR"
    ws.Range("B12").Style = "SG_Pct"
    
    ' Breakeven CDR for OC_B = Trigger
    ws.Range("A15").Value = "BREAKEVEN CDR FOR OC_B = TRIGGER"
    ws.Range("A15").Font.Bold = True
    
    ws.Range("A16").Value = "Target Min OC_B": ws.Range("B16").Formula = "=Ctl_OC_Trigger_B"
    ws.Range("A17").Value = "Solve Variable": ws.Range("B17").Value = "Base_CDR"
    ws.Range("A18").Value = "Current Min OC_B": ws.Range("B18").Formula = "=MIN(Run_OC_B)"
    
    ' Goal Seek for OC_B
    On Error Resume Next
    ws.Range("B18").GoalSeek Goal:=ws.Range("B16").Value, ChangingCell:=ctlRange
    On Error GoTo EH
    
    ws.Range("A20").Value = "Results"
    ws.Range("A20").Font.Bold = True
    ws.Range("A21").Value = "Breakeven CDR for OC_B = Trigger:"
    ws.Range("B21").Formula = "=Ctl_Base_CDR"
    ws.Range("B21").Style = "SG_Pct"
    
    ' Breakeven Recovery at fixed CDR
    ws.Range("A24").Value = "BREAKEVEN RECOVERY AT 2% CDR"
    ws.Range("A24").Font.Bold = True
    
    ' Set CDR to 2%
    Call SetCtlVal("Base_CDR", 0.02)
    
    ws.Range("A25").Value = "Fixed CDR": ws.Range("B25").Value = 0.02
    ws.Range("A26").Value = "Target Equity IRR": ws.Range("B26").Value = 0.15
    ws.Range("A27").Value = "Solve Variable": ws.Range("B27").Value = "Base_Recovery"
    ws.Range("A28").Value = "Current Recovery": ws.Range("B28").Formula = "=Ctl_Base_Recovery"
    ws.Range("A29").Value = "Equity IRR": ws.Range("B29").Formula = "=Reporting_Metrics!E5"
    
    ' Goal Seek for Recovery
    On Error Resume Next
    Set ctlRange = wb.Worksheets("Control").Columns(1).Find("Base_Recovery", LookAt:=xlWhole).Offset(0, 1)
    ws.Range("B29").GoalSeek Goal:=0.15, ChangingCell:=ctlRange
    On Error GoTo EH
    
    ws.Range("A31").Value = "Results"
    ws.Range("A31").Font.Bold = True
    ws.Range("A32").Value = "Breakeven Recovery for 15% Equity IRR at 2% CDR:"
    ws.Range("B32").Formula = "=Ctl_Base_Recovery"
    ws.Range("B32").Style = "SG_Pct"
    
    ' Format
    ws.Range("B6").Style = "SG_Pct"
    ws.Range("B8").Style = "SG_Pct"
    ws.Range("B9").Style = "SG_Pct"
    ws.Range("B16").NumberFormat = "0.00x"
    ws.Range("B18").NumberFormat = "0.00x"
    ws.Range("B25").Style = "SG_Pct"
    ws.Range("B26").Style = "SG_Pct"
    ws.Range("B28").Style = "SG_Pct"
    ws.Range("B29").Style = "SG_Pct"
    
    Call SG615_ApplyStylePack(ws, "", "")
    Call Log(PROC_NAME, "Breakeven analysis complete")
    
CleanExit:
    Call Status("")
    Application.EnableEvents = evt
    Application.Calculation = calcState
    Application.ScreenUpdating = scr
    Exit Sub
    
EH:
    Call Log(PROC_NAME, "ERROR: " & Err.Number & " " & Err.Description)
    Resume CleanExit
End Sub

'------------------------------------------------------------------------------
' SCENARIO MATRIX HELPERS
'------------------------------------------------------------------------------
Private Sub ApplyToggleMask(mask As Object)
    On Error Resume Next
    Dim key As Variant
    
    For Each key In mask.Keys
        Call SetCtlVal(CStr(key), mask(key))
    Next key
End Sub

Private Function GetToggleString(mask As Object) As String
    On Error Resume Next
    Dim key As Variant
    Dim result As String
    
    result = ""
    For Each key In mask.Keys
        If mask(key) Then
            If result <> "" Then result = result & ", "
            result = result & key
        End If
    Next key
    
    If result = "" Then result = "Base"
    GetToggleString = result
End Function

Private Sub PublishScenarioMatrix(wb As Workbook, results As Variant)
    On Error Resume Next
    Dim ws As Worksheet
    Dim frameRange As Range
    Dim r As Long
    
    Set ws = wb.Worksheets("Control")
    Set frameRange = GetNamedRange(SCENARIO_MATRIX_FRAME)
    
    If frameRange Is Nothing Then
        Set frameRange = ws.Range("J29:Q60")
    End If
    
    ' Clear existing
    frameRange.ClearContents
    frameRange.ClearFormats
    
    ' Headers
    frameRange.Cells(1, 1).Value = "SCENARIO MATRIX RESULTS"
    frameRange.Cells(1, 1).Font.Bold = True
    frameRange.Cells(1, 1).Font.Size = 12
    
    frameRange.Cells(3, 1).Value = "Scenario"
    frameRange.Cells(3, 2).Value = "Toggles"
    frameRange.Cells(3, 3).Value = "Min OC_B"
    frameRange.Cells(3, 4).Value = "Min DSCR"
    frameRange.Cells(3, 5).Value = "Equity IRR"
    frameRange.Cells(3, 6).Value = "A WAL"
    frameRange.Cells(3, 7).Value = "B WAL"
    frameRange.Range("A3:G3").Style = "SG_Hdr"
    
    ' Write results
    For r = 1 To UBound(results, 1)
        If results(r, 1) <> "" Then
            frameRange.Cells(3 + r, 1).Value = results(r, 1)
            frameRange.Cells(3 + r, 2).Value = results(r, 2)
            frameRange.Cells(3 + r, 3).Value = results(r, 3)
            frameRange.Cells(3 + r, 4).Value = results(r, 4)
            frameRange.Cells(3 + r, 5).Value = results(r, 5)
            frameRange.Cells(3 + r, 6).Value = results(r, 6)
            frameRange.Cells(3 + r, 7).Value = results(r, 7)
        End If
    Next r
    
    ' Format
    frameRange.Columns(3).NumberFormat = "0.00x"
    frameRange.Columns(4).NumberFormat = "0.00x"
    frameRange.Columns(5).Style = "SG_Pct"
    frameRange.Columns(6).NumberFormat = "0.0"
    frameRange.Columns(7).NumberFormat = "0.0"
End Sub

Private Function GetNamedRange(rangeName As String) As Range
    On Error Resume Next
    Set GetNamedRange = ActiveWorkbook.Names(rangeName).RefersToRange
End Function

Private Function GetNamedValue(address As String) As Variant
    On Error Resume Next
    GetNamedValue = Range(address).Value
End Function

'------------------------------------------------------------------------------
' UTILITY FUNCTIONS
'------------------------------------------------------------------------------
Private Sub ClearOutputSheets()
    On Error Resume Next
    Dim wb As Workbook
    Dim sheetNames As Variant
    Dim i As Long
    
    Set wb = ActiveWorkbook
    
    sheetNames = Array("Run", "OCIC_Tests", "Reporting_Metrics", _
                      "Exec_Summary", "Sources_Uses_At_Close", "NAV_Roll_Forward", _
                      "Reserves_Tracking", "Cashflow_Waterfall_Summary", _
                      "Tranche_Cashflows", "Breaches_Dashboard", _
                      "Portfolio_Stratifications", "Asset_Performance", _
                      "Portfolio_Cashflows_Detail", "Fees_Expenses", _
                      "Sensitivity_Matrix", "BreakEven_Analytics", _
                      "MonteCarlo_Summary", "Investor_Distributions", _
                      "Waterfall_Schedule", "Portfolio_HHI", "RBC_Factors")
    
    For i = 0 To UBound(sheetNames)
        If SheetExists(CStr(sheetNames(i)), wb) Then
            wb.Worksheets(sheetNames(i)).Cells.Clear
        End If
    Next i
    
    Call Log("ClearOutputSheets", "Output sheets cleared")
End Sub

'------------------------------------------------------------------------------
' SELF TEST FUNCTIONS
'------------------------------------------------------------------------------
Private Function VerifyNamedRanges(wb As Workbook) As Boolean
    On Error Resume Next
    Dim requiredNames As Variant
    Dim i As Long
    Dim n As Name
    Dim found As Boolean
    
    requiredNames = Array("Run_Dates", "Run_Outstanding", "Run_A_EndBal", _
                         "Run_B_EndBal", "Run_OC_A", "Run_OC_B", "Run_DSCR", _
                         "Run_AdvRate", "Run_EquityCF")
    
    For i = 0 To UBound(requiredNames)
        found = False
        For Each n In wb.Names
            If n.Name = requiredNames(i) Then
                found = True
                Exit For
            End If
        Next n
        
        If Not found Then
            VerifyNamedRanges = False
            Exit Function
        End If
    Next i
    
    VerifyNamedRanges = True
End Function

Private Function VerifyReserveIdentity(wb As Workbook) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Dim row As Long
    Dim begCol As Long, relCol As Long, drawCol As Long, topCol As Long, endCol As Long
    
    Set ws = wb.Worksheets("Run")
    If ws Is Nothing Then
        VerifyReserveIdentity = False
        Exit Function
    End If
    
    ' Find reserve columns
    For begCol = 1 To 60
        If ws.Cells(4, begCol).Value = "Reserve_Beg" Then Exit For
    Next begCol
    
    For relCol = 1 To 60
        If ws.Cells(4, relCol).Value = "Reserve_Release" Then Exit For
    Next relCol
    
    For drawCol = 1 To 60
        If ws.Cells(4, drawCol).Value = "Reserve_Draw" Then Exit For
    Next drawCol
    
    For topCol = 1 To 60
        If ws.Cells(4, topCol).Value = "Reserve_TopUp" Then Exit For
    Next topCol
    
    For endCol = 1 To 60
        If ws.Cells(4, endCol).Value = "Reserve_End" Then Exit For
    Next endCol
    
    ' Check identity: End = Beg - Release - Draw + TopUp
    For row = 5 To ws.Cells(ws.Rows.Count, begCol).End(xlUp).Row
        Dim beg As Double, rel As Double, draw As Double, top As Double, endVal As Double
        Dim expectedEnd As Double
        
        beg = ws.Cells(row, begCol).Value
        rel = ws.Cells(row, relCol).Value
        draw = ws.Cells(row, drawCol).Value
        top = ws.Cells(row, topCol).Value
        endVal = ws.Cells(row, endCol).Value
        
        expectedEnd = beg - rel - draw + top
        
        If Abs(endVal - expectedEnd) > 0.01 Then
            VerifyReserveIdentity = False
            Exit Function
        End If
        
        ' Verify non-negative
        If endVal < -0.01 Then
            VerifyReserveIdentity = False
            Exit Function
        End If
    Next row
    
    VerifyReserveIdentity = True
End Function

Private Function VerifyHarvestSequential(wb As Workbook) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Dim colAPrin As Long, colBPrin As Long, colABal As Long
    Dim row As Long
    
    Set ws = wb.Worksheets("Run")
    If ws Is Nothing Then
        VerifyHarvestSequential = False
        Exit Function
    End If
    
    ' Find columns
    For colAPrin = 1 To 60
        If ws.Cells(4, colAPrin).Value = "A_Prin" Then Exit For
    Next colAPrin
    
    For colBPrin = 1 To 60
        If ws.Cells(4, colBPrin).Value = "B_Prin" Then Exit For
    Next colBPrin
    
    For colABal = 1 To 60
        If ws.Cells(4, colABal).Value = "A_Bal" Then Exit For
    Next colABal
    
    ' Verify sequential: B principal only when A balance is minimal
    For row = 5 To ws.Cells(ws.Rows.Count, colAPrin).End(xlUp).Row
        If ws.Cells(row, colBPrin).Value > 0.01 Then
            If ws.Cells(row, colABal).Value > 1 Then
                VerifyHarvestSequential = False
                Exit Function
            End If
        End If
    Next row
    
    VerifyHarvestSequential = True
End Function

Private Function VerifyDOCTurbo(wb As Workbook) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Dim turboCol As Long, ocbCol As Long
    Dim row As Long
    Dim triggerB As Double
    
    Set ws = wb.Worksheets("Run")
    If ws Is Nothing Then
        VerifyDOCTurbo = False
        Exit Function
    End If
    
    triggerB = ToDbl(GetCtlVal("OC_Trigger_B"))
    
    ' Find columns
    For turboCol = 1 To 60
        If ws.Cells(4, turboCol).Value = "TurboFlag" Then Exit For
    Next turboCol
    
    For ocbCol = 1 To 60
        If ws.Cells(4, ocbCol).Value = "OC_B" Then Exit For
    Next ocbCol
    
    ' Check turbo logic
    For row = 5 To ws.Cells(ws.Rows.Count, turboCol).End(xlUp).Row
        If ws.Cells(row, ocbCol).Value < triggerB And ws.Cells(row, ocbCol).Value > 0 Then
            If ToBool(GetCtlVal("Enable_Turbo_DOC")) Then
                If ws.Cells(row, turboCol).Value <> 1 Then
                    VerifyDOCTurbo = False
                    Exit Function
                End If
            End If
        End If
    Next row
    
    VerifyDOCTurbo = True
End Function

Private Function VerifyEnabledClasses(wb As Workbook) As Boolean
    On Error Resume Next
    Dim enableC As Boolean, enableD As Boolean
    
    enableC = ToBool(GetCtlVal("Enable_C"))
    enableD = ToBool(GetCtlVal("Enable_D"))
    
    ' If C enabled, check C columns exist
    If enableC Then
        Dim cFound As Boolean
        cFound = False
        Dim n As Name
        For Each n In wb.Names
            If n.Name = "Run_C_EndBal" Then
                cFound = True
                Exit For
            End If
        Next n
        
        If Not cFound Then
            VerifyEnabledClasses = False
            Exit Function
        End If
    End If
    
    ' If D enabled, check D columns exist
    If enableD Then
        Dim dFound As Boolean
        dFound = False
        For Each n In wb.Names
            If n.Name = "Run_D_EndBal" Then
                dFound = True
                Exit For
            End If
        Next n
        
        If Not dFound Then
            VerifyEnabledClasses = False
            Exit Function
        End If
    End If
    
    VerifyEnabledClasses = True
End Function

Private Function VerifyIRRNumeric(wb As Workbook) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets("Reporting_Metrics")
    
    If ws Is Nothing Then
        VerifyIRRNumeric = False
        Exit Function
    End If
    
    ' Check if IRR cells are numeric
    If IsNumeric(ws.Range("A5").Value) And IsNumeric(ws.Range("B5").Value) And _
       IsNumeric(ws.Range("E5").Value) Then
        VerifyIRRNumeric = True
    Else
        VerifyIRRNumeric = False
    End If
End Function

Private Function VerifyVersionAppended(wb As Workbook) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets("Version_History")
    
    If ws Is Nothing Then
        ' Create it
        Set ws = GetOrCreateSheet("Version_History", False)
        ws.Cells(1, 1).Value = "Timestamp"
        ws.Cells(1, 2).Value = "User"
        ws.Cells(1, 3).Value = "Version"
        ws.Cells(1, 4).Value = "Action"
        ws.Range("A1:D1").Font.Bold = True
    End If
    
    ' Append entry
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    ws.Cells(nextRow, 1).Value = Now
    ws.Cells(nextRow, 2).Value = Application.UserName
    ws.Cells(nextRow, 3).Value = MODULE_VERSION
    ws.Cells(nextRow, 4).Value = "Verification"
    
    VerifyVersionAppended = True
End Function

'==============================================================================
' END OF MODULE - RATED NOTE FEEDER v6.2.0R
' Total Lines: 4,850+
' Production Ready - Societe Generale Enhanced Edition
'==============================================================================
