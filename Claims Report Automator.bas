Attribute VB_Name = "Module1"
Sub getinput()
Dim FinalMonthStr As String, FinalMonth As Integer, FinalYearStr As String, FinalYear As Integer, NYears As Integer, NDivisions As Integer
Dim StartMonthStr As String, StartMonth As Integer, StartYearStr As String, StartYear As Integer
Dim HaveMedical As Boolean, HaveVision As Boolean, Answer As String, Pollsheet As String, PollNCells As Integer
Dim DivisionArray(3) As Variant, d As Integer, MedicalArray(4) As Variant, VisionArray(4) As Variant, RxArray(2) As Variant, m As Integer, v As Integer, r As Integer
Dim NCells As Integer, j As Integer, exists As Boolean, CopyingBook As Workbook, CopyingBookName As String, PastingSheetName As String, wb As Workbook, i As Integer, xrow As Integer, Cellval As String

Dim sh As Worksheet

MedicalArray(0) = "Medical Claims"
MedicalArray(1) = "Medical Count"
MedicalArray(2) = "Medical Per Emp"
MedicalArray(3) = "Medical Per Life"
MedicalArray(4) = "Medical Unit Cost"

VisionArray(0) = "Vision Claims"
VisionArray(1) = "Vision Count"
VisionArray(2) = "Vision Per Emp"
VisionArray(3) = "Vision Per Life"
VisionArray(4) = "Vision Unit Cost"

RxArray(0) = "Rx Claims"
RxArray(1) = "Rx Per Emp"
RxArray(2) = "Rx Per Life"

Application.DisplayAlerts = False


For m = 0 To 4
    For Each sh In Application.Worksheets
        If sh.Name = MedicalArray(m) Then
            sh.Delete
        End If
    Next sh
Next m
For v = 0 To 4
    For Each sh In Application.Worksheets
        If sh.Name = VisionArray(v) Then
            sh.Delete
        End If
    Next sh
Next v
For r = 0 To 2
    For Each sh In Application.Worksheets
        If sh.Name = RxArray(r) Then
            sh.Delete
        End If
    Next sh
Next r
For Each sh In Application.Worksheets
    If sh.Name = "Rx Claims" Then
        sh.Delete
    End If
Next sh
For Each sh In Application.Worksheets
    If sh.Name = "Enrollment" Then
        sh.Delete
    End If
Next sh
For Each sh In Application.Worksheets
    If sh.Name = "Dashboard" Then
        sh.Delete
    End If
Next sh
For Each sh In Application.Worksheets
    If sh.Name = "Dashboard Graphs" Then
        sh.Delete
    End If
Next sh
For Each sh In Application.Worksheets
    If sh.Name = "Medical Graphs" Then
        sh.Delete
    End If
Next sh
For Each sh In Application.Worksheets
    If sh.Name = "Rx Graphs" Then
        sh.Delete
    End If
Next sh
For Each sh In Application.Worksheets
    If sh.Name = "Vision Graphs" Then
        sh.Delete
    End If
Next sh
For Each sh In Application.Worksheets
    If sh.Name = "CLM25" Then
        sh.Delete
    End If
Next sh
For Each sh In Application.Worksheets
    If sh.Name = "CS" Then
        sh.Delete
    End If
Next sh
For Each sh In Application.Worksheets
    If sh.Name = "MDCLMS" Then
        sh.Delete
    End If
Next sh
For Each sh In Application.Worksheets
    If sh.Name = "RXCLMS" Then
        sh.Delete
    End If
Next sh
For Each sh In Application.Worksheets
    If sh.Name = "VSCLMS" Then
        sh.Delete
    End If
Next sh

Application.DisplayAlerts = True

Dim FilePath As String, FolderPath As String, StringFind As Integer, BackSlash As String

'''''Old Opening Code

'''''FilePath = Application.ActiveWorkbook.Path
'''''
'''''StringFind = 0
'''''
'''''BackSlash = "NOTBACKSLASH"
'''''
'''''Do While BackSlash <> "\"
'''''
'''''If Mid(FilePath, Len(FilePath) - StringFind, 1) = "\" Then
'''''BackSlash = "\"
'''''End If
'''''
'''''StringFind = StringFind + 1
'''''
'''''Loop
'''''
'''''Dim qr As Queries
'''''
'''''FolderPath = Right(FilePath, StringFind - 1)
''''
''''''On Error GoTo MakeQuery
'''''    ActiveWorkbook.Queries(FolderPath).Delete
'''''MakeQuery:
''''''On Error GoTo Handler
'''''    ActiveWorkbook.Queries.Add Name:=FolderPath, Formula:="let" & Chr(13) & "" & Chr(10) & "Source = Folder.Files(" & """" & FilePath & """" & ")" & Chr(13) & "" & Chr(10) & "in" & Chr(13) & "" & Chr(10) & "Source"
'''''    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & FolderPath & ";Extended Properties=""""", Destination:=Range("$A$1")).QueryTable
'''''        .CommandType = xlCmdSql
'''''        .CommandText = Array("SELECT * FROM [" & FolderPath & "]")
'''''        .RowNumbers = False
'''''        .FillAdjacentFormulas = False
'''''        .PreserveFormatting = True
'''''        .RefreshOnFileOpen = False
'''''        .BackgroundQuery = True
'''''        .RefreshStyle = xlInsertDeleteCells
'''''        .SavePassword = False
'''''        .SaveData = True
'''''        .AdjustColumnWidth = True
'''''        .RefreshPeriod = 0
'''''        .PreserveColumnInfo = True
'''''        '.ListObject.DisplayName = FolderPath
'''''        .Refresh BackgroundQuery:=False
'''''    End With
'''''    'Application.CommandBars("Queries and Connections").Visible = False
''''
'''''i = 0
'''''xrow = 2
'''''ActiveSheet.Range("H1") = "NCells"
'''''ActiveSheet.Range("H2") = "=COUNTA(A:A)"
'''''
'''''Dim NCells As Integer
'''''NCells = ActiveSheet.Range("H2")
'''''
'''''ReDim FileArray(0)
'''''
'''''Cellval = ActiveSheet.Cells(xrow, 1)
'''''
'''''Do While xrow <= NCells
'''''    If Cellval <> "" And Left(Cellval, 1) <> "~" And Cellval <> ThisWorkbook.Name Then
'''''    FileArray(i) = Cellval
'''''    End If
'''''    xrow = xrow + 1
'''''    i = i + 1
'''''    Cellval = ActiveSheet.Cells(xrow, 1)
'''''    If Cellval <> "" And Left(Cellval, 1) <> "~" And Cellval <> ThisWorkbook.Name Then
'''''    ReDim Preserve FileArray(i)
'''''    End If
'''''Loop
'''''
'''''Dim wb As Workbook
'''''
'''''i = 0
'''''
'''''Dim OpeningBookName As String
'''''
'''''FilePath = Sheets("MACRO").Range("F2")
'''''
'''''For i = LBound(FileArray, 1) To UBound(FileArray, 1)
'''''    OpeningBookName = FileArray(i)
'''''    Workbooks.Open Filename:=FilePath & "\" & OpeningBookName
'''''Next i
'''''
'''''Dim j As Integer, exists As Boolean, CopyingBook As Workbook, CopyingBookName As String, PastingSheetName As String
'''''
'''''Application.DisplayAlerts = False
'''''
'''''exists = False
'''''i = 0
'''''
'''''ThisWorkbook.Activate
'''''
'''''For i = LBound(FileArray, 1) To UBound(FileArray, 1)
'''''
'''''    CopyingBookName = Left(FileArray(i), Len(FileArray(i)) - 4)
'''''
'''''    PastingSheetName = Left(CopyingBookName, 31)
'''''
'''''    j = 1
'''''
'''''    For j = 1 To Worksheets.Count
'''''        If Worksheets(j).Name = PastingSheetName Then
'''''           exists = True
'''''        End If
'''''    Next j
'''''
'''''    If exists = True Then
'''''        Sheets(PastingSheetName).Delete
'''''    End If
'''''    Sheets.Add After:=ActiveSheet
'''''    ActiveSheet.Select
'''''    ActiveSheet.Name = PastingSheetName
'''''
'''''    Set CopyingBook = Workbooks.Open(Application.ActiveWorkbook.Path & "\" & CopyingBookName)
'''''
'''''    Sheets(PastingSheetName).UsedRange.Copy
'''''    ThisWorkbook.Activate
'''''    Sheets(PastingSheetName).Cells.PasteSpecial Paste:=xlPasteValues
'''''
'''''    exists = False
'''''
'''''Next i
''''
'''''Rename Sheets
''''
'''''j = 1
'''''
'''''Dim DeleteSheet As Boolean
'''''
'''''DeleteSheet = False
'''''
'''''For Each sh In Application.Worksheets
'''''    If sh.Name = "CS" Then
'''''        DeleteSheet = True
'''''        GoTo DeleteSheet
'''''    End If
'''''    If sh.Name = "CLM25" Then
'''''        DeleteSheet = True
'''''        GoTo DeleteSheet
'''''    End If
'''''    If sh.Name = "VSCLMS" Then
'''''        DeleteSheet = True
'''''        GoTo DeleteSheet
'''''    End If
'''''    If sh.Name = "RXCLMS" Then
'''''        DeleteSheet = True
'''''        GoTo DeleteSheet
'''''    End If
'''''    If sh.Name = "MDCLMS" Then
'''''        DeleteSheet = True
'''''        GoTo DeleteSheet
'''''    End If
'''''    If sh.Name = "Dashboard" Then
'''''        DeleteSheet = True
'''''        GoTo DeleteSheet
'''''    End If
'''''DeleteSheet:
'''''    If DeleteSheet = True Then
'''''        sh.Delete
'''''    End If
'''''    DeleteSheet = False
'''''Next sh
'''''
'''''j = 1
'''''
'''''For j = 1 To Worksheets.Count
'''''    Worksheets(j).Activate
'''''    If Left(Worksheets(j).Name, 2) = "CS" Then
'''''        ActiveSheet.Select
'''''        ActiveSheet.Name = "CS"
'''''    End If
'''''    If Left(Worksheets(j).Name, 5) = "CLM25" Then
'''''        ActiveSheet.Select
'''''        ActiveSheet.Name = "CLM25"
'''''    End If
'''''    If Left(Worksheets(j).Name, 6) = "RXCLMS" Then
'''''        ActiveSheet.Select
'''''        ActiveSheet.Name = "RXCLMS"
'''''    End If
'''''    If Left(Worksheets(j).Name, 6) = "VSCLMS" Then
'''''        ActiveSheet.Select
'''''        ActiveSheet.Name = "VSCLMS"
'''''    End If
'''''    If Left(Worksheets(j).Name, 6) = "MDCLMS" Then
'''''        ActiveSheet.Select
'''''        ActiveSheet.Name = "MDCLMS"
'''''    End If
'''''Next j

'Open and Copy Sheets

Sheets("MACRO").Activate

Dim FileArray(4) As Variant

FileArray(0) = "CLM25.csv"
FileArray(1) = "CS.csv"
FileArray(2) = "MDCLMS.csv"
FileArray(3) = "RXCLMS.csv"
FileArray(4) = "VSCLMS.csv"

For i = LBound(FileArray, 1) To UBound(FileArray, 1)
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Select
    ActiveSheet.Name = FileArray(i)
Next i

Sheets("MACRO").Activate

For i = LBound(FileArray, 1) To UBound(FileArray, 1)

    Workbooks.Open Application.ActiveWorkbook.Path & "\" & FileArray(i)
    Windows(FileArray(i)).Activate
    Cells.Select
    Selection.Copy
    Windows("new messa WIP").Activate
    Sheets(FileArray(i)).Activate
    Sheets(FileArray(i)).Range("A1").Select
    Selection.PasteSpecial

Next i

For Each sh In Application.Worksheets
    If sh.Name = "CLM25.csv" Then
        Sheets("CLM25.csv").Activate
        ActiveSheet.Select
        ActiveSheet.Name = "CLM25"
    End If
    If sh.Name = "CS.csv" Then
        Sheets("CS.csv").Activate
        ActiveSheet.Select
        ActiveSheet.Name = "CS"
    End If
    If sh.Name = "MDCLMS.csv" Then
        Sheets("MDCLMS.csv").Activate
        ActiveSheet.Select
        ActiveSheet.Name = "MDCLMS"
    End If
    If sh.Name = "RXCLMS.csv" Then
        Sheets("RXCLMS.csv").Activate
        ActiveSheet.Select
        ActiveSheet.Name = "RXCLMS"
    End If
    If sh.Name = "VSCLMS.csv" Then
        Sheets("VSCLMS.csv").Activate
        ActiveSheet.Select
        ActiveSheet.Name = "VSCLMS"
    End If
Next sh

'Prompt for data

Application.DisplayAlerts = True

HaveMedical = False
HaveVision = False

For Each sh In Application.Worksheets
    If sh.Name = "MDCLMS" Then
        HaveMedical = True
    End If
    If sh.Name = "VSCLMS" Then
        HaveVision = True
    End If
Next sh

If HaveMedical = False And HaveVision = False Then
    GoTo Handler
End If

If HaveMedical = True Then
    Pollsheet = "MDCLMS"
Else
    Pollsheet = "VSCLMS"
End If

PollNCells = Application.WorksheetFunction.CountA(Sheets(Pollsheet).Range("A:A")) + 1

StartYearStr = Left(Sheets(Pollsheet).Cells(5, 1), 4)
StartYear = StartYearStr

FinalYearStr = Left(Sheets(Pollsheet).Cells(PollNCells, 1), 4)
FinalYear = FinalYearStr

If Mid(Sheets(Pollsheet).Cells(5, 1), Len(Sheets(Pollsheet).Cells(5, 1)) - 1, 1) = "/" Then
    StartMonthStr = Right(Sheets(Pollsheet).Cells(5, 1), 1)
Else
    StartMonthStr = Right(Sheets(Pollsheet).Cells(5, 1), 2)
End If
StartMonth = StartMonthStr

If Mid(Sheets(Pollsheet).Cells(5, 1), Len(Sheets(Pollsheet).Cells(PollNCells, 1)) - 1, 1) = "/" Then
    FinalMonthStr = Right(Sheets(Pollsheet).Cells(5, 1), 1)
Else
    FinalMonthStr = Right(Sheets(Pollsheet).Cells(5, 1), 2)
End If
FinalMonth = FinalMonthStr

NYears = FinalYear - StartYear + 1

'''DEBUG MODE

'FinalMonthStr = InputBox("What is the last month with data?")
'
'If FinalMonthStr = "January" Then
'    FinalMonth = 1
'End If
'If FinalMonthStr = "February" Then
'    FinalMonth = 2
'End If
'If FinalMonthStr = "March" Then
'    FinalMonth = 3
'End If
'If FinalMonthStr = "April" Then
'    FinalMonth = 4
'End If
'If FinalMonthStr = "May" Then
'    FinalMonth = 5
'End If
'If FinalMonthStr = "June" Then
'    FinalMonth = 6
'End If
'If FinalMonthStr = "July" Then
'    FinalMonth = 7
'End If
'If FinalMonthStr = "August" Then
'    FinalMonth = 8
'End If
'If FinalMonthStr = "September" Then
'    FinalMonth = 9
'End If
'If FinalMonthStr = "October" Then
'    FinalMonth = 10
'End If
'If FinalMonthStr = "November" Then
'    FinalMonth = 11
'End If
'If FinalMonthStr = "December" Then
'    FinalMonth = 12
'    Else
'    FinalMonth = FinalMonthStr
'End If
'
'If FinalMonth > 12 Or FinalMonth < 1 Then
'    GoTo Handler
'End If
'
'StartMonthStr = InputBox("What is the first month with data?")
'
'If StartMonthStr = "January" Then
'    StartMonth = 1
'End If
'If StartMonthStr = "February" Then
'    StartMonth = 2
'End If
'If StartMonthStr = "March" Then
'    StartMonth = 3
'End If
'If StartMonthStr = "April" Then
'    StartMonth = 4
'End If
'If StartMonthStr = "May" Then
'    StartMonth = 5
'End If
'If StartMonthStr = "June" Then
'    StartMonth = 6
'End If
'If StartMonthStr = "July" Then
'    StartMonth = 7
'End If
'If StartMonthStr = "August" Then
'    StartMonth = 8
'End If
'If StartMonthStr = "September" Then
'    StartMonth = 9
'End If
'If StartMonthStr = "October" Then
'    StartMonth = 10
'End If
'If StartMonthStr = "November" Then
'    StartMonth = 11
'End If
'If StartMonthStr = "December" Then
'    StartMonth = 12
'    Else
'    StartMonth = StartMonthStr
'End If
'
'If StartMonth > 12 Or StartMonth < 1 Then
'    GoTo Handler
'End If
'
'FinalYear = InputBox("What is the last year of data?")
'NYears = InputBox("How many years of data are you using?")
'StartYear = FinalYear + 1 - NYears
'
'Answer = MsgBox("Does this group have medical coverage?", vbQuestion + vbYesNo + vbDefaultButton2, "Medical Coverage?")
'If Answer = vbYes Then
'    HaveMedical = True
'Else
'    HaveMedical = False
'End If
'Answer = MsgBox("Does this group have vision coverage?", vbQuestion + vbYesNo + vbDefaultButton2, "Vision Coverage?")
'If Answer = vbYes Then
'    HaveVision = True
'Else
'    HaveVision = False
'End If
'
'If HaveVision = False And HaveMedical = False Then
'    GoTo Handler
'End If

'Build Sheets

DivisionArray(0) = "Medical"
DivisionArray(1) = "Rx"
DivisionArray(2) = "Vision"

m = 0
v = 0
r = 0

Dim ycolumn As Integer, monthrow As Integer, MonthNumber As Integer, YearCounter As Integer, CurrentYear As Integer

'Medical Sheets

If HaveMedical = True Then

    For m = LBound(MedicalArray, 1) To UBound(MedicalArray, 1)

        Sheets.Add After:=ActiveSheet
        ActiveSheet.Select
        ActiveSheet.Name = MedicalArray(m)
    
        xrow = 1
        ycolumn = 1
        d = 0
        monthrow = 2
        MonthNumber = 1
        YearCounter = NYears
        CurrentYear = StartYear
            
        Do While YearCounter >= 1
        
            Cells(xrow, ycolumn) = CurrentYear
            Cells(xrow, ycolumn + 1) = "Medical"
            Cells(xrow + 1, ycolumn) = "Month"
            
            'PopulateMonths
            Do While MonthNumber <= 12
                Cells(xrow + monthrow, ycolumn) = MonthNumber
                monthrow = monthrow + 1
                MonthNumber = MonthNumber + 1
            Loop
            monthrow = 2
            MonthNumber = 1
            Cells(xrow + 1, ycolumn + 1) = "In-Patient"
            Cells(xrow + 1, ycolumn + 2) = "Lab/X-Ray"
            Cells(xrow + 1, ycolumn + 3) = "Medical/Surgical"
            Cells(xrow + 1, ycolumn + 4) = "Other Equipment"
            Cells(xrow + 1, ycolumn + 5) = "Out-Patient"
            Cells(xrow + 1, ycolumn + 6) = "Total"
            Cells(xrow + 14, ycolumn) = "Total"
            
            YearCounter = YearCounter - 1
            CurrentYear = CurrentYear + 1
            
            ycolumn = ycolumn + 8
            
        Loop
        
    Next m

End If

'Rx Sheets

If HaveMedical = True Then

    For r = LBound(RxArray, 1) To UBound(RxArray, 1)

        Sheets.Add After:=ActiveSheet
        ActiveSheet.Select
        ActiveSheet.Name = RxArray(r)
        
        xrow = 1
        ycolumn = 1
        monthrow = 2
        MonthNumber = 1
        YearCounter = NYears
        CurrentYear = StartYear
            
        Do While YearCounter >= 1
            Cells(xrow, ycolumn) = CurrentYear
            Cells(xrow, ycolumn + 1) = "Rx"
            Cells(xrow + 1, ycolumn) = "Month"
            'PopulateMonths
            Do While MonthNumber <= 12
                Cells(xrow + monthrow, ycolumn) = MonthNumber
                monthrow = monthrow + 1
                MonthNumber = MonthNumber + 1
            Loop
            monthrow = 2
            MonthNumber = 1
            
            Cells(xrow + 1, ycolumn + 1) = "Brand"
            Cells(xrow + 1, ycolumn + 2) = "Generic"
            Cells(xrow + 1, ycolumn + 3) = "Specialty"
            Cells(xrow + 1, ycolumn + 4) = "Total"
            Cells(xrow + 14, ycolumn) = "Total"
            
            YearCounter = YearCounter - 1
            CurrentYear = CurrentYear + 1
            
            ycolumn = ycolumn + 8
            
        Loop
        
    Next r

End If

'Vision Sheets

If HaveVision = True Then

    For v = LBound(VisionArray, 1) To UBound(VisionArray, 1)

        Sheets.Add After:=ActiveSheet
        ActiveSheet.Select
        ActiveSheet.Name = VisionArray(v)
        
        xrow = 1
        ycolumn = 1
        YearCounter = NYears
        CurrentYear = StartYear
            
        Do While YearCounter >= 1
            
            Cells(xrow, ycolumn) = CurrentYear
            Cells(xrow, ycolumn + 1) = "Vision"
            Cells(xrow + 1, ycolumn) = "Month"
            'PopulateMonths
            Do While MonthNumber <= 12
            Cells(xrow + monthrow, ycolumn) = MonthNumber
            monthrow = monthrow + 1
            MonthNumber = MonthNumber + 1
            Loop
            monthrow = 2
            MonthNumber = 1
            'Populate column titles
            Cells(xrow + 1, ycolumn + 1) = "Contact Lenses"
            Cells(xrow + 1, ycolumn + 2) = "Exam"
            Cells(xrow + 1, ycolumn + 3) = "Frames"
            Cells(xrow + 1, ycolumn + 4) = "Lenses"
            Cells(xrow + 1, ycolumn + 5) = "Miscellaneous"
            Cells(xrow + 1, ycolumn + 6) = "Total"
            Cells(xrow + 14, ycolumn) = "Total"
            
            YearCounter = YearCounter - 1
            CurrentYear = CurrentYear + 1
            
            ycolumn = ycolumn + 8
            
        Loop

    Next v

End If

'Enrollment Sheet

Sheets.Add After:=ActiveSheet
ActiveSheet.Select
ActiveSheet.Name = "Enrollment"

Dim CSNCells As Integer
CSNCells = Application.WorksheetFunction.CountA(Sheets("CS").Range("A:A")) + 1

Range("A1") = Sheets("CS").Range("E2")
Range("A3") = "Medical/RX"
Range("A4") = "Average birth year"
Range("A5") = Application.WorksheetFunction.AverageIf(Sheets("CS").Range("D5:D" & CSNCells), "<>*No Coverage", Sheets("CS").Range("A5:A" & CSNCells))
Range("B4") = "Employees by gender"
Range("B5") = "Male:"
Range("B6") = "Female:"
Range("B7") = "Total:"
Range("C5") = Application.WorksheetFunction.CountIf(Sheets("CS").Range("B5:B" & CSNCells), "M") - Application.WorksheetFunction.CountIfs(Sheets("CS").Range("B5:B" & CSNCells), "M", Sheets("CS").Range("D5:D" & CSNCells), "No Coverage")
Range("C6") = Application.WorksheetFunction.CountIf(Sheets("CS").Range("B5:B" & CSNCells), "F") - Application.WorksheetFunction.CountIfs(Sheets("CS").Range("B5:B" & CSNCells), "F", Sheets("CS").Range("D5:D" & CSNCells), "No Coverage")
Range("C7") = Application.WorksheetFunction.Sum(Sheets("Enrollment").Range("C5:C6"))
Range("D4") = "Covered Lives"
Range("D5") = Application.WorksheetFunction.Sum(Sheets("CS").Range("E5:E" & CSNCells))

Range("F3") = "Vision"
Range("F4") = "Average birth year"
Range("F5") = Application.WorksheetFunction.AverageIf(Sheets("CS").Range("F5:F" & CSNCells), "<>*No Coverage", Sheets("CS").Range("A5:A" & CSNCells))
Range("G4") = "Employees by gender"
Range("G5") = "Male:"
Range("G6") = "Female:"
Range("G7") = "Total:"
Range("H5") = Application.WorksheetFunction.CountIf(Sheets("CS").Range("B5:B" & CSNCells), "M") - Application.WorksheetFunction.CountIfs(Sheets("CS").Range("B5:B" & CSNCells), "M", Sheets("CS").Range("F5:F" & CSNCells), "No Coverage")
Range("H6") = Application.WorksheetFunction.CountIf(Sheets("CS").Range("B5:B" & CSNCells), "F") - Application.WorksheetFunction.CountIfs(Sheets("CS").Range("B5:B" & CSNCells), "F", Sheets("CS").Range("F5:F" & CSNCells), "No Coverage")
Range("H7") = Application.WorksheetFunction.Sum(Sheets("Enrollment").Range("H5:H6"))
Range("I4") = "Covered Lives"
Range("I5") = Application.WorksheetFunction.Sum(Sheets("CS").Range("G5:G" & CSNCells))

'Populate Sheets

Dim ClaimNCells As Integer, FinalMonthTemp As Integer, uprow As Integer, backcolumn As Integer, SLoop As Integer, RxLoops As Integer, holdrow As Integer, yearrow As Integer
Dim ColumnSelection As Integer, ClaimsSheet As Worksheet, PopulationSheet As Worksheet

d = 0
SLoop = 1

If HaveMedical = True Then
    GoTo MedicalFill
End If

If HaveVision = True Then
    GoTo VisionFill
End If

GoTo Handler

MedicalFill:
    
    d = 0

    Set ClaimsSheet = Worksheets("MDCLMS")
    
    Set PopulationSheet = Worksheets("Medical Claims")
    
    ColumnSelection = 4
    
    ClaimNCells = Application.WorksheetFunction.CountA(ClaimsSheet.Range("A:A")) + 1
    
    GoTo StartPopulation
    
RxFill:

    d = 1

    Set ClaimsSheet = Worksheets("RXCLMS")
    
    Set PopulationSheet = Worksheets("Rx Claims")
    
    ColumnSelection = 9
    
    ClaimNCells = Application.WorksheetFunction.CountA(ClaimsSheet.Range("A:A")) + 1
    
    GoTo StartPopulation

VisionFill:

    d = 2

    Set ClaimsSheet = Worksheets("VSCLMS")
    
    Set PopulationSheet = Worksheets("Vision Claims")
    
    ColumnSelection = 5
    
    ClaimNCells = Application.WorksheetFunction.CountA(ClaimsSheet.Range("A:A")) + 1
    
    GoTo StartPopulation
    
StartPopulation:
    
FinalMonthTemp = FinalMonth

uprow = 0
backcolumn = 0
holdrow = 0
yearrow = 0

PopulationSheet.Activate

ycolumn = 8 * (NYears - 1) + 1

'Medical/Vision Population

    If d = 0 Or d = 2 Then

FinalYear:
        If NYears > 1 Then
                
            Do While backcolumn < 5
                Cells(FinalMonthTemp + 2, ycolumn + 5 - backcolumn) = ClaimsSheet.Cells(ClaimNCells - uprow, ColumnSelection)
                uprow = uprow + 1
                backcolumn = backcolumn + 1
            Loop
            
            backcolumn = 0
            FinalMonthTemp = FinalMonthTemp - 1
            If FinalMonthTemp > 0 Then
                GoTo FinalYear
            End If
            
            SLoop = 1
            
            Do While SLoop < 6
                Cells(15, ycolumn + SLoop) = Application.WorksheetFunction.Sum(Range(Cells(3, ycolumn + SLoop), Cells(14, ycolumn + SLoop)))
                SLoop = SLoop + 1
            Loop
            SLoop = 1
            
            Do While SLoop < 14
                Cells(SLoop + 2, ycolumn + 6) = Application.WorksheetFunction.Sum(Range(Cells(SLoop + 2, ycolumn + 1), Cells(SLoop + 2, ycolumn + 5)))
                SLoop = SLoop + 1
            Loop
            SLoop = 1
            
            YearCounter = NYears
            
        Else
            
            Do While backcolumn < 5
                Cells(FinalMonthTemp + 2, ycolumn + 5 - backcolumn) = ClaimsSheet.Cells(ClaimNCells - uprow, ColumnSelection)
                uprow = uprow + 1
                backcolumn = backcolumn + 1
            Loop
            
            backcolumn = 0
            FinalMonthTemp = FinalMonthTemp - 1
            If FinalMonthTemp >= StartMonth Then
                GoTo FinalYear
            End If
            
            SLoop = 1
            
            Do While SLoop < 6
                Cells(15, ycolumn + SLoop) = Application.WorksheetFunction.Sum(Range(Cells(3, ycolumn + SLoop), Cells(14, ycolumn + SLoop)))
                SLoop = SLoop + 1
            Loop
            SLoop = 1
            
            Do While SLoop < 14
                Cells(SLoop + 2, ycolumn + 6) = Application.WorksheetFunction.Sum(Range(Cells(SLoop + 2, ycolumn + 1), Cells(SLoop + 2, ycolumn + 5)))
                SLoop = SLoop + 1
            Loop
            SLoop = 1
            
            YearCounter = NYears
            
        End If
        
NewYear:
        
        YearCounter = YearCounter - 1
        
        FinalMonthTemp = 12
        
        ycolumn = ycolumn - 8
        
        Do While YearCounter > 0
        
            If YearCounter > 1 Then
        
                Do While backcolumn < 5
                    Cells(FinalMonthTemp + 2, ycolumn + 5 - backcolumn) = ClaimsSheet.Cells(ClaimNCells - uprow, ColumnSelection)
                    uprow = uprow + 1
                    backcolumn = backcolumn + 1
                Loop
                
                backcolumn = 0
                FinalMonthTemp = FinalMonthTemp - 1
                If FinalMonthTemp = 0 Then
                    SLoop = 1
                    
                    Do While SLoop < 6
                        Cells(15, ycolumn + SLoop) = Application.WorksheetFunction.Sum(Range(Cells(3, ycolumn + SLoop), Cells(14, ycolumn + SLoop)))
                        SLoop = SLoop + 1
                    Loop
                    SLoop = 1
                    
                    Do While SLoop < 14
                        Cells(SLoop + 2, ycolumn + 6) = Application.WorksheetFunction.Sum(Range(Cells(SLoop + 2, ycolumn + 1), Cells(SLoop + 2, ycolumn + 5)))
                        SLoop = SLoop + 1
                    Loop
                    SLoop = 1
                    GoTo NewYear
                End If
                
            Else
            
                Do While backcolumn < 5
                    Cells(FinalMonthTemp + 2, ycolumn + 5 - backcolumn) = ClaimsSheet.Cells(ClaimNCells - uprow, ColumnSelection)
                    uprow = uprow + 1
                    backcolumn = backcolumn + 1
                Loop
                
                backcolumn = 0
                FinalMonthTemp = FinalMonthTemp - 1
                If FinalMonthTemp < StartMonth Then
                    SLoop = 1
                    
                    Do While SLoop < 6
                        Cells(15, ycolumn + SLoop) = Application.WorksheetFunction.Sum(Range(Cells(3, ycolumn + SLoop), Cells(14, ycolumn + SLoop)))
                        SLoop = SLoop + 1
                    Loop
                    SLoop = 1
                    
                    Do While SLoop < 14
                        Cells(SLoop + 2, ycolumn + 6) = Application.WorksheetFunction.Sum(Range(Cells(SLoop + 2, ycolumn + 1), Cells(SLoop + 2, ycolumn + 5)))
                        SLoop = SLoop + 1
                    Loop
                    SLoop = 1
                    GoTo NewYear
                End If
                
            End If
        
        Loop
    
    End If

'Rx Population

    If d = 1 Then

FinalYearRx:

        If NYears > 1 Then

            ColumnSelection = 9
            
            RxLoops = 4
            
            Do While RxLoops > 0
        
                Do While FinalMonthTemp > 0
                    Cells(2 + FinalMonthTemp, ycolumn + RxLoops) = ClaimsSheet.Cells(ClaimNCells - uprow, ColumnSelection)
                    uprow = uprow + 1
                    FinalMonthTemp = FinalMonthTemp - 1
                Loop
                
                RxLoops = RxLoops - 1
                FinalMonthTemp = FinalMonth
                holdrow = uprow
                uprow = 0
                ColumnSelection = ColumnSelection - 2
                
            Loop
            
            Do While SLoop < 5
                Cells(15, ycolumn + SLoop) = Application.WorksheetFunction.Sum(Range(Cells(3, ycolumn + SLoop), Cells(14, ycolumn + SLoop)))
                SLoop = SLoop + 1
            Loop
            SLoop = 1
        
            YearCounter = NYears
            
        Else
            
            ColumnSelection = 9
            
            RxLoops = 4
            
            Do While RxLoops > 0
        
                Do While FinalMonthTemp >= StartMonth
                    Cells(2 + FinalMonthTemp, ycolumn + RxLoops) = ClaimsSheet.Cells(ClaimNCells - uprow, ColumnSelection)
                    uprow = uprow + 1
                    FinalMonthTemp = FinalMonthTemp - 1
                Loop
                
                RxLoops = RxLoops - 1
                FinalMonthTemp = FinalMonth
                uprow = 0
                ColumnSelection = ColumnSelection - 2
                
            Loop
            
            Do While SLoop < 5
                Cells(15, ycolumn + SLoop) = Application.WorksheetFunction.Sum(Range(Cells(3, ycolumn + SLoop), Cells(14, ycolumn + SLoop)))
                SLoop = SLoop + 1
            Loop
            SLoop = 1
        
            YearCounter = NYears
            
        End If
            
NewYearRx:
    
            YearCounter = YearCounter - 1
            
            FinalMonthTemp = 12
                
            ycolumn = ycolumn - 8
            
            ColumnSelection = 9
            
            RxLoops = 4
            
            Do While YearCounter > 0
            
                If YearCounter > 1 Then
            
                    Do While RxLoops > 0
                
                        Do While FinalMonthTemp > 0
                            Cells(2 + FinalMonthTemp, ycolumn + RxLoops) = ClaimsSheet.Cells(ClaimNCells - uprow - holdrow - yearrow, ColumnSelection)
                            uprow = uprow + 1
                            FinalMonthTemp = FinalMonthTemp - 1
                        Loop
                        
                        RxLoops = RxLoops - 1
                        FinalMonthTemp = 12
                        uprow = 0
                        ColumnSelection = ColumnSelection - 2
                        
                    Loop
                    
                    Do While SLoop < 5
                        Cells(15, ycolumn + SLoop) = Application.WorksheetFunction.Sum(Range(Cells(3, ycolumn + SLoop), Cells(14, ycolumn + SLoop)))
                        SLoop = SLoop + 1
                    Loop
                    SLoop = 1
                    yearrow = yearrow + 12
                    GoTo NewYearRx
                    
                Else
                
                    Do While RxLoops > 0
                
                        Do While FinalMonthTemp >= StartMonth
                            Cells(2 + FinalMonthTemp, ycolumn + RxLoops) = ClaimsSheet.Cells(ClaimNCells - uprow - holdrow - yearrow, ColumnSelection)
                            uprow = uprow + 1
                            FinalMonthTemp = FinalMonthTemp - 1
                        Loop
                        
                        RxLoops = RxLoops - 1
                        FinalMonthTemp = 12
                        uprow = 0
                        ColumnSelection = ColumnSelection - 2
                        
                    Loop
                    
                    Do While SLoop < 5
                        Cells(15, ycolumn + SLoop) = Application.WorksheetFunction.Sum(Range(Cells(3, ycolumn + SLoop), Cells(14, ycolumn + SLoop)))
                        SLoop = SLoop + 1
                    Loop
                    SLoop = 1
                    GoTo NewYearRx
                    
                End If
            
            Loop
    
    End If
    
If HaveMedical = True And d = 0 Then
    GoTo RxFill
End If
If HaveVision = True And d = 1 Then
    GoTo VisionFill
End If
    
'Medical/Vision Count population

d = 0
SLoop = 1

If HaveMedical = True Then
    GoTo MedicalFill2
End If

If HaveVision = True Then
    GoTo VisionFill2
End If

GoTo Handler

MedicalFill2:
    
    d = 0

    Set ClaimsSheet = Worksheets("MDCLMS")
    
    Set PopulationSheet = Worksheets("Medical Count")
    
    ColumnSelection = 3
    
    ClaimNCells = Application.WorksheetFunction.CountA(ClaimsSheet.Range("A:A")) + 1
    
    GoTo StartPopulation2

VisionFill2:

    d = 2

    Set ClaimsSheet = Worksheets("VSCLMS")
    
    Set PopulationSheet = Worksheets("Vision Count")
    
    ColumnSelection = 3
    
    ClaimNCells = Application.WorksheetFunction.CountA(ClaimsSheet.Range("A:A")) + 1
    
    GoTo StartPopulation2
    
StartPopulation2:
    
FinalMonthTemp = FinalMonth

uprow = 0
backcolumn = 0

PopulationSheet.Activate

ycolumn = 8 * (NYears - 1) + 1

'Loop to populate data

'Medical/Vision Population

    If d = 0 Or d = 2 Then

FinalYear2:

        If NYears > 1 Then
        
            Do While backcolumn < 5
                Cells(FinalMonthTemp + 2, ycolumn + 5 - backcolumn) = ClaimsSheet.Cells(ClaimNCells - uprow, ColumnSelection)
                uprow = uprow + 1
                backcolumn = backcolumn + 1
            Loop
            
            backcolumn = 0
            FinalMonthTemp = FinalMonthTemp - 1
            If FinalMonthTemp > 0 Then
                GoTo FinalYear2
            End If
            
            SLoop = 1
            
            Do While SLoop < 6
                Cells(15, ycolumn + SLoop) = Application.WorksheetFunction.Sum(Range(Cells(3, ycolumn + SLoop), Cells(14, ycolumn + SLoop)))
                SLoop = SLoop + 1
            Loop
            SLoop = 1
            
            Do While SLoop < 14
                Cells(SLoop + 2, ycolumn + 6) = Application.WorksheetFunction.Sum(Range(Cells(SLoop + 2, ycolumn + 1), Cells(SLoop + 2, ycolumn + 5)))
                SLoop = SLoop + 1
            Loop
            SLoop = 1
            
            YearCounter = NYears
            
        Else
        
            Do While backcolumn < 5
                Cells(FinalMonthTemp + 2, ycolumn + 5 - backcolumn) = ClaimsSheet.Cells(ClaimNCells - uprow, ColumnSelection)
                uprow = uprow + 1
                backcolumn = backcolumn + 1
            Loop
            
            backcolumn = 0
            FinalMonthTemp = FinalMonthTemp - 1
            If FinalMonthTemp >= StartMonth Then
                GoTo FinalYear2
            End If
            
            SLoop = 1
            
            Do While SLoop < 6
                Cells(15, ycolumn + SLoop) = Application.WorksheetFunction.Sum(Range(Cells(3, ycolumn + SLoop), Cells(14, ycolumn + SLoop)))
                SLoop = SLoop + 1
            Loop
            SLoop = 1
            
            Do While SLoop < 14
                Cells(SLoop + 2, ycolumn + 6) = Application.WorksheetFunction.Sum(Range(Cells(SLoop + 2, ycolumn + 1), Cells(SLoop + 2, ycolumn + 5)))
                SLoop = SLoop + 1
            Loop
            SLoop = 1
            
            YearCounter = NYears
            
        End If
        
NewYear2:
        
        YearCounter = YearCounter - 1
        
        FinalMonthTemp = 12
        
        ycolumn = ycolumn - 8
        
        Do While YearCounter > 0
        
            If YearCounter > 1 Then
        
                Do While backcolumn < 5
                    Cells(FinalMonthTemp + 2, ycolumn + 5 - backcolumn) = ClaimsSheet.Cells(ClaimNCells - uprow, ColumnSelection)
                    uprow = uprow + 1
                    backcolumn = backcolumn + 1
                Loop
                
                backcolumn = 0
                FinalMonthTemp = FinalMonthTemp - 1
                If FinalMonthTemp = 0 Then
                    SLoop = 1
                    
                    Do While SLoop < 6
                        Cells(15, ycolumn + SLoop) = Application.WorksheetFunction.Sum(Range(Cells(3, ycolumn + SLoop), Cells(14, ycolumn + SLoop)))
                        SLoop = SLoop + 1
                    Loop
                    SLoop = 1
                    
                    Do While SLoop < 14
                        Cells(SLoop + 2, ycolumn + 6) = Application.WorksheetFunction.Sum(Range(Cells(SLoop + 2, ycolumn + 1), Cells(SLoop + 2, ycolumn + 5)))
                        SLoop = SLoop + 1
                    Loop
                    SLoop = 1
                    GoTo NewYear2
                End If
                
            Else
            
                Do While backcolumn < 5
                    Cells(FinalMonthTemp + 2, ycolumn + 5 - backcolumn) = ClaimsSheet.Cells(ClaimNCells - uprow, ColumnSelection)
                    uprow = uprow + 1
                    backcolumn = backcolumn + 1
                Loop
                
                backcolumn = 0
                FinalMonthTemp = FinalMonthTemp - 1
                If FinalMonthTemp < StartMonth Then
                    SLoop = 1
                    
                    Do While SLoop < 6
                        Cells(15, ycolumn + SLoop) = Application.WorksheetFunction.Sum(Range(Cells(3, ycolumn + SLoop), Cells(14, ycolumn + SLoop)))
                        SLoop = SLoop + 1
                    Loop
                    SLoop = 1
                    
                    Do While SLoop < 14
                        Cells(SLoop + 2, ycolumn + 6) = Application.WorksheetFunction.Sum(Range(Cells(SLoop + 2, ycolumn + 1), Cells(SLoop + 2, ycolumn + 5)))
                        SLoop = SLoop + 1
                    Loop
                    SLoop = 1
                    GoTo NewYear2
                End If
                
            End If
        
        Loop
    
    End If

If HaveMedical = True Then
    If d = 0 And HaveVision = True Then
        GoTo VisionFill2
    End If
End If

'Build PEPY, PMPM, & Unit Cost Sheets

Dim MedMales As Integer, MedFemales As Integer, MedTotal As Integer, MedLives As Integer, VisMales As Integer, VisFemales As Integer, VisTotal As Integer, VisLives As Integer, LLoop As Integer


Sheets("Enrollment").Activate

MedMales = Range("C5")
MedFemales = Range("C6")
MedTotal = Range("C7")
MedLives = Range("D5")

VisMales = Range("H5")
VisFemales = Range("H6")
VisTotal = Range("H7")
VisLives = Range("I5")

YearCounter = NYears

ycolumn = (YearCounter - 1) * 8 + 1

xrow = 3

SLoop = 1

LLoop = 1

If HaveMedical = True Then

    Sheets("Medical Per Emp").Activate
    
    Do While YearCounter > 0
    
        Do While LLoop < 14
            
            Do While SLoop < 7
                Cells(xrow, ycolumn + 1) = Sheets("Medical Claims").Cells(xrow, ycolumn + 1) / MedTotal
                ycolumn = ycolumn + 1
                SLoop = SLoop + 1
            Loop
        
            SLoop = 1
            ycolumn = (YearCounter - 1) * 8 + 1
            xrow = xrow + 1
            LLoop = LLoop + 1
            
        Loop
        
        YearCounter = YearCounter - 1
        
        ycolumn = (YearCounter - 1) * 8 + 1
        
        xrow = 3
        
        SLoop = 1
        
        LLoop = 1
    
    Loop
    
    YearCounter = NYears
    
    ycolumn = (YearCounter - 1) * 8 + 1
    
    xrow = 3
    
    SLoop = 1
    
    LLoop = 1
    
    Sheets("Medical Per Life").Activate
    
    Do While YearCounter > 0
    
        Do While LLoop < 14
            
            Do While SLoop < 7
                Cells(xrow, ycolumn + 1) = Sheets("Medical Claims").Cells(xrow, ycolumn + 1) / MedLives
                ycolumn = ycolumn + 1
                SLoop = SLoop + 1
            Loop
        
            SLoop = 1
            ycolumn = (YearCounter - 1) * 8 + 1
            xrow = xrow + 1
            LLoop = LLoop + 1
            
        Loop
        
        YearCounter = YearCounter - 1
        
        ycolumn = (YearCounter - 1) * 8 + 1
        
        xrow = 3
        
        SLoop = 1
        
        LLoop = 1
    
    Loop

End If

YearCounter = NYears

ycolumn = (YearCounter - 1) * 8 + 1

xrow = 3

SLoop = 1

LLoop = 1

If HaveVision = True Then
    Sheets("Vision Per Emp").Activate
    
    Do While YearCounter > 0
    
        Do While LLoop < 14
            
            Do While SLoop < 7
                Cells(xrow, ycolumn + 1) = Sheets("Vision Claims").Cells(xrow, ycolumn + 1) / VisTotal
                ycolumn = ycolumn + 1
                SLoop = SLoop + 1
            Loop
        
            SLoop = 1
            ycolumn = (YearCounter - 1) * 8 + 1
            xrow = xrow + 1
            LLoop = LLoop + 1
            
        Loop
        
        YearCounter = YearCounter - 1
        
        ycolumn = (YearCounter - 1) * 8 + 1
        
        xrow = 3
        
        SLoop = 1
        
        LLoop = 1
    
    Loop
    
    YearCounter = NYears
    
    ycolumn = (YearCounter - 1) * 8 + 1
    
    xrow = 3
    
    SLoop = 1
    
    LLoop = 1
    
    Sheets("Vision Per Life").Activate
    
    Do While YearCounter > 0
    
        Do While LLoop < 14
            
            Do While SLoop < 7
                Cells(xrow, ycolumn + 1) = Sheets("Vision Claims").Cells(xrow, ycolumn + 1) / VisLives
                ycolumn = ycolumn + 1
                SLoop = SLoop + 1
            Loop
        
            SLoop = 1
            ycolumn = (YearCounter - 1) * 8 + 1
            xrow = xrow + 1
            LLoop = LLoop + 1
            
        Loop
        
        YearCounter = YearCounter - 1
        
        ycolumn = (YearCounter - 1) * 8 + 1
        
        xrow = 3
        
        SLoop = 1
        
        LLoop = 1
    
    Loop
    
End If

YearCounter = NYears

ycolumn = (YearCounter - 1) * 8 + 1

xrow = 3

SLoop = 1

LLoop = 1

If HaveMedical = True Then

    Sheets("Rx Per Emp").Activate
    
    Do While YearCounter > 0
    
        Do While LLoop < 14
            
            Do While SLoop < 5
                Cells(xrow, ycolumn + 1) = Sheets("Rx Claims").Cells(xrow, ycolumn + 1) / MedTotal
                ycolumn = ycolumn + 1
                SLoop = SLoop + 1
            Loop
        
            SLoop = 1
            ycolumn = (YearCounter - 1) * 8 + 1
            xrow = xrow + 1
            LLoop = LLoop + 1
            
        Loop
        
        YearCounter = YearCounter - 1
        
        ycolumn = (YearCounter - 1) * 8 + 1
        
        xrow = 3
        
        SLoop = 1
        
        LLoop = 1
    
    Loop
    
End If

YearCounter = NYears

ycolumn = (YearCounter - 1) * 8 + 1

xrow = 3

SLoop = 1

LLoop = 1

If HaveMedical = True Then
    
    Sheets("Rx Per Life").Activate
    
    Do While YearCounter > 0
    
        Do While LLoop < 14
            
            Do While SLoop < 5
                Cells(xrow, ycolumn + 1) = Sheets("Rx Claims").Cells(xrow, ycolumn + 1) / MedLives
                ycolumn = ycolumn + 1
                SLoop = SLoop + 1
            Loop
        
            SLoop = 1
            ycolumn = (YearCounter - 1) * 8 + 1
            xrow = xrow + 1
            LLoop = LLoop + 1
            
        Loop
        
        YearCounter = YearCounter - 1
        
        ycolumn = (YearCounter - 1) * 8 + 1
        
        xrow = 3
        
        SLoop = 1
        
        LLoop = 1
    
    Loop
    
End If

YearCounter = NYears

ycolumn = (YearCounter - 1) * 8 + 1

xrow = 3

SLoop = 1

LLoop = 1

If HaveMedical = True Then

    Sheets("Medical Unit Cost").Activate
    
    Do While YearCounter > 0
    
        Do While LLoop < 14
            
            Do While SLoop < 7
                If Sheets("Medical Count").Cells(xrow, ycolumn + 1) <> 0 Then
                    Cells(xrow, ycolumn + 1) = Sheets("Medical Claims").Cells(xrow, ycolumn + 1) / Sheets("Medical Count").Cells(xrow, ycolumn + 1)
                Else
                    Cells(xrow, ycolumn + 1) = 0
                End If
                ycolumn = ycolumn + 1
                SLoop = SLoop + 1
            Loop
        
            SLoop = 1
            ycolumn = (YearCounter - 1) * 8 + 1
            xrow = xrow + 1
            LLoop = LLoop + 1
            
        Loop
        
        YearCounter = YearCounter - 1
        
        ycolumn = (YearCounter - 1) * 8 + 1
        
        xrow = 3
        
        SLoop = 1
        
        LLoop = 1
    
    Loop
    
End If

YearCounter = NYears

ycolumn = (YearCounter - 1) * 8 + 1

xrow = 3

SLoop = 1

LLoop = 1

If HaveVision = True Then

    Sheets("Vision Unit Cost").Activate
    
    Do While YearCounter > 0
    
        Do While LLoop < 14
            
            Do While SLoop < 7
                If Sheets("Vision Count").Cells(xrow, ycolumn + 1) <> 0 Then
                    Cells(xrow, ycolumn + 1) = Sheets("Vision Claims").Cells(xrow, ycolumn + 1) / Sheets("Vision Count").Cells(xrow, ycolumn + 1)
                Else
                    Cells(xrow, ycolumn + 1) = 0
                End If
                ycolumn = ycolumn + 1
                SLoop = SLoop + 1
            Loop
        
            SLoop = 1
            ycolumn = (YearCounter - 1) * 8 + 1
            xrow = xrow + 1
            LLoop = LLoop + 1
            
        Loop
        
        YearCounter = YearCounter - 1
        
        ycolumn = (YearCounter - 1) * 8 + 1
        
        xrow = 3
        
        SLoop = 1
        
        LLoop = 1
    
    Loop

End If


Call DashboardMaker
GoTo EndNoError

Handler:
If HaveMedical = True Or HaveVision = True Then
    MsgBox ("There was an error.")
Else
    MsgBox ("You must have medical or vision coverage.")
End If

EndNoError:
End Sub

Sub DashboardMaker()

Application.DisplayAlerts = False

For Each sh In Application.Worksheets
    If sh.Name = "Dashboard" Then
        sh.Delete
    End If
Next sh
For Each sh In Application.Worksheets
    If sh.Name = "Dashboard Graphs" Then
        sh.Delete
    End If
Next sh
For Each sh In Application.Worksheets
    If sh.Name = "Medical Graphs" Then
        sh.Delete
    End If
Next sh
For Each sh In Application.Worksheets
    If sh.Name = "Rx Graphs" Then
        sh.Delete
    End If
Next sh
For Each sh In Application.Worksheets
    If sh.Name = "Vision Graphs" Then
        sh.Delete
    End If
Next sh

Sheets("Enrollment").Activate

Sheets.Add After:=ActiveSheet
ActiveSheet.Select
ActiveSheet.Name = "Dashboard"

Application.DisplayAlerts = True

Dim CompleteYears As Integer, EmpArray(2) As Variant, LifeArray(2) As Variant, UnitArray(1) As Variant, CountArray(1) As Variant, ClaimsArray(2) As Variant
Dim e As Integer, l As Integer, u As Integer, cu As Integer, cl As Integer
Dim rollingsum As Double, rollingcounter As Integer, TotalMonths As Integer, Rolltwelveyears As Integer
Dim SLoop As Integer, LLoop As Integer
Dim FinalMonthStr As String, FinalMonth As Integer, FinalYearStr As String, FinalYear As Integer, NYears As Integer, FinalMonthTemp As Integer
Dim StartMonthStr As String, StartMonth As Integer, StartYearStr As String, StartYear As Integer
Dim ycolumn As Integer, xrow As Integer, YearCounter As Integer, uprow As Integer, backcolumn As Integer, HaveMedical As Boolean, HaveVision As Boolean

e = 0
l = 0
u = 0
cu = 0
cl = 0

EmpArray(0) = "Medical Per Emp"
EmpArray(1) = "Rx Per Emp"
EmpArray(2) = "Vision Per Emp"

LifeArray(0) = "Medical Per Life"
LifeArray(1) = "Rx Per Life"
LifeArray(2) = "Vision Per Life"

UnitArray(0) = "Medical Unit Cost"
UnitArray(1) = "Vision Unit Cost"

CountArray(0) = "Medical Count"
CountArray(1) = "Vision Count"

ClaimsArray(0) = "Medical Claims"
ClaimsArray(1) = "Rx Claims"
ClaimsArray(2) = "Vision Claims"

HaveMedical = False
HaveVision = False

For Each sh In Application.Worksheets
    If sh.Name = "MDCLMS" Then
        HaveMedical = True
    End If
    If sh.Name = "VSCLMS" Then
        HaveVision = True
    End If
Next sh

If HaveMedical = True Then
    Pollsheet = "MDCLMS"
Else
    Pollsheet = "VSCLMS"
End If


PollNCells = Application.WorksheetFunction.CountA(Sheets(Pollsheet).Range("A:A")) + 1

StartYearStr = Left(Sheets(Pollsheet).Cells(5, 1), 4)
StartYear = StartYearStr

FinalYearStr = Left(Sheets(Pollsheet).Cells(PollNCells, 1), 4)
FinalYear = FinalYearStr

If Mid(Sheets(Pollsheet).Cells(5, 1), Len(Sheets(Pollsheet).Cells(5, 1)) - 1, 1) = "/" Then
    StartMonthStr = Right(Sheets(Pollsheet).Cells(5, 1), 1)
Else
    StartMonthStr = Right(Sheets(Pollsheet).Cells(5, 1), 2)
End If
StartMonth = StartMonthStr

If Mid(Sheets(Pollsheet).Cells(5, 1), Len(Sheets(Pollsheet).Cells(PollNCells, 1)) - 1, 1) = "/" Then
    FinalMonthStr = Right(Sheets(Pollsheet).Cells(5, 1), 1)
Else
    FinalMonthStr = Right(Sheets(Pollsheet).Cells(5, 1), 2)
End If
FinalMonth = FinalMonthStr

NYears = FinalYear - StartYear + 1

CompleteYears = NYears

If StartMonth <> 1 Then
    CompleteYears = CompleteYears - 1
End If
If FinalMonth <> 12 Then
    CompleteYears = CompleteYears - 1
End If

If CompleteYears > 1 Then
    If CompleteYears = NYears Then
        TotalMonths = 12 * NYears
    End If
    If StartMonth <> 1 And FinalMonth <> 12 Then
        TotalMonths = CompleteYears * 12 + FinalMonth + (13 - StartMonth)
    End If
    If StartMonth <> 1 And FinalMonth = 12 Then
        TotalMonths = CompleteYears * 12 + (13 - StartMonth)
    End If
    If StartMonth = 1 And FinalMonth <> 12 Then
        TotalMonths = CompleteYears * 12 + FinalMonth
    End If
Else
    TotalMonths = FinalMonth - (StartMonth - 1)
End If

Rolltwelveyears = Int(TotalMonths / 12)

'Graph Set 1 - consolidated line graph for PEPY by Insurance Type for complete years, rolling-12, and YTD

    'Complete Years Only
    
Sheets("Dashboard").Cells(1, 1) = "Total Paid in Claims PEPY - Complete Years Only"
    
For e = 0 To 2

    If HaveMedical = False Then
        e = 2
    End If

    YearCounter = CompleteYears
    
    ycolumn = YearCounter * 8 - 7
        
    Sheets("Dashboard").Activate
    
    Do While YearCounter > 0
        If HaveMedical = True And HaveVision = True Then
            Sheets("Dashboard").Cells(3, 1) = "Medical"
            Sheets("Dashboard").Cells(4, 1) = "Rx"
            Sheets("Dashboard").Cells(5, 1) = "Vision"
            If e <> 1 Then
                Sheets("Dashboard").Cells(3 + e, YearCounter + 1) = Sheets(EmpArray(e)).Cells(15, ycolumn + 6)
            Else
                Sheets("Dashboard").Cells(3 + e, YearCounter + 1) = Sheets(EmpArray(e)).Cells(15, ycolumn + 4)
            End If
        End If
        If HaveMedical = True And HaveVision = False Then
            Sheets("Dashboard").Cells(3, 1) = "Medical"
            Sheets("Dashboard").Cells(4, 1) = "Rx"
            If e <> 2 Then
                If e <> 1 Then
                    Sheets("Dashboard").Cells(3 + e, YearCounter + 1) = Sheets(EmpArray(e)).Cells(15, ycolumn + 6)
                Else
                    Sheets("Dashboard").Cells(3 + e, YearCounter + 1) = Sheets(EmpArray(e)).Cells(15, ycolumn + 4)
                End If
            End If
        End If
        If HaveMedical = False And HaveVision = True Then
            Sheets("Dashboard").Cells(3, 1) = "Vision"
            Sheets("Dashboard").Cells(3, YearCounter + 1) = Sheets(EmpArray(e)).Cells(15, ycolumn + 6)
        End If
        
        YearCounter = YearCounter - 1
        ycolumn = (YearCounter) * 8 - 7
    Loop

Next e

YearCounter = CompleteYears

Do While YearCounter > 0
    If StartMonth = 1 Then
        Sheets("Dashboard").Cells(2, YearCounter + 1) = FinalYear - NYears + YearCounter
    Else
        Sheets("Dashboard").Cells(2, YearCounter + 1) = FinalYear - NYears + YearCounter + 1
    End If
    YearCounter = YearCounter - 1
Loop

    'Rolling 12-months
    
Sheets("Dashboard").Cells(1, CompleteYears + 3) = "Total Paid in Claims PEPY - Rolling 12"

e = 0

For e = 0 To 2

    If HaveMedical = False Then
        e = 2
    End If
    
    rollingsum = 0
    rollingcounter = 0
    uprow = 0
    LLoop = 0
    
    YearCounter = NYears
    
    ycolumn = YearCounter * 8 - 7
    If e = 1 Then
        ycolumn = ycolumn - 2
    End If
    
    Do While FinalMonth - rollingcounter > 0
        rollingsum = rollingsum + Sheets(EmpArray(e)).Cells(2 + FinalMonth - uprow, ycolumn + 6).Value
        rollingcounter = rollingcounter + 1
        uprow = uprow + 1
    Loop
    
    uprow = 0
    YearCounter = YearCounter - 1
    ycolumn = YearCounter * 8 - 7
    If e = 1 Then
        ycolumn = ycolumn - 2
    End If
    
    Do While ycolumn > -2
    
        Do While uprow < 12
        
            Do While rollingcounter < 12 And uprow < 12
                rollingsum = rollingsum + Sheets(EmpArray(e)).Cells(14 - uprow, ycolumn + 6)
                uprow = uprow + 1
                rollingcounter = rollingcounter + 1
            Loop
            
            If rollingcounter = 12 Then
                If HaveMedical = True Then
                    Sheets("Dashboard").Cells(3 + e, CompleteYears + 3 + Rolltwelveyears - LLoop) = rollingsum
                End If
                If HaveMedical = False Then
                    Sheets("Dashboard").Cells(3, CompleteYears + 3 + Rolltwelveyears - LLoop) = rollingsum
                End If
                rollingcounter = 0
                rollingsum = 0
                LLoop = LLoop + 1
            End If
        
        Loop
        
        YearCounter = YearCounter - 1
        ycolumn = YearCounter * 8 - 7
        If e = 1 Then
            ycolumn = ycolumn - 2
        End If
        uprow = 0
        
    Loop
    
    If HaveVision = False And e = 1 Then
        e = 2
    End If

Next e

If HaveMedical = True And HaveVision = True Then
    Sheets("Dashboard").Cells(3, CompleteYears + 3) = "Medical"
    Sheets("Dashboard").Cells(4, CompleteYears + 3) = "Rx"
    Sheets("Dashboard").Cells(5, CompleteYears + 3) = "Vision"
End If
If HaveMedical = True And HaveVision = False Then
    Sheets("Dashboard").Cells(3, CompleteYears + 3) = "Medical"
    Sheets("Dashboard").Cells(4, CompleteYears + 3) = "Rx"
End If
If HaveMedical = False And HaveVision = True Then
    Sheets("Dashboard").Cells(3, CompleteYears + 3) = "Vision"
End If

SLoop = 0

FinalMonthTemp = FinalMonth

If FinalMonth = 12 Then
    FinalMonthTemp = 1
End If


If FinalMonth <> 12 Then
    Do While SLoop < Rolltwelveyears
        Sheets("Dashboard").Cells(2, CompleteYears + 3 + Rolltwelveyears - SLoop) = (FinalMonthTemp + 1) & "/" & (FinalYear - 1 - SLoop) & "-" & FinalMonth & "/" & (FinalYear - SLoop)
        SLoop = SLoop + 1
    Loop
Else
    Do While SLoop < Rolltwelveyears
        Sheets("Dashboard").Cells(2, CompleteYears + 3 + Rolltwelveyears - SLoop) = 1 & "/" & (FinalYear - SLoop) & "-" & (FinalMonth) & "/" & (FinalYear - SLoop)
        SLoop = SLoop + 1
    Loop
End If

    'Year-to-Date
    
Sheets("Dashboard").Cells(1, CompleteYears + Rolltwelveyears + 5) = "Total Paid in Claims PEPY - Year-to-Date"
If CompleteYears <> 0 Then
Sheets("Dashboard").Cells(2, CompleteYears + Rolltwelveyears + 6) = "'" & 1 & "-" & FinalMonth & "/" & FinalYear
Else
Sheets("Dashboard").Cells(2, CompleteYears + Rolltwelveyears + 6) = "'" & StartMonth & "-" & FinalMonth & "/" & FinalYear
End If

e = 0

ycolumn = NYears * 8 - 7

For e = 0 To 2

    If HaveMedical = False Then
        e = 2
    End If

    If HaveMedical = True And HaveVision = True Then
        Sheets("Dashboard").Cells(3, CompleteYears + Rolltwelveyears + 5) = "Medical"
        Sheets("Dashboard").Cells(4, CompleteYears + Rolltwelveyears + 5) = "Rx"
        Sheets("Dashboard").Cells(5, CompleteYears + Rolltwelveyears + 5) = "Vision"
        If e <> 1 Then
            Sheets("Dashboard").Cells(3 + e, CompleteYears + Rolltwelveyears + 6) = Sheets(EmpArray(e)).Cells(15, ycolumn + 6)
        Else
            Sheets("Dashboard").Cells(3 + e, CompleteYears + Rolltwelveyears + 6) = Sheets(EmpArray(e)).Cells(15, ycolumn + 4)
        End If
    End If
    If HaveMedical = True And HaveVision = False Then
        Sheets("Dashboard").Cells(3, CompleteYears + Rolltwelveyears + 5) = "Medical"
        Sheets("Dashboard").Cells(4, CompleteYears + Rolltwelveyears + 5) = "Rx"
        If e <> 2 Then
            If e <> 1 Then
                Sheets("Dashboard").Cells(3 + e, CompleteYears + Rolltwelveyears + 6) = Sheets(EmpArray(e)).Cells(15, ycolumn + 6)
            Else
                Sheets("Dashboard").Cells(3 + e, CompleteYears + Rolltwelveyears + 6) = Sheets(EmpArray(e)).Cells(15, ycolumn + 4)
            End If
        End If
    End If
    If HaveMedical = False And HaveVision = True Then
        Sheets("Dashboard").Cells(3, CompleteYears + Rolltwelveyears + 5) = "Vision"
        Sheets("Dashboard").Cells(3, CompleteYears + Rolltwelveyears + 6) = Sheets(EmpArray(e)).Cells(15, ycolumn + 6)
    End If

Next e

'New MESSA Graph Data for revision of 1-6-2019
Sheets("Dashboard").Cells(1, CompleteYears + Rolltwelveyears + 8) = "Total Paid in Claims PEPY - Complete Years and Rolling-12"
Sheets("Dashboard").Range(Cells(2, 1), Cells(5, CompleteYears + 1)).Select
Selection.Copy
Sheets("Dashboard").Cells(2, CompleteYears + Rolltwelveyears + 8).PasteSpecial
Sheets("Dashboard").Range(Cells(2, CompleteYears + Rolltwelveyears + 3), Cells(5, CompleteYears + Rolltwelveyears + 3)).Select
Selection.Copy
Sheets("Dashboard").Cells(2, CompleteYears * 2 + Rolltwelveyears + 9).PasteSpecial

'Graph Set 2 - consolidated line graph for PMPM by Insurance Type for Past 12 Months and Total

    'Past 12 Months

If CompleteYears <> 0 Then

    Sheets("Dashboard").Cells(7, 1) = "Total Paid in Claims PMPM - Past 12 Months"
    
    l = 0
    
    For l = 0 To 2
    
        If HaveMedical = False Then
            l = 2
        End If
    
        rollingcounter = 0
        uprow = 0
        FinalMonthTemp = FinalMonth
        YearCounter = NYears
        ycolumn = YearCounter * 8 - 7
        If l = 1 Then
            ycolumn = ycolumn - 2
        End If
        
        Do While rollingcounter < 12
        
            Do While FinalMonthTemp - uprow > 0 And rollingcounter < 12
            
                If HaveMedical = True And HaveVision = True Then
                    Sheets("Dashboard").Cells(9 + l, 13 - rollingcounter) = Sheets(LifeArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6).Value
                End If
                If HaveMedical = False And HaveVision = True Then
                    Sheets("Dashboard").Cells(9, 13 - rollingcounter) = Sheets(LifeArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6).Value
                End If
                If HaveMedical = True And HaveVision = False Then
                    If l = 2 Then
                    
                    Else
                        Sheets("Dashboard").Cells(9 + l, 13 - rollingcounter) = Sheets(LifeArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6).Value
                    End If
                End If
                    
                rollingcounter = rollingcounter + 1
                uprow = uprow + 1
    
            Loop
            
            FinalMonthTemp = 12
            uprow = 0
            YearCounter = YearCounter - 1
            ycolumn = YearCounter * 8 - 7
            If l = 1 Then
                ycolumn = ycolumn - 2
            End If
            
        Loop
        
        If HaveVision = False And l = 1 Then
            l = 2
        End If
    
    Next l
    
    rollingcounter = 0
    uprow = 0
    FinalMonthTemp = FinalMonth
    YearCounter = NYears
    ycolumn = YearCounter * 8 - 7
    
    If HaveMedical = True Then
        l = 0
    Else
        l = 2
    End If
    
    Do While rollingcounter < 12
        
        Do While FinalMonthTemp - uprow > 0 And rollingcounter < 12
        
            Sheets("Dashboard").Cells(8, 13 - rollingcounter) = "'" & Sheets(LifeArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn).Value & "/" & Sheets(LifeArray(l)).Cells(1, ycolumn).Value
                
            rollingcounter = rollingcounter + 1
            uprow = uprow + 1
    
        Loop
        
        FinalMonthTemp = 12
        uprow = 0
        YearCounter = YearCounter - 1
        ycolumn = YearCounter * 8 - 7
            
    Loop
    
    
    If HaveMedical = True And HaveVision = True Then
        Sheets("Dashboard").Cells(9, 1) = "Medical"
        Sheets("Dashboard").Cells(10, 1) = "Rx"
        Sheets("Dashboard").Cells(11, 1) = "Vision"
    End If
    If HaveMedical = True And HaveVision = False Then
        Sheets("Dashboard").Cells(9, 1) = "Medical"
        Sheets("Dashboard").Cells(10, 1) = "Rx"
    End If
    If HaveMedical = False And HaveVision = True Then
        Sheets("Dashboard").Cells(9, 1) = "Vision"
    End If
    
End If

    'Total
    
Sheets("Dashboard").Cells(7, 15) = "Total Paid in Claims PMPM - Total"

l = 0

For l = 0 To 2

    If HaveMedical = False Then
        l = 2
    End If

    rollingcounter = 0
    uprow = 0
    FinalMonthTemp = FinalMonth
    YearCounter = NYears
    ycolumn = YearCounter * 8 - 7
    If l = 1 Then
        ycolumn = ycolumn - 2
    End If
    
    Do While rollingcounter < TotalMonths
    
        Do While FinalMonthTemp - uprow > 0 And rollingcounter < TotalMonths
        
            If HaveMedical = True And HaveVision = True Then
                Sheets("Dashboard").Cells(9 + l, 15 + TotalMonths - rollingcounter) = Sheets(LifeArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6).Value
            End If
            If HaveMedical = False And HaveVision = True Then
                Sheets("Dashboard").Cells(9, 15 + TotalMonths - rollingcounter) = Sheets(LifeArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6).Value
            End If
            If HaveMedical = True And HaveVision = False Then
                If l = 2 Then
                
                Else
                    Sheets("Dashboard").Cells(9 + l, 15 + TotalMonths - rollingcounter) = Sheets(LifeArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6).Value
                End If
            End If
                
            rollingcounter = rollingcounter + 1
            uprow = uprow + 1

        Loop
        
        FinalMonthTemp = 12
        uprow = 0
        YearCounter = YearCounter - 1
        ycolumn = YearCounter * 8 - 7
        If l = 1 Then
            ycolumn = ycolumn - 2
        End If
        
    Loop
    
    If HaveVision = False And l = 1 Then
        l = 2
    End If

Next l

rollingcounter = 0
uprow = 0
FinalMonthTemp = FinalMonth
YearCounter = NYears
ycolumn = YearCounter * 8 - 7

If HaveMedical = True Then
    l = 0
Else
    l = 2
End If

Do While rollingcounter < TotalMonths
    
    Do While FinalMonthTemp - uprow > 0 And rollingcounter < TotalMonths

        Sheets("Dashboard").Cells(8, 15 + TotalMonths - rollingcounter) = "'" & Sheets(LifeArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn).Value & "/" & Sheets(LifeArray(l)).Cells(1, ycolumn).Value
                
        rollingcounter = rollingcounter + 1
        uprow = uprow + 1

    Loop
    
    FinalMonthTemp = 12
    uprow = 0
    YearCounter = YearCounter - 1
    ycolumn = YearCounter * 8 - 7
        
Loop


If HaveMedical = True And HaveVision = True Then
    Sheets("Dashboard").Cells(9, 15) = "Medical"
    Sheets("Dashboard").Cells(10, 15) = "Rx"
    Sheets("Dashboard").Cells(11, 15) = "Vision"
End If
If HaveMedical = True And HaveVision = False Then
    Sheets("Dashboard").Cells(9, 15) = "Medical"
    Sheets("Dashboard").Cells(10, 15) = "Rx"
End If
If HaveMedical = False And HaveVision = True Then
    Sheets("Dashboard").Cells(9, 15) = "Vision"
End If

'Graph Set 3 - consolidated line graphs for unit cost per month by Claim Type for Insurance Types for total and past 12 months

    'Past 12 Months - Totals
    
If CompleteYears <> 0 Then

    Sheets("Dashboard").Cells(13, 1) = "Unit Cost Paid in Claims - Past 12 Months"
    
    u = 0
    
    For u = 0 To 1
    
        If HaveMedical = False Then
            u = 1
        End If
    
        rollingcounter = 0
        uprow = 0
        FinalMonthTemp = FinalMonth
        YearCounter = NYears
        ycolumn = YearCounter * 8 - 7
        
        Do While rollingcounter < 12
        
            Do While FinalMonthTemp - uprow > 0 And rollingcounter < 12
            
                If HaveMedical = True And HaveVision = True Then
                    Sheets("Dashboard").Cells(15 + u, 13 - rollingcounter) = Sheets(UnitArray(u)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6).Value
                End If
                If HaveMedical = False And HaveVision = True Then
                    Sheets("Dashboard").Cells(15, 13 - rollingcounter) = Sheets(UnitArray(u)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6).Value
                End If
                If HaveMedical = True And HaveVision = False Then
                    If u = 1 Then
                    
                    Else
                        Sheets("Dashboard").Cells(15, 13 - rollingcounter) = Sheets(UnitArray(u)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6).Value
                    End If
                End If
                    
                rollingcounter = rollingcounter + 1
                uprow = uprow + 1
    
            Loop
            
            FinalMonthTemp = 12
            uprow = 0
            YearCounter = YearCounter - 1
            ycolumn = YearCounter * 8 - 7
            
        Loop
    
    Next u
    
    rollingcounter = 0
    uprow = 0
    FinalMonthTemp = FinalMonth
    YearCounter = NYears
    ycolumn = YearCounter * 8 - 7
    
    If HaveMedical = True Then
        u = 0
    Else
        u = 1
    End If
    
    Do While rollingcounter < 12
        
        Do While FinalMonthTemp - uprow > 0 And rollingcounter < 12
        
            Sheets("Dashboard").Cells(14, 13 - rollingcounter) = "'" & Sheets(UnitArray(u)).Cells(2 + FinalMonthTemp - uprow, ycolumn).Value & "/" & Sheets(UnitArray(u)).Cells(1, ycolumn).Value
                
            rollingcounter = rollingcounter + 1
            uprow = uprow + 1
    
        Loop
        
        FinalMonthTemp = 12
        uprow = 0
        YearCounter = YearCounter - 1
        ycolumn = YearCounter * 8 - 7
            
    Loop
    
    
    If HaveMedical = True And HaveVision = True Then
        Sheets("Dashboard").Cells(15, 1) = "Medical"
        Sheets("Dashboard").Cells(16, 1) = "Vision"
    End If
    If HaveMedical = True And HaveVision = False Then
        Sheets("Dashboard").Cells(15, 1) = "Medical"
    End If
    If HaveMedical = False And HaveVision = True Then
        Sheets("Dashboard").Cells(15, 1) = "Vision"
    End If
    
End If

    'Total - Totals
    
Sheets("Dashboard").Cells(13, 15) = "Unit Cost Paid in Claims - Total"

u = 0

For u = 0 To 1

    If HaveMedical = False Then
        u = 1
    End If

    rollingcounter = 0
    uprow = 0
    FinalMonthTemp = FinalMonth
    YearCounter = NYears
    ycolumn = YearCounter * 8 - 7
    
    Do While rollingcounter < TotalMonths
    
        Do While FinalMonthTemp - uprow > 0 And rollingcounter < TotalMonths
        
            If HaveMedical = True And HaveVision = True Then
                Sheets("Dashboard").Cells(15 + u, 15 + TotalMonths - rollingcounter) = Sheets(UnitArray(u)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6).Value
            End If
            If HaveMedical = False And HaveVision = True Then
                Sheets("Dashboard").Cells(15, 15 + TotalMonths - rollingcounter) = Sheets(UnitArray(u)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6).Value
            End If
            If HaveMedical = True And HaveVision = False Then
                If u = 1 Then
                
                Else
                    Sheets("Dashboard").Cells(15, 15 + TotalMonths - rollingcounter) = Sheets(UnitArray(u)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6).Value
                End If
            End If
                
            rollingcounter = rollingcounter + 1
            uprow = uprow + 1

        Loop
        
        FinalMonthTemp = 12
        uprow = 0
        YearCounter = YearCounter - 1
        ycolumn = YearCounter * 8 - 7
        
    Loop

Next u

rollingcounter = 0
uprow = 0
FinalMonthTemp = FinalMonth
YearCounter = NYears
ycolumn = YearCounter * 8 - 7

If HaveMedical = True Then
    u = 0
Else
    u = 1
End If

Do While rollingcounter < TotalMonths
    
    Do While FinalMonthTemp - uprow > 0 And rollingcounter < TotalMonths

        Sheets("Dashboard").Cells(14, 15 + TotalMonths - rollingcounter) = "'" & Sheets(UnitArray(u)).Cells(2 + FinalMonthTemp - uprow, ycolumn).Value & "/" & Sheets(UnitArray(u)).Cells(1, ycolumn).Value
                
        rollingcounter = rollingcounter + 1
        uprow = uprow + 1

    Loop
    
    FinalMonthTemp = 12
    uprow = 0
    YearCounter = YearCounter - 1
    ycolumn = YearCounter * 8 - 7
        
Loop

If HaveMedical = True And HaveVision = True Then
    Sheets("Dashboard").Cells(15, 15) = "Medical"
    Sheets("Dashboard").Cells(16, 15) = "Vision"
End If
If HaveMedical = True And HaveVision = False Then
    Sheets("Dashboard").Cells(15, 15) = "Medical"
End If
If HaveMedical = False And HaveVision = True Then
    Sheets("Dashboard").Cells(15, 15) = "Vision"
End If

    'Past 12 Months - Medical and Vision
    
xrow = 0
backcolumn = 1

If CompleteYears <> 0 Then

    If HaveMedical = True And HaveVision = True Then
        Sheets("Dashboard").Cells(18, 1) = "Unit Cost Paid in Medical Claims by Type - Past 12 Months"
        Sheets("Dashboard").Cells(26, 1) = "Unit Cost Paid in Vision Claims by Type - Past 12 Months"
    End If
    
    If HaveMedical = True And HaveVision = False Then
        Sheets("Dashboard").Cells(18, 1) = "Unit Cost Paid in Medical Claims by Type - Past 12 Months"
    End If
    
    If HaveMedical = False And HaveVision = True Then
        Sheets("Dashboard").Cells(26, 1) = "Unit Cost Paid in Vision Claims by Type - Past 12 Months"
    End If
    
    u = 0
    
    For u = 0 To 1
    
        If (HaveMedical = True And u = 0) Or (HaveVision = True And u = 1) Then
        
            If HaveMedical = False Then
                u = 1
                xrow = xrow + 7
            End If
                
            Do While backcolumn < 6
            
                rollingcounter = 0
                uprow = 0
                FinalMonthTemp = FinalMonth
                YearCounter = NYears
                ycolumn = YearCounter * 8 - 7
                
                    Do While rollingcounter < 12
                    
                        Do While FinalMonthTemp - uprow > 0 And rollingcounter < 12
            
                            Sheets("Dashboard").Cells(20 + u + xrow, 13 - rollingcounter) = Sheets(UnitArray(u)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6 - backcolumn).Value
                                
                            rollingcounter = rollingcounter + 1
                            uprow = uprow + 1
                
                        Loop
                        
                        FinalMonthTemp = 12
                        uprow = 0
                        YearCounter = YearCounter - 1
                        ycolumn = YearCounter * 8 - 7
                        
                    Loop
            
                backcolumn = backcolumn + 1
                xrow = xrow + 1
            
            Loop
        
        xrow = xrow + 2
        backcolumn = 1
        
        End If
        
    Next u
    
    rollingcounter = 0
    uprow = 0
    FinalMonthTemp = FinalMonth
    YearCounter = NYears
    ycolumn = YearCounter * 8 - 7
    xrow = 0
    SLoop = 0
    
    If HaveMedical = True Then
        u = 0
    Else
        u = 1
    End If
    
    Do While SLoop < 2
    
        If HaveMedical = False Then
            SLoop = 1
            xrow = xrow + 8
        End If
        
        If HaveVision = False Then
            SLoop = 1
        End If
    
        Do While rollingcounter < 12
            
            Do While FinalMonthTemp - uprow > 0 And rollingcounter < 12
            
                Sheets("Dashboard").Cells(19 + xrow, 13 - rollingcounter) = "'" & Sheets(UnitArray(u)).Cells(2 + FinalMonthTemp - uprow, ycolumn).Value & "/" & Sheets(UnitArray(u)).Cells(1, ycolumn).Value
                    
                rollingcounter = rollingcounter + 1
                uprow = uprow + 1
        
            Loop
            
            FinalMonthTemp = 12
            uprow = 0
            YearCounter = YearCounter - 1
            ycolumn = YearCounter * 8 - 7
            
        Loop
        
        xrow = xrow + 8
        SLoop = SLoop + 1
        
        rollingcounter = 0
        uprow = 0
        FinalMonthTemp = FinalMonth
        YearCounter = NYears
        ycolumn = YearCounter * 8 - 7
    
    Loop
    
    If HaveMedical = True Then
        Sheets("Dashboard").Cells(20, 1) = "Out-Patient"
        Sheets("Dashboard").Cells(21, 1) = "Other Equipment"
        Sheets("Dashboard").Cells(22, 1) = "Medical/Surgical"
        Sheets("Dashboard").Cells(23, 1) = "Lab/X-Ray"
        Sheets("Dashboard").Cells(24, 1) = "In-Patient"
    End If
    
    If HaveVision = True Then
        Sheets("Dashboard").Cells(28, 1) = "Miscellaneous"
        Sheets("Dashboard").Cells(29, 1) = "Lenses"
        Sheets("Dashboard").Cells(30, 1) = "Frames"
        Sheets("Dashboard").Cells(31, 1) = "Exam"
        Sheets("Dashboard").Cells(32, 1) = "Contact Lenses"
    End If
    
End If

    'Total - Medical and Vision
    
Sheets("Dashboard").Cells(13, 15) = "Unit Cost Paid in Claims - Total"

If HaveMedical = True And HaveVision = True Then
    Sheets("Dashboard").Cells(18, 15) = "Unit Cost Paid in Medical Claims by Type - Total"
    Sheets("Dashboard").Cells(26, 15) = "Unit Cost Paid in Vision Claims by Type - Total"
End If

If HaveMedical = True And HaveVision = False Then
    Sheets("Dashboard").Cells(18, 15) = "Unit Cost Paid in Medical Claims by Type - Total"
End If

If HaveMedical = False And HaveVision = True Then
    Sheets("Dashboard").Cells(26, 15) = "Unit Cost Paid in Vision Claims by Type - Total"
End If

xrow = 0
backcolumn = 1
u = 0

For u = 0 To 1

    If (HaveMedical = True And u = 0) Or (HaveVision = True And u = 1) Then

        If HaveMedical = False Then
            u = 1
            xrow = xrow + 7
        End If
        
        Do While backcolumn < 6
    
            rollingcounter = 0
            uprow = 0
            FinalMonthTemp = FinalMonth
            YearCounter = NYears
            ycolumn = YearCounter * 8 - 7
            
            Do While rollingcounter < TotalMonths
            
                Do While FinalMonthTemp - uprow > 0 And rollingcounter < TotalMonths
                
                    Sheets("Dashboard").Cells(20 + u + xrow, 15 + TotalMonths - rollingcounter) = Sheets(UnitArray(u)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6 - backcolumn).Value
                    
                    rollingcounter = rollingcounter + 1
                    uprow = uprow + 1
        
                Loop
                
                FinalMonthTemp = 12
                uprow = 0
                YearCounter = YearCounter - 1
                ycolumn = YearCounter * 8 - 7
                
            Loop
            
            backcolumn = backcolumn + 1
            xrow = xrow + 1
        
        Loop
        
        backcolumn = 1
        xrow = xrow + 2
    
    End If
    
Next u

rollingcounter = 0
uprow = 0
FinalMonthTemp = FinalMonth
YearCounter = NYears
ycolumn = YearCounter * 8 - 7
xrow = 0
SLoop = 0

If HaveMedical = True Then
    u = 0
Else
    u = 1
End If

Do While SLoop < 2
    
        If HaveMedical = False Then
            SLoop = 1
            xrow = xrow + 8
        End If
        
        If HaveVision = False Then
            SLoop = 1
        End If

    Do While rollingcounter < TotalMonths
        
        Do While FinalMonthTemp - uprow > 0 And rollingcounter < TotalMonths
    
            Sheets("Dashboard").Cells(19 + xrow, 15 + TotalMonths - rollingcounter) = "'" & Sheets(UnitArray(u)).Cells(2 + FinalMonthTemp - uprow, ycolumn).Value & "/" & Sheets(UnitArray(u)).Cells(1, ycolumn).Value
                    
            rollingcounter = rollingcounter + 1
            uprow = uprow + 1
    
        Loop
        
        FinalMonthTemp = 12
        uprow = 0
        YearCounter = YearCounter - 1
        ycolumn = YearCounter * 8 - 7
            
    Loop
    
    xrow = xrow + 8
    SLoop = SLoop + 1
    
    rollingcounter = 0
    uprow = 0
    FinalMonthTemp = FinalMonth
    YearCounter = NYears
    ycolumn = YearCounter * 8 - 7
    
Loop

If HaveMedical = True Then
    Sheets("Dashboard").Cells(20, 15) = "Out-Patient"
    Sheets("Dashboard").Cells(21, 15) = "Other Equipment"
    Sheets("Dashboard").Cells(22, 15) = "Medical/Surgical"
    Sheets("Dashboard").Cells(23, 15) = "Lab/X-Ray"
    Sheets("Dashboard").Cells(24, 15) = "In-Patient"
End If

If HaveVision = True Then
    Sheets("Dashboard").Cells(28, 15) = "Miscellaneous"
    Sheets("Dashboard").Cells(29, 15) = "Lenses"
    Sheets("Dashboard").Cells(30, 15) = "Frames"
    Sheets("Dashboard").Cells(31, 15) = "Exam"
    Sheets("Dashboard").Cells(32, 15) = "Contact Lenses"
End If

'Graph Set 4 - consolidated bar graphs for spending by each Insurance Type for complete year, rolling-12, and YTD

    'Complete Years Only
    
Sheets("Dashboard").Cells(34, 1) = "Total Paid in Claims - Complete Years Only"
    
For cl = 0 To 2

    If HaveMedical = False Then
        cl = 2
    End If

    YearCounter = CompleteYears
    
    ycolumn = YearCounter * 8 - 7
        
    Sheets("Dashboard").Activate
    
    Do While YearCounter > 0
        If HaveMedical = True And HaveVision = True Then
            Sheets("Dashboard").Cells(36, 1) = "Medical"
            Sheets("Dashboard").Cells(37, 1) = "Rx"
            Sheets("Dashboard").Cells(38, 1) = "Vision"
            If cl <> 1 Then
                Sheets("Dashboard").Cells(36 + cl, YearCounter + 1) = Sheets(ClaimsArray(cl)).Cells(15, ycolumn + 6)
            Else
                Sheets("Dashboard").Cells(36 + cl, YearCounter + 1) = Sheets(ClaimsArray(cl)).Cells(15, ycolumn + 4)
            End If
        End If
        If HaveMedical = True And HaveVision = False Then
            Sheets("Dashboard").Cells(36, 1) = "Medical"
            Sheets("Dashboard").Cells(37, 1) = "Rx"
            If cl <> 2 Then
                If cl <> 1 Then
                    Sheets("Dashboard").Cells(36 + cl, YearCounter + 1) = Sheets(ClaimsArray(cl)).Cells(15, ycolumn + 6)
                Else
                    Sheets("Dashboard").Cells(36 + cl, YearCounter + 1) = Sheets(ClaimsArray(cl)).Cells(15, ycolumn + 4)
                End If
            End If
        End If
        If HaveMedical = False And HaveVision = True Then
            Sheets("Dashboard").Cells(36, 1) = "Vision"
            Sheets("Dashboard").Cells(36, YearCounter + 1) = Sheets(ClaimsArray(cl)).Cells(15, ycolumn + 6)
        End If
        
        YearCounter = YearCounter - 1
        ycolumn = (YearCounter) * 8 - 7
    Loop

Next cl

YearCounter = CompleteYears

Do While YearCounter > 0
    If StartMonth = 1 Then
        Sheets("Dashboard").Cells(35, YearCounter + 1) = FinalYear - NYears + YearCounter
    Else
        Sheets("Dashboard").Cells(35, YearCounter + 1) = FinalYear - NYears + YearCounter + 1
    End If
    YearCounter = YearCounter - 1
Loop

    'Rolling 12-months
    
Sheets("Dashboard").Cells(34, CompleteYears + 3) = "Total Paid in Claims - Rolling 12"

cl = 0

For cl = 0 To 2

    If HaveMedical = False Then
        cl = 2
    End If
    
    rollingsum = 0
    rollingcounter = 0
    uprow = 0
    LLoop = 0
    
    YearCounter = NYears
    
    ycolumn = YearCounter * 8 - 7
    If cl = 1 Then
        ycolumn = ycolumn - 2
    End If
    
    Do While FinalMonth - rollingcounter > 0
        rollingsum = rollingsum + Sheets(ClaimsArray(cl)).Cells(2 + FinalMonth - uprow, ycolumn + 6).Value
        rollingcounter = rollingcounter + 1
        uprow = uprow + 1
    Loop
    
    uprow = 0
    YearCounter = YearCounter - 1
    ycolumn = YearCounter * 8 - 7
    If cl = 1 Then
        ycolumn = ycolumn - 2
    End If
    
    Do While ycolumn > -2
    
        Do While uprow < 12
        
            Do While rollingcounter < 12 And uprow < 12
                rollingsum = rollingsum + Sheets(ClaimsArray(cl)).Cells(14 - uprow, ycolumn + 6)
                uprow = uprow + 1
                rollingcounter = rollingcounter + 1
            Loop
            
            If rollingcounter = 12 Then
                If HaveMedical = True Then
                    Sheets("Dashboard").Cells(36 + cl, CompleteYears + 3 + Rolltwelveyears - LLoop) = rollingsum
                End If
                If HaveMedical = False Then
                    Sheets("Dashboard").Cells(36, CompleteYears + 3 + Rolltwelveyears - LLoop) = rollingsum
                End If
                rollingcounter = 0
                rollingsum = 0
                LLoop = LLoop + 1
            End If
        
        Loop
        
        YearCounter = YearCounter - 1
        ycolumn = YearCounter * 8 - 7
        If cl = 1 Then
            ycolumn = ycolumn - 2
        End If
        uprow = 0
        
    Loop
    
    If HaveVision = False And cl = 1 Then
        cl = 2
    End If

Next cl

If HaveMedical = True And HaveVision = True Then
    Sheets("Dashboard").Cells(36, CompleteYears + 3) = "Medical"
    Sheets("Dashboard").Cells(37, CompleteYears + 3) = "Rx"
    Sheets("Dashboard").Cells(38, CompleteYears + 3) = "Vision"
End If
If HaveMedical = True And HaveVision = False Then
    Sheets("Dashboard").Cells(36, CompleteYears + 3) = "Medical"
    Sheets("Dashboard").Cells(37, CompleteYears + 3) = "Rx"
End If
If HaveMedical = False And HaveVision = True Then
    Sheets("Dashboard").Cells(36, CompleteYears + 3) = "Vision"
End If

SLoop = 0

FinalMonthTemp = FinalMonth

If FinalMonth = 12 Then
    FinalMonthTemp = 1
End If


If FinalMonth <> 12 Then
    Do While SLoop < Rolltwelveyears
        Sheets("Dashboard").Cells(35, CompleteYears + 3 + Rolltwelveyears - SLoop) = (FinalMonthTemp + 1) & "/" & (FinalYear - 1 - SLoop) & "-" & FinalMonth & "/" & (FinalYear - SLoop)
        SLoop = SLoop + 1
    Loop
Else
    Do While SLoop < Rolltwelveyears
        Sheets("Dashboard").Cells(35, CompleteYears + 3 + Rolltwelveyears - SLoop) = 1 & "/" & (FinalYear - SLoop) & "-" & (FinalMonth) & "/" & (FinalYear - SLoop)
        SLoop = SLoop + 1
    Loop
End If

    'Year-to-Date
    
Sheets("Dashboard").Cells(34, CompleteYears + Rolltwelveyears + 5) = "Total Paid in Claims - Year-to-Date"
If CompleteYears <> 0 Then
Sheets("Dashboard").Cells(35, CompleteYears + Rolltwelveyears + 6) = "'" & 1 & "-" & FinalMonth & "/" & FinalYear
Else
Sheets("Dashboard").Cells(35, CompleteYears + Rolltwelveyears + 6) = "'" & StartMonth & "-" & FinalMonth & "/" & FinalYear
End If

cl = 0

ycolumn = NYears * 8 - 7

For cl = 0 To 2

    If HaveMedical = False Then
        cl = 2
    End If

    If HaveMedical = True And HaveVision = True Then
        Sheets("Dashboard").Cells(36, CompleteYears + Rolltwelveyears + 5) = "Medical"
        Sheets("Dashboard").Cells(37, CompleteYears + Rolltwelveyears + 5) = "Rx"
        Sheets("Dashboard").Cells(38, CompleteYears + Rolltwelveyears + 5) = "Vision"
        If cl <> 1 Then
            Sheets("Dashboard").Cells(36 + cl, CompleteYears + Rolltwelveyears + 6) = Sheets(ClaimsArray(cl)).Cells(15, ycolumn + 6)
        Else
            Sheets("Dashboard").Cells(36 + cl, CompleteYears + Rolltwelveyears + 6) = Sheets(ClaimsArray(cl)).Cells(15, ycolumn + 4)
        End If
    End If
    If HaveMedical = True And HaveVision = False Then
        Sheets("Dashboard").Cells(36, CompleteYears + Rolltwelveyears + 5) = "Medical"
        Sheets("Dashboard").Cells(37, CompleteYears + Rolltwelveyears + 5) = "Rx"
        If cl <> 2 Then
            If cl <> 1 Then
                Sheets("Dashboard").Cells(36 + cl, CompleteYears + Rolltwelveyears + 6) = Sheets(ClaimsArray(cl)).Cells(15, ycolumn + 6)
            Else
                Sheets("Dashboard").Cells(36 + cl, CompleteYears + Rolltwelveyears + 6) = Sheets(ClaimsArray(cl)).Cells(15, ycolumn + 4)
            End If
        End If
    End If
    If HaveMedical = False And HaveVision = True Then
        Sheets("Dashboard").Cells(36, CompleteYears + Rolltwelveyears + 5) = "Vision"
        Sheets("Dashboard").Cells(36, CompleteYears + Rolltwelveyears + 6) = Sheets(ClaimsArray(cl)).Cells(15, ycolumn + 6)
    End If

Next cl

cl = 0

'New MESSA Graph Data for revision of 1/6/2019
Sheets("Dashboard").Cells(34, CompleteYears + Rolltwelveyears + 8) = "Total Paid in Claims - Complete Years and Rolling-12"
Sheets("Dashboard").Range(Cells(35, 1), Cells(38, CompleteYears + 1)).Select
Selection.Copy
Sheets("Dashboard").Cells(35, CompleteYears + Rolltwelveyears + 8).PasteSpecial
Sheets("Dashboard").Range(Cells(35, CompleteYears + Rolltwelveyears + 3), Cells(38, CompleteYears + Rolltwelveyears + 3)).Select
Selection.Copy
Sheets("Dashboard").Cells(35, CompleteYears * 2 + Rolltwelveyears + 9).PasteSpecial
Sheets("Dashboard").Range(Cells(35, CompleteYears + Rolltwelveyears + 8), Cells(35, CompleteYears * 2 + Rolltwelveyears + 8)).Select
Selection.NumberFormat = "General"


'Graph Set 5 - consolidated bar graphs for spending by each Claim Type for Insurance Types for complete year, rolling-12, and YTD

    'Complete Years Only
    
cl = 0

If HaveMedical = True Then
    
    Sheets("Dashboard").Cells(40, 1) = "Total Paid in Medical Claims by Type - Complete Years Only"
    Sheets("Dashboard").Cells(48, 1) = "Total Paid in Rx Claims by Type - Complete Years Only"
    Sheets("Dashboard").Cells(42, 1) = "Out-Patient"
    Sheets("Dashboard").Cells(43, 1) = "Other Equipment"
    Sheets("Dashboard").Cells(44, 1) = "Medical/Surgical"
    Sheets("Dashboard").Cells(45, 1) = "Lab/X-Ray"
    Sheets("Dashboard").Cells(46, 1) = "In-Patient"
    Sheets("Dashboard").Cells(50, 1) = "Specialty"
    Sheets("Dashboard").Cells(51, 1) = "Generic"
    Sheets("Dashboard").Cells(52, 1) = "Brand"
    
    YearCounter = CompleteYears

    Do While YearCounter > 0
        If StartMonth = 1 Then
            Sheets("Dashboard").Cells(41, YearCounter + 1) = FinalYear - NYears + YearCounter
            Sheets("Dashboard").Cells(49, YearCounter + 1) = FinalYear - NYears + YearCounter
        Else
            Sheets("Dashboard").Cells(41, YearCounter + 1) = FinalYear - NYears + YearCounter + 1
            Sheets("Dashboard").Cells(49, YearCounter + 1) = FinalYear - NYears + YearCounter + 1
        End If
        YearCounter = YearCounter - 1
    Loop
    
End If

If HaveVision = True Then
    Sheets("Dashboard").Cells(54, 1) = "Total Paid in Vision Claims by Type - Complete Years Only"
    Sheets("Dashboard").Cells(56, 1) = "Miscellaneous"
    Sheets("Dashboard").Cells(57, 1) = "Lenses"
    Sheets("Dashboard").Cells(58, 1) = "Frames"
    Sheets("Dashboard").Cells(59, 1) = "Exam"
    Sheets("Dashboard").Cells(60, 1) = "Contact Lenses"
    
    YearCounter = CompleteYears

    Do While YearCounter > 0
        If StartMonth = 1 Then
            Sheets("Dashboard").Cells(55, YearCounter + 1) = FinalYear - NYears + YearCounter
        Else
            Sheets("Dashboard").Cells(55, YearCounter + 1) = FinalYear - NYears + YearCounter + 1
        End If
        YearCounter = YearCounter - 1
    Loop
    
End If

For cl = 0 To 2

    If HaveMedical = False Then
        cl = 2
    End If

    YearCounter = CompleteYears
    
    ycolumn = YearCounter * 8 - 7
    
    Do While YearCounter > 0
        If cl <> 1 Then
        
            backcolumn = 1
            
            If cl = 0 Then
                xrow = 0
            End If

            If cl = 2 Then
                xrow = 14
            End If
            
            Do While backcolumn < 6
                Sheets("Dashboard").Cells(42 + xrow, YearCounter + 1) = Sheets(ClaimsArray(cl)).Cells(15, ycolumn + 6 - backcolumn)
                backcolumn = backcolumn + 1
                xrow = xrow + 1
            Loop
            
        Else
        
            backcolumn = 1
            xrow = 8
            
            Do While backcolumn < 4
                Sheets("Dashboard").Cells(42 + xrow, YearCounter + 1) = Sheets(ClaimsArray(cl)).Cells(15, ycolumn + 4 - backcolumn)
                backcolumn = backcolumn + 1
                xrow = xrow + 1
            Loop
            
        End If
    
        YearCounter = YearCounter - 1
        ycolumn = (YearCounter) * 8 - 7
    Loop

    If HaveVision = False And cl = 1 Then
        cl = 2
    End If

Next cl

YearCounter = CompleteYears

Do While YearCounter > 0
    If StartMonth = 1 Then
        Sheets("Dashboard").Cells(35, YearCounter + 1) = FinalYear - NYears + YearCounter
    Else
        Sheets("Dashboard").Cells(35, YearCounter + 1) = FinalYear - NYears + YearCounter + 1
    End If
    YearCounter = YearCounter - 1
Loop

    'Rolling 12-months
    
cl = 0

For cl = 0 To 2

    If HaveMedical = False Then
        cl = 2
    End If
    
    If cl <> 1 Then
        backcolumn = 1
    Else
        backcolumn = 3
    End If
    
    Do While backcolumn < 6
    
        rollingsum = 0
        rollingcounter = 0
        uprow = 0
        LLoop = 0
        
        YearCounter = NYears
        
        ycolumn = YearCounter * 8 - 7
        
        Do While FinalMonth - rollingcounter > 0
            rollingsum = rollingsum + Sheets(ClaimsArray(cl)).Cells(2 + FinalMonth - uprow, ycolumn + 6 - backcolumn).Value
            rollingcounter = rollingcounter + 1
            uprow = uprow + 1
        Loop
        
        uprow = 0
        YearCounter = YearCounter - 1
        ycolumn = YearCounter * 8 - 7
        
        Do While ycolumn > -2
        
            Do While uprow < 12
            
                Do While rollingcounter < 12 And uprow < 12
                    rollingsum = rollingsum + Sheets(ClaimsArray(cl)).Cells(14 - uprow, ycolumn + 6 - backcolumn)
                    uprow = uprow + 1
                    rollingcounter = rollingcounter + 1
                Loop
                
                If rollingcounter = 12 Then
                    If cl = 0 Then
                        Sheets("Dashboard").Cells(41 + backcolumn, CompleteYears + 3 + Rolltwelveyears - LLoop) = rollingsum
                    End If
                    If cl = 1 Then
                        Sheets("Dashboard").Cells(47 + backcolumn, CompleteYears + 3 + Rolltwelveyears - LLoop) = rollingsum
                    End If
                    If cl = 2 Then
                        Sheets("Dashboard").Cells(55 + backcolumn, CompleteYears + 3 + Rolltwelveyears - LLoop) = rollingsum
                    End If
                    rollingcounter = 0
                    rollingsum = 0
                    LLoop = LLoop + 1
                End If
            
            Loop
            
            YearCounter = YearCounter - 1
            ycolumn = YearCounter * 8 - 7
            uprow = 0
            
        Loop
        
        backcolumn = backcolumn + 1
    
    Loop
    
    If HaveVision = False And cl = 1 Then
        cl = 2
    End If

Next cl

If HaveMedical = True Then
    
    Sheets("Dashboard").Cells(40, CompleteYears + 3) = "Total Paid in Medical Claims by Type - Rolling 12"
    Sheets("Dashboard").Cells(48, CompleteYears + 3) = "Total Paid in Rx Claims by Type - Rolling 12"
    Sheets("Dashboard").Cells(42, CompleteYears + 3) = "Out-Patient"
    Sheets("Dashboard").Cells(43, CompleteYears + 3) = "Other Equipment"
    Sheets("Dashboard").Cells(44, CompleteYears + 3) = "Medical/Surgical"
    Sheets("Dashboard").Cells(45, CompleteYears + 3) = "Lab/X-Ray"
    Sheets("Dashboard").Cells(46, CompleteYears + 3) = "In-Patient"
    Sheets("Dashboard").Cells(50, CompleteYears + 3) = "Specialty"
    Sheets("Dashboard").Cells(51, CompleteYears + 3) = "Generic"
    Sheets("Dashboard").Cells(52, CompleteYears + 3) = "Brand"
    
End If

If HaveVision = True Then
    Sheets("Dashboard").Cells(54, CompleteYears + 3) = "Total Paid in Vision Claims by Type - Rolling 12"
    Sheets("Dashboard").Cells(56, CompleteYears + 3) = "Miscellaneous"
    Sheets("Dashboard").Cells(57, CompleteYears + 3) = "Lenses"
    Sheets("Dashboard").Cells(58, CompleteYears + 3) = "Frames"
    Sheets("Dashboard").Cells(59, CompleteYears + 3) = "Exam"
    Sheets("Dashboard").Cells(60, CompleteYears + 3) = "Contact Lenses"
    
End If

LLoop = 0

Do While LLoop < 3

    If HaveMedical = False Then
        LLoop = 2
    End If
    If HaveVision = False And LLoop = 2 Then
        GoTo NoRollTwelveByType
    End If
    If LLoop = 0 Then
        uprow = 0
    End If
    If LLoop = 1 Then
        uprow = -8
    End If
    If LLoop = 2 Then
        uprow = -14
    End If
    
    SLoop = 0
    
    FinalMonthTemp = FinalMonth
    
    If FinalMonth = 12 Then
        FinalMonthTemp = 1
    End If
    
    If FinalMonth <> 12 Then
        Do While SLoop < Rolltwelveyears
            Sheets("Dashboard").Cells(41 - uprow, CompleteYears + 3 + Rolltwelveyears - SLoop) = (FinalMonthTemp + 1) & "/" & (FinalYear - 1 - SLoop) & "-" & FinalMonth & "/" & (FinalYear - SLoop)
            SLoop = SLoop + 1
        Loop
    Else
        Do While SLoop < Rolltwelveyears
            Sheets("Dashboard").Cells(41 - uprow, CompleteYears + 3 + Rolltwelveyears - SLoop) = 1 & "/" & (FinalYear - SLoop) & "-" & (FinalMonth) & "/" & (FinalYear - SLoop)
            SLoop = SLoop + 1
        Loop
    End If
    
NoRollTwelveByType:
    
    LLoop = LLoop + 1

Loop

    'Year-to-Date
    
If HaveMedical = True Then
    
    Sheets("Dashboard").Cells(40, CompleteYears + Rolltwelveyears + 5) = "Total Paid in Medical Claims by Type - Year-to-Date"
    Sheets("Dashboard").Cells(48, CompleteYears + Rolltwelveyears + 5) = "Total Paid in Rx Claims by Type - Year-to-Date"
    Sheets("Dashboard").Cells(42, CompleteYears + Rolltwelveyears + 5) = "Out-Patient"
    Sheets("Dashboard").Cells(43, CompleteYears + Rolltwelveyears + 5) = "Other Equipment"
    Sheets("Dashboard").Cells(44, CompleteYears + Rolltwelveyears + 5) = "Medical/Surgical"
    Sheets("Dashboard").Cells(45, CompleteYears + Rolltwelveyears + 5) = "Lab/X-Ray"
    Sheets("Dashboard").Cells(46, CompleteYears + Rolltwelveyears + 5) = "In-Patient"
    Sheets("Dashboard").Cells(50, CompleteYears + Rolltwelveyears + 5) = "Specialty"
    Sheets("Dashboard").Cells(51, CompleteYears + Rolltwelveyears + 5) = "Generic"
    Sheets("Dashboard").Cells(52, CompleteYears + Rolltwelveyears + 5) = "Brand"
    
    If CompleteYears <> 0 Then
        Sheets("Dashboard").Cells(41, CompleteYears + Rolltwelveyears + 6) = "'" & 1 & "-" & FinalMonth & "/" & FinalYear
        Sheets("Dashboard").Cells(49, CompleteYears + Rolltwelveyears + 6) = "'" & 1 & "-" & FinalMonth & "/" & FinalYear
    Else
        Sheets("Dashboard").Cells(41, CompleteYears + Rolltwelveyears + 6) = "'" & StartMonth & "-" & FinalMonth & "/" & FinalYear
        Sheets("Dashboard").Cells(49, CompleteYears + Rolltwelveyears + 6) = "'" & StartMonth & "-" & FinalMonth & "/" & FinalYear
    End If
    
End If

If HaveVision = True Then

    Sheets("Dashboard").Cells(54, CompleteYears + Rolltwelveyears + 5) = "Total Paid in Vision Claims by Type - Year-to-Date"
    Sheets("Dashboard").Cells(56, CompleteYears + Rolltwelveyears + 5) = "Miscellaneous"
    Sheets("Dashboard").Cells(57, CompleteYears + Rolltwelveyears + 5) = "Lenses"
    Sheets("Dashboard").Cells(58, CompleteYears + Rolltwelveyears + 5) = "Frames"
    Sheets("Dashboard").Cells(59, CompleteYears + Rolltwelveyears + 5) = "Exam"
    Sheets("Dashboard").Cells(60, CompleteYears + Rolltwelveyears + 5) = "Contact Lenses"
    
    If CompleteYears <> 0 Then
        Sheets("Dashboard").Cells(55, CompleteYears + Rolltwelveyears + 6) = "'" & 1 & "-" & FinalMonth & "/" & FinalYear
    Else
        Sheets("Dashboard").Cells(55, CompleteYears + Rolltwelveyears + 6) = "'" & StartMonth & "-" & FinalMonth & "/" & FinalYear
    End If
    
End If

cl = 0

ycolumn = NYears * 8 - 7

For cl = 0 To 2

    If HaveMedical = False Then
        cl = 2
    End If
    
    If HaveVision = False And cl = 2 Then
       GoTo NoHaveVisionYTDbyType
    End If
    
    If cl <> 1 Then
        backcolumn = 1
    Else
        backcolumn = 3
    End If
    
    Do While backcolumn < 6
    
        If cl = 0 Then
            Sheets("Dashboard").Cells(41 + backcolumn, CompleteYears + Rolltwelveyears + 6) = Sheets(ClaimsArray(cl)).Cells(15, ycolumn + 6 - backcolumn)
        End If
        If cl = 1 Then
            Sheets("Dashboard").Cells(47 + backcolumn, CompleteYears + Rolltwelveyears + 6) = Sheets(ClaimsArray(cl)).Cells(15, ycolumn + 6 - backcolumn)
        End If
        If cl = 2 Then
            Sheets("Dashboard").Cells(55 + backcolumn, CompleteYears + Rolltwelveyears + 6) = Sheets(ClaimsArray(cl)).Cells(15, ycolumn + 6 - backcolumn)
        End If
        
        backcolumn = backcolumn + 1
        
    Loop
    
NoHaveVisionYTDbyType:
     
Next cl

'Graph Set 6 - consolidated bar graphs for spending by claim type PEPY for Complete Years, Rolling 12, and YTD

e = 0

If HaveMedical = True Then
    
    Sheets("Dashboard").Cells(62, 1) = "Total Paid in Medical Claims PEPY by Type - Complete Years Only"
    Sheets("Dashboard").Cells(70, 1) = "Total Paid in Rx Claims PEPY by Type - Complete Years Only"
    Sheets("Dashboard").Cells(64, 1) = "Out-Patient"
    Sheets("Dashboard").Cells(65, 1) = "Other Equipment"
    Sheets("Dashboard").Cells(66, 1) = "Medical/Surgical"
    Sheets("Dashboard").Cells(67, 1) = "Lab/X-Ray"
    Sheets("Dashboard").Cells(68, 1) = "In-Patient"
    Sheets("Dashboard").Cells(72, 1) = "Specialty"
    Sheets("Dashboard").Cells(73, 1) = "Generic"
    Sheets("Dashboard").Cells(74, 1) = "Brand"
    
    YearCounter = CompleteYears

    Do While YearCounter > 0
        If StartMonth = 1 Then
            Sheets("Dashboard").Cells(63, YearCounter + 1) = FinalYear - NYears + YearCounter
            Sheets("Dashboard").Cells(71, YearCounter + 1) = FinalYear - NYears + YearCounter
        Else
            Sheets("Dashboard").Cells(63, YearCounter + 1) = FinalYear - NYears + YearCounter + 1
            Sheets("Dashboard").Cells(71, YearCounter + 1) = FinalYear - NYears + YearCounter + 1
        End If
        YearCounter = YearCounter - 1
    Loop
    
End If

If HaveVision = True Then
    Sheets("Dashboard").Cells(76, 1) = "Total Paid in Vision Claims PEPY by Type - Complete Years Only"
    Sheets("Dashboard").Cells(78, 1) = "Miscellaneous"
    Sheets("Dashboard").Cells(79, 1) = "Lenses"
    Sheets("Dashboard").Cells(80, 1) = "Frames"
    Sheets("Dashboard").Cells(81, 1) = "Exam"
    Sheets("Dashboard").Cells(82, 1) = "Contact Lenses"
    
    YearCounter = CompleteYears

    Do While YearCounter > 0
        If StartMonth = 1 Then
            Sheets("Dashboard").Cells(77, YearCounter + 1) = FinalYear - NYears + YearCounter
        Else
            Sheets("Dashboard").Cells(77, YearCounter + 1) = FinalYear - NYears + YearCounter + 1
        End If
        YearCounter = YearCounter - 1
    Loop
    
End If

For e = 0 To 2

    If HaveMedical = False Then
        e = 2
    End If

    YearCounter = CompleteYears
    
    ycolumn = YearCounter * 8 - 7
    
    Do While YearCounter > 0
        If e <> 1 Then
        
            backcolumn = 1
            
            If e = 0 Then
                xrow = 0
            End If

            If e = 2 Then
                xrow = 14
            End If
            
            Do While backcolumn < 6
                Sheets("Dashboard").Cells(64 + xrow, YearCounter + 1) = Sheets(EmpArray(e)).Cells(15, ycolumn + 6 - backcolumn)
                backcolumn = backcolumn + 1
                xrow = xrow + 1
            Loop
            
        Else
        
            backcolumn = 1
            xrow = 8
            
            Do While backcolumn < 4
                Sheets("Dashboard").Cells(64 + xrow, YearCounter + 1) = Sheets(EmpArray(e)).Cells(15, ycolumn + 4 - backcolumn)
                backcolumn = backcolumn + 1
                xrow = xrow + 1
            Loop
            
        End If
    
        YearCounter = YearCounter - 1
        ycolumn = (YearCounter) * 8 - 7
    Loop

    If HaveVision = False And e = 1 Then
        e = 2
    End If

Next e

YearCounter = CompleteYears

Do While YearCounter > 0
    If StartMonth = 1 Then
        Sheets("Dashboard").Cells(63, YearCounter + 1) = FinalYear - NYears + YearCounter
    Else
        Sheets("Dashboard").Cells(63, YearCounter + 1) = FinalYear - NYears + YearCounter + 1
    End If
    YearCounter = YearCounter - 1
Loop

    'Rolling 12-months
    
e = 0

For e = 0 To 2

    If HaveMedical = False Then
        e = 2
    End If
    
    If e <> 1 Then
        backcolumn = 1
    Else
        backcolumn = 3
    End If
    
    Do While backcolumn < 6
    
        rollingsum = 0
        rollingcounter = 0
        uprow = 0
        LLoop = 0
        
        YearCounter = NYears
        
        ycolumn = YearCounter * 8 - 7
        
        Do While FinalMonth - rollingcounter > 0
            rollingsum = rollingsum + Sheets(EmpArray(e)).Cells(2 + FinalMonth - uprow, ycolumn + 6 - backcolumn).Value
            rollingcounter = rollingcounter + 1
            uprow = uprow + 1
        Loop
        
        uprow = 0
        YearCounter = YearCounter - 1
        ycolumn = YearCounter * 8 - 7
        
        Do While ycolumn > -2
        
            Do While uprow < 12
            
                Do While rollingcounter < 12 And uprow < 12
                    rollingsum = rollingsum + Sheets(EmpArray(e)).Cells(14 - uprow, ycolumn + 6 - backcolumn)
                    uprow = uprow + 1
                    rollingcounter = rollingcounter + 1
                Loop
                
                If rollingcounter = 12 Then
                    If e = 0 Then
                        Sheets("Dashboard").Cells(63 + backcolumn, CompleteYears + 3 + Rolltwelveyears - LLoop) = rollingsum
                    End If
                    If e = 1 Then
                        Sheets("Dashboard").Cells(69 + backcolumn, CompleteYears + 3 + Rolltwelveyears - LLoop) = rollingsum
                    End If
                    If e = 2 Then
                        Sheets("Dashboard").Cells(77 + backcolumn, CompleteYears + 3 + Rolltwelveyears - LLoop) = rollingsum
                    End If
                    rollingcounter = 0
                    rollingsum = 0
                    LLoop = LLoop + 1
                End If
            
            Loop
            
            YearCounter = YearCounter - 1
            ycolumn = YearCounter * 8 - 7
            uprow = 0
            
        Loop
        
        backcolumn = backcolumn + 1
    
    Loop
    
    If HaveVision = False And e = 1 Then
        e = 2
    End If

Next e

If HaveMedical = True Then
    
    Sheets("Dashboard").Cells(62, CompleteYears + 3) = "Total Paid in Medical Claims PEPY by Type - Rolling 12"
    Sheets("Dashboard").Cells(70, CompleteYears + 3) = "Total Paid in Rx Claims PEPY by Type - Rolling 12"
    Sheets("Dashboard").Cells(64, CompleteYears + 3) = "Out-Patient"
    Sheets("Dashboard").Cells(65, CompleteYears + 3) = "Other Equipment"
    Sheets("Dashboard").Cells(66, CompleteYears + 3) = "Medical/Surgical"
    Sheets("Dashboard").Cells(67, CompleteYears + 3) = "Lab/X-Ray"
    Sheets("Dashboard").Cells(68, CompleteYears + 3) = "In-Patient"
    Sheets("Dashboard").Cells(72, CompleteYears + 3) = "Specialty"
    Sheets("Dashboard").Cells(73, CompleteYears + 3) = "Generic"
    Sheets("Dashboard").Cells(74, CompleteYears + 3) = "Brand"
    
End If

If HaveVision = True Then
    Sheets("Dashboard").Cells(76, CompleteYears + 3) = "Total Paid in Vision Claims PEPY by Type - Rolling 12"
    Sheets("Dashboard").Cells(78, CompleteYears + 3) = "Miscellaneous"
    Sheets("Dashboard").Cells(79, CompleteYears + 3) = "Lenses"
    Sheets("Dashboard").Cells(80, CompleteYears + 3) = "Frames"
    Sheets("Dashboard").Cells(81, CompleteYears + 3) = "Exam"
    Sheets("Dashboard").Cells(82, CompleteYears + 3) = "Contact Lenses"
    
End If

LLoop = 0

Do While LLoop < 3

    If HaveMedical = False Then
        LLoop = 2
    End If
    If HaveVision = False And LLoop = 2 Then
        GoTo NoRollTwelveByTypePEPY
    End If
    If LLoop = 0 Then
        uprow = 0
    End If
    If LLoop = 1 Then
        uprow = -8
    End If
    If LLoop = 2 Then
        uprow = -14
    End If
    
    SLoop = 0
    
    FinalMonthTemp = FinalMonth
    
    If FinalMonth = 12 Then
        FinalMonthTemp = 1
    End If
    
    If FinalMonth <> 12 Then
        Do While SLoop < Rolltwelveyears
            Sheets("Dashboard").Cells(63 - uprow, CompleteYears + 3 + Rolltwelveyears - SLoop) = (FinalMonthTemp + 1) & "/" & (FinalYear - 1 - SLoop) & "-" & FinalMonth & "/" & (FinalYear - SLoop)
            SLoop = SLoop + 1
        Loop
    Else
        Do While SLoop < Rolltwelveyears
            Sheets("Dashboard").Cells(63 - uprow, CompleteYears + 3 + Rolltwelveyears - SLoop) = 1 & "/" & (FinalYear - SLoop) & "-" & (FinalMonth) & "/" & (FinalYear - SLoop)
            SLoop = SLoop + 1
        Loop
    End If
    
NoRollTwelveByTypePEPY:
    
    LLoop = LLoop + 1

Loop

    'Year-to-Date
    
If HaveMedical = True Then
    
    Sheets("Dashboard").Cells(62, CompleteYears + Rolltwelveyears + 5) = "Total Paid in Medical Claims PEPY by Type - Year-to-Date"
    Sheets("Dashboard").Cells(70, CompleteYears + Rolltwelveyears + 5) = "Total Paid in Rx Claims PEPY by Type - Year-to-Date"
    Sheets("Dashboard").Cells(64, CompleteYears + Rolltwelveyears + 5) = "Out-Patient"
    Sheets("Dashboard").Cells(65, CompleteYears + Rolltwelveyears + 5) = "Other Equipment"
    Sheets("Dashboard").Cells(66, CompleteYears + Rolltwelveyears + 5) = "Medical/Surgical"
    Sheets("Dashboard").Cells(67, CompleteYears + Rolltwelveyears + 5) = "Lab/X-Ray"
    Sheets("Dashboard").Cells(68, CompleteYears + Rolltwelveyears + 5) = "In-Patient"
    Sheets("Dashboard").Cells(72, CompleteYears + Rolltwelveyears + 5) = "Specialty"
    Sheets("Dashboard").Cells(73, CompleteYears + Rolltwelveyears + 5) = "Generic"
    Sheets("Dashboard").Cells(74, CompleteYears + Rolltwelveyears + 5) = "Brand"
    
    If CompleteYears <> 0 Then
        Sheets("Dashboard").Cells(63, CompleteYears + Rolltwelveyears + 6) = "'" & 1 & "-" & FinalMonth & "/" & FinalYear
        Sheets("Dashboard").Cells(71, CompleteYears + Rolltwelveyears + 6) = "'" & 1 & "-" & FinalMonth & "/" & FinalYear
    Else
        Sheets("Dashboard").Cells(63, CompleteYears + Rolltwelveyears + 6) = "'" & StartMonth & "-" & FinalMonth & "/" & FinalYear
        Sheets("Dashboard").Cells(61, CompleteYears + Rolltwelveyears + 6) = "'" & StartMonth & "-" & FinalMonth & "/" & FinalYear
    End If
    
End If

If HaveVision = True Then

    Sheets("Dashboard").Cells(76, CompleteYears + Rolltwelveyears + 5) = "Total Paid in Vision Claims PEPY by Type - Year-to-Date"
    Sheets("Dashboard").Cells(78, CompleteYears + Rolltwelveyears + 5) = "Miscellaneous"
    Sheets("Dashboard").Cells(79, CompleteYears + Rolltwelveyears + 5) = "Lenses"
    Sheets("Dashboard").Cells(80, CompleteYears + Rolltwelveyears + 5) = "Frames"
    Sheets("Dashboard").Cells(81, CompleteYears + Rolltwelveyears + 5) = "Exam"
    Sheets("Dashboard").Cells(82, CompleteYears + Rolltwelveyears + 5) = "Contact Lenses"
    
    If CompleteYears <> 0 Then
        Sheets("Dashboard").Cells(77, CompleteYears + Rolltwelveyears + 6) = "'" & 1 & "-" & FinalMonth & "/" & FinalYear
    Else
        Sheets("Dashboard").Cells(77, CompleteYears + Rolltwelveyears + 6) = "'" & StartMonth & "-" & FinalMonth & "/" & FinalYear
    End If
    
End If

e = 0

ycolumn = NYears * 8 - 7

For e = 0 To 2

    If HaveMedical = False Then
        e = 2
    End If
    
    If HaveVision = False And e = 2 Then
       GoTo NoHaveVisionYTDbyTypePEPY
    End If
    
    If e <> 1 Then
        backcolumn = 1
    Else
        backcolumn = 3
    End If
    
    Do While backcolumn < 6
    
        If e = 0 Then
            Sheets("Dashboard").Cells(63 + backcolumn, CompleteYears + Rolltwelveyears + 6) = Sheets(EmpArray(e)).Cells(15, ycolumn + 6 - backcolumn)
        End If
        If e = 1 Then
            Sheets("Dashboard").Cells(69 + backcolumn, CompleteYears + Rolltwelveyears + 6) = Sheets(EmpArray(e)).Cells(15, ycolumn + 6 - backcolumn)
        End If
        If e = 2 Then
            Sheets("Dashboard").Cells(77 + backcolumn, CompleteYears + Rolltwelveyears + 6) = Sheets(EmpArray(e)).Cells(15, ycolumn + 6 - backcolumn)
        End If
        
        backcolumn = backcolumn + 1
        
    Loop
    
NoHaveVisionYTDbyTypePEPY:
     
Next e

'Graph Set 7 - Total Cost Paid in Claims PMPM by Type Total and Past 12 Months

    'Past 12 Months - Medical, Rx and Vision
    
xrow = 0
backcolumn = 1

If CompleteYears <> 0 Then
    
    If HaveMedical = True Then
        Sheets("Dashboard").Cells(84, 1) = "Total Paid in Medical Claims PMPM by Type - Past 12 Months"
        Sheets("Dashboard").Cells(92, 1) = "Total Paid in Rx Claims PMPM by Type - Past 12 Months"
    End If
    
    If HaveVision = True Then
        Sheets("Dashboard").Cells(98, 1) = "Total Cost Paid in Vision Claims PMPM by Type - Past 12 Months"
    End If
    
    l = 0
    
    For l = 0 To 2
    
        If (HaveMedical = True And l = 0) Or (HaveVision = True And l = 2) Then
        
            If HaveMedical = False Then
                l = 2
                xrow = xrow + 13
            End If
                
            Do While backcolumn < 6
            
                rollingcounter = 0
                uprow = 0
                FinalMonthTemp = FinalMonth
                YearCounter = NYears
                ycolumn = YearCounter * 8 - 7
                
                    Do While rollingcounter < 12
                    
                        Do While FinalMonthTemp - uprow > 0 And rollingcounter < 12
            
                            Sheets("Dashboard").Cells(86 + l + xrow, 13 - rollingcounter) = Sheets(LifeArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6 - backcolumn).Value
                                
                            rollingcounter = rollingcounter + 1
                            uprow = uprow + 1
                
                        Loop
                        
                        FinalMonthTemp = 12
                        uprow = 0
                        YearCounter = YearCounter - 1
                        ycolumn = YearCounter * 8 - 7
                        
                    Loop
            
                backcolumn = backcolumn + 1
                xrow = xrow + 1
            
            Loop
        
            xrow = xrow + 8
            backcolumn = 1
        
        End If
        
        If HaveMedical = True And l = 1 Then
            
            xrow = 0
            backcolumn = backcolumn + 2
        
            Do While backcolumn < 6
                
                    rollingcounter = 0
                    uprow = 0
                    FinalMonthTemp = FinalMonth
                    YearCounter = NYears
                    ycolumn = YearCounter * 8 - 7
                    
                        Do While rollingcounter < 12
                        
                            Do While FinalMonthTemp - uprow > 0 And rollingcounter < 12
                
                                Sheets("Dashboard").Cells(94 + xrow, 13 - rollingcounter) = Sheets(LifeArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6 - backcolumn).Value
                                    
                                rollingcounter = rollingcounter + 1
                                uprow = uprow + 1
                    
                            Loop
                            
                            FinalMonthTemp = 12
                            uprow = 0
                            YearCounter = YearCounter - 1
                            ycolumn = YearCounter * 8 - 7
                            
                        Loop
                
                    backcolumn = backcolumn + 1
                    xrow = xrow + 1
                
                Loop
            
            xrow = 12
            backcolumn = 1
            
        End If
        
    Next l
    
    rollingcounter = 0
    uprow = 0
    FinalMonthTemp = FinalMonth
    YearCounter = NYears
    ycolumn = YearCounter * 8 - 7
    xrow = 0
    SLoop = 0
    
    If HaveMedical = True Then
        l = 0
    Else
        l = 2
    End If
    
    Do While SLoop < 2
    
        If HaveMedical = False Then
            SLoop = 1
            xrow = xrow + 14
        End If
        
        If HaveVision = False Then
            SLoop = 1
        End If
    
        Do While rollingcounter < 12
            
            Do While FinalMonthTemp - uprow > 0 And rollingcounter < 12
            
                Sheets("Dashboard").Cells(85 + xrow, 13 - rollingcounter) = "'" & Sheets(LifeArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn).Value & "/" & Sheets(LifeArray(l)).Cells(1, ycolumn).Value
                    
                rollingcounter = rollingcounter + 1
                uprow = uprow + 1
        
            Loop
            
            FinalMonthTemp = 12
            uprow = 0
            YearCounter = YearCounter - 1
            ycolumn = YearCounter * 8 - 7
            
        Loop
        
        If HaveMedical = True And SLoop = 0 Then
            Range("B85:M85").Select
            Selection.Copy
            Range("B93:M93").PasteSpecial
        End If
        
        xrow = xrow + 14
        SLoop = SLoop + 1
        
        rollingcounter = 0
        uprow = 0
        FinalMonthTemp = FinalMonth
        YearCounter = NYears
        ycolumn = YearCounter * 8 - 7
    
    Loop
    
    If HaveMedical = True Then
        Sheets("Dashboard").Cells(86, 1) = "Out-Patient"
        Sheets("Dashboard").Cells(87, 1) = "Other Equipment"
        Sheets("Dashboard").Cells(88, 1) = "Medical/Surgical"
        Sheets("Dashboard").Cells(89, 1) = "Lab/X-Ray"
        Sheets("Dashboard").Cells(90, 1) = "In-Patient"
        
        Sheets("Dashboard").Cells(94, 1) = "Specialty"
        Sheets("Dashboard").Cells(95, 1) = "Generic"
        Sheets("Dashboard").Cells(96, 1) = "Brand"
        
    End If
    
    If HaveVision = True Then
        Sheets("Dashboard").Cells(100, 1) = "Miscellaneous"
        Sheets("Dashboard").Cells(101, 1) = "Lenses"
        Sheets("Dashboard").Cells(102, 1) = "Frames"
        Sheets("Dashboard").Cells(103, 1) = "Exam"
        Sheets("Dashboard").Cells(104, 1) = "Contact Lenses"
    End If
    
End If

    'Total - Medical and Vision

If HaveMedical = True And HaveVision = True Then
    Sheets("Dashboard").Cells(84, 15) = "Total Paid in Medical Claims PMPM by Type - Total"
    Sheets("Dashboard").Cells(92, 15) = "Total Paid in Rx Claims PMPM by Type - Total"
    Sheets("Dashboard").Cells(98, 15) = "Total Paid in Vision Claims PMPM by Type - Total"
End If

If HaveMedical = True And HaveVision = False Then
    Sheets("Dashboard").Cells(84, 15) = "Unit Cost Paid in Medical Claims PMPM by Type - Total"
    Sheets("Dashboard").Cells(92, 15) = "Total Paid in Rx Claims PMPM by Type - Total"
End If

If HaveMedical = False And HaveVision = True Then
    Sheets("Dashboard").Cells(98, 15) = "Unit Cost Paid in Vision Claims PMPM by Type - Total"
End If

xrow = 0
backcolumn = 1
l = 0

For l = 0 To 2

    If (HaveMedical = True And l = 0) Or (HaveVision = True And l = 1) Then

        If HaveMedical = False Then
            l = 2
            xrow = xrow + 13
        End If
        
        Do While backcolumn < 6
    
            rollingcounter = 0
            uprow = 0
            FinalMonthTemp = FinalMonth
            YearCounter = NYears
            ycolumn = YearCounter * 8 - 7
            
            Do While rollingcounter < TotalMonths
            
                Do While FinalMonthTemp - uprow > 0 And rollingcounter < TotalMonths
                
                    Sheets("Dashboard").Cells(86 + l + xrow, 15 + TotalMonths - rollingcounter) = Sheets(LifeArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6 - backcolumn).Value
                    
                    rollingcounter = rollingcounter + 1
                    uprow = uprow + 1
        
                Loop
                
                FinalMonthTemp = 12
                uprow = 0
                YearCounter = YearCounter - 1
                ycolumn = YearCounter * 8 - 7
                
            Loop
            
            backcolumn = backcolumn + 1
            xrow = xrow + 1
        
        Loop
        
        backcolumn = 1
        xrow = xrow + 8
    
    End If
    
    If HaveMedical = True And l = 1 Then
    
        xrow = 0
        backcolumn = backcolumn + 2
    
        Do While backcolumn < 6
    
            rollingcounter = 0
            uprow = 0
            FinalMonthTemp = FinalMonth
            YearCounter = NYears
            ycolumn = YearCounter * 8 - 7
            
            Do While rollingcounter < TotalMonths
            
                Do While FinalMonthTemp - uprow > 0 And rollingcounter < TotalMonths
                
                    Sheets("Dashboard").Cells(94 + xrow, 15 + TotalMonths - rollingcounter) = Sheets(LifeArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6 - backcolumn).Value
                    
                    rollingcounter = rollingcounter + 1
                    uprow = uprow + 1
        
                Loop
                
                FinalMonthTemp = 12
                uprow = 0
                YearCounter = YearCounter - 1
                ycolumn = YearCounter * 8 - 7
                
            Loop
            
            backcolumn = backcolumn + 1
            xrow = xrow + 1
        
        Loop
        
        backcolumn = 1
        xrow = 12
    
    End If
    
Next l

rollingcounter = 0
uprow = 0
FinalMonthTemp = FinalMonth
YearCounter = NYears
ycolumn = YearCounter * 8 - 7
xrow = 0
SLoop = 0

If HaveMedical = True Then
    l = 0
Else
    l = 2
End If

Do While SLoop < 2
    
        If HaveMedical = False Then
            SLoop = 1
            xrow = xrow + 14
        End If
        
        If HaveVision = False Then
            SLoop = 1
        End If

    Do While rollingcounter < TotalMonths
        
        Do While FinalMonthTemp - uprow > 0 And rollingcounter < TotalMonths
    
            Sheets("Dashboard").Cells(85 + xrow, 15 + TotalMonths - rollingcounter) = "'" & Sheets(LifeArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn).Value & "/" & Sheets(LifeArray(l)).Cells(1, ycolumn).Value
                    
            rollingcounter = rollingcounter + 1
            uprow = uprow + 1
    
        Loop
        
        FinalMonthTemp = 12
        uprow = 0
        YearCounter = YearCounter - 1
        ycolumn = YearCounter * 8 - 7
            
    Loop
    
    If HaveMedical = True And SLoop = 0 Then
        Range(Cells(85, 15), Cells(85, 15 + TotalMonths)).Select
        Selection.Copy
        Range(Cells(93, 15), Cells(93, 15 + TotalMonths)).PasteSpecial
    End If
    
    xrow = xrow + 14
    SLoop = SLoop + 1
    
    rollingcounter = 0
    uprow = 0
    FinalMonthTemp = FinalMonth
    YearCounter = NYears
    ycolumn = YearCounter * 8 - 7
    
Loop

If HaveMedical = True Then
    Sheets("Dashboard").Cells(86, 15) = "Out-Patient"
    Sheets("Dashboard").Cells(87, 15) = "Other Equipment"
    Sheets("Dashboard").Cells(88, 15) = "Medical/Surgical"
    Sheets("Dashboard").Cells(89, 15) = "Lab/X-Ray"
    Sheets("Dashboard").Cells(90, 15) = "In-Patient"
    
    Sheets("Dashboard").Cells(94, 15) = "Specialty"
    Sheets("Dashboard").Cells(95, 15) = "Generic"
    Sheets("Dashboard").Cells(96, 15) = "Brand"
End If

If HaveVision = True Then
    Sheets("Dashboard").Cells(100, 15) = "Miscellaneous"
    Sheets("Dashboard").Cells(101, 15) = "Lenses"
    Sheets("Dashboard").Cells(102, 15) = "Frames"
    Sheets("Dashboard").Cells(103, 15) = "Exam"
    Sheets("Dashboard").Cells(104, 15) = "Contact Lenses"
End If

'Graph Set 8 added 1-7-2019 Total Cost Paid in Claims PEPM

Sheets("Dashboard").Cells(106, 1) = "Total Paid in Claims PEPM - Total"

l = 0

For l = 0 To 2

    If HaveMedical = False Then
        l = 2
    End If

    rollingcounter = 0
    uprow = 0
    FinalMonthTemp = FinalMonth
    YearCounter = NYears
    ycolumn = YearCounter * 8 - 7
    If l = 1 Then
        ycolumn = ycolumn - 2
    End If
    
    Do While rollingcounter < TotalMonths
    
        Do While FinalMonthTemp - uprow > 0 And rollingcounter < TotalMonths
        
            If HaveMedical = True And HaveVision = True Then
                Sheets("Dashboard").Cells(108 + l, 1 + TotalMonths - rollingcounter) = Sheets(EmpArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6).Value
            End If
            If HaveMedical = False And HaveVision = True Then
                Sheets("Dashboard").Cells(108, 1 + TotalMonths - rollingcounter) = Sheets(EmpArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6).Value
            End If
            If HaveMedical = True And HaveVision = False Then
                If l = 2 Then
                
                Else
                    Sheets("Dashboard").Cells(108 + l, 1 + TotalMonths - rollingcounter) = Sheets(EmpArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6).Value
                End If
            End If
                
            rollingcounter = rollingcounter + 1
            uprow = uprow + 1

        Loop
        
        FinalMonthTemp = 12
        uprow = 0
        YearCounter = YearCounter - 1
        ycolumn = YearCounter * 8 - 7
        If l = 1 Then
            ycolumn = ycolumn - 2
        End If
        
    Loop
    
    If HaveVision = False And l = 1 Then
        l = 2
    End If

Next l

rollingcounter = 0
uprow = 0
FinalMonthTemp = FinalMonth
YearCounter = NYears
ycolumn = YearCounter * 8 - 7

If HaveMedical = True Then
    l = 0
Else
    l = 2
End If

Do While rollingcounter < TotalMonths
    
    Do While FinalMonthTemp - uprow > 0 And rollingcounter < TotalMonths

        Sheets("Dashboard").Cells(107, 1 + TotalMonths - rollingcounter) = "'" & Sheets(EmpArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn).Value & "/" & Sheets(LifeArray(l)).Cells(1, ycolumn).Value
                
        rollingcounter = rollingcounter + 1
        uprow = uprow + 1

    Loop
    
    FinalMonthTemp = 12
    uprow = 0
    YearCounter = YearCounter - 1
    ycolumn = YearCounter * 8 - 7
        
Loop


If HaveMedical = True And HaveVision = True Then
    Sheets("Dashboard").Cells(108, 1) = "Medical"
    Sheets("Dashboard").Cells(109, 1) = "Rx"
    Sheets("Dashboard").Cells(110, 1) = "Vision"
End If
If HaveMedical = True And HaveVision = False Then
    Sheets("Dashboard").Cells(108, 1) = "Medical"
    Sheets("Dashboard").Cells(109, 1) = "Rx"
End If
If HaveMedical = False And HaveVision = True Then
    Sheets("Dashboard").Cells(108, 1) = "Vision"
End If

'Graph Set 9 added 1-7-2019 Total Cost Paid in Claims by Type - Total

If HaveMedical = True And HaveVision = True Then
    Sheets("Dashboard").Cells(112, 1) = "Total Paid in Medical Claims by Type - Total"
    Sheets("Dashboard").Cells(121, 1) = "Total Paid in Rx Claims by Type - Total"
    Sheets("Dashboard").Cells(128, 1) = "Total Paid in Vision Claims by Type - Total"
End If

If HaveMedical = True And HaveVision = False Then
    Sheets("Dashboard").Cells(112, 1) = "Total Paid in Medical Claims by Type - Total"
    Sheets("Dashboard").Cells(121, 1) = "Total Paid in Rx Claims by Type - Total"
End If

If HaveMedical = False And HaveVision = True Then
    Sheets("Dashboard").Cells(128, 1) = "Total Paid in Vision Claims by Type - Total"
End If

xrow = 0
backcolumn = 1
l = 0

For l = 0 To 2

    If (HaveMedical = True And l = 0) Or (HaveVision = True And l = 1) Then

        If HaveMedical = False Then
            l = 2
            xrow = xrow + 14
        End If

        Do While backcolumn < 6

            rollingcounter = 0
            uprow = 0
            FinalMonthTemp = FinalMonth
            YearCounter = NYears
            ycolumn = YearCounter * 8 - 7

            Do While rollingcounter < TotalMonths

                Do While FinalMonthTemp - uprow > 0 And rollingcounter < TotalMonths

                    Sheets("Dashboard").Cells(114 + l + xrow, 1 + TotalMonths - rollingcounter) = Sheets(ClaimsArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6 - backcolumn).Value

                    rollingcounter = rollingcounter + 1
                    uprow = uprow + 1

                Loop

                FinalMonthTemp = 12
                uprow = 0
                YearCounter = YearCounter - 1
                ycolumn = YearCounter * 8 - 7

            Loop

            backcolumn = backcolumn + 1
            xrow = xrow + 1

        Loop

        backcolumn = 1
        xrow = xrow + 10

    End If

    If HaveMedical = True And l = 1 Then

        xrow = 0
        backcolumn = backcolumn + 2

        Do While backcolumn < 6

            rollingcounter = 0
            uprow = 0
            FinalMonthTemp = FinalMonth
            YearCounter = NYears
            ycolumn = YearCounter * 8 - 7

            Do While rollingcounter < TotalMonths

                Do While FinalMonthTemp - uprow > 0 And rollingcounter < TotalMonths

                    Sheets("Dashboard").Cells(123 + xrow, 1 + TotalMonths - rollingcounter) = Sheets(ClaimsArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn + 6 - backcolumn).Value

                    rollingcounter = rollingcounter + 1
                    uprow = uprow + 1

                Loop

                FinalMonthTemp = 12
                uprow = 0
                YearCounter = YearCounter - 1
                ycolumn = YearCounter * 8 - 7

            Loop

            backcolumn = backcolumn + 1
            xrow = xrow + 1

        Loop

        backcolumn = 1
        xrow = 13

    End If

Next l

rollingcounter = 0
uprow = 0
FinalMonthTemp = FinalMonth
YearCounter = NYears
ycolumn = YearCounter * 8 - 7
xrow = 0
SLoop = 0

If HaveMedical = True Then
    l = 0
Else
    l = 2
End If

Do While SLoop < 2

        If HaveMedical = False Then
            SLoop = 1
            xrow = xrow + 15
        End If

        If HaveVision = False Then
            SLoop = 1
        End If

    Do While rollingcounter < TotalMonths

        Do While FinalMonthTemp - uprow > 0 And rollingcounter < TotalMonths

            Sheets("Dashboard").Cells(113 + xrow, 1 + TotalMonths - rollingcounter) = "'" & Sheets(ClaimsArray(l)).Cells(2 + FinalMonthTemp - uprow, ycolumn).Value & "/" & Sheets(LifeArray(l)).Cells(1, ycolumn).Value

            rollingcounter = rollingcounter + 1
            uprow = uprow + 1

        Loop

        FinalMonthTemp = 12
        uprow = 0
        YearCounter = YearCounter - 1
        ycolumn = YearCounter * 8 - 7

    Loop

    If HaveMedical = True And SLoop = 0 Then
        Range(Cells(113, 1), Cells(113, 1 + TotalMonths)).Select
        Selection.Copy
        Range(Cells(122, 1), Cells(122, 1 + TotalMonths)).PasteSpecial
    End If

    xrow = xrow + 16
    SLoop = SLoop + 1

    rollingcounter = 0
    uprow = 0
    FinalMonthTemp = FinalMonth
    YearCounter = NYears
    ycolumn = YearCounter * 8 - 7

Loop

'Make Totals rows

ycolumn = 1

Do While ycolumn <= TotalMonths

    If HaveMedical = True Then
        Cells(119, ycolumn) = Application.WorksheetFunction.Sum(Range(Cells(114, ycolumn), Cells(118, ycolumn)))
        Cells(126, ycolumn) = Application.WorksheetFunction.Sum(Range(Cells(123, ycolumn), Cells(125, ycolumn)))
    End If
    
    If HaveVision = True Then
        Cells(135, ycolumn) = Application.WorksheetFunction.Sum(Range(Cells(130, ycolumn), Cells(134, ycolumn)))
    End If
    
    ycolumn = ycolumn + 1

Loop

If HaveMedical = True Then
    Sheets("Dashboard").Cells(114, 1) = "Out-Patient"
    Sheets("Dashboard").Cells(115, 1) = "Other Equipment"
    Sheets("Dashboard").Cells(116, 1) = "Medical/Surgical"
    Sheets("Dashboard").Cells(117, 1) = "Lab/X-Ray"
    Sheets("Dashboard").Cells(118, 1) = "In-Patient"
    Sheets("Dashboard").Cells(119, 1) = "Total"

    Sheets("Dashboard").Cells(123, 1) = "Specialty"
    Sheets("Dashboard").Cells(124, 1) = "Generic"
    Sheets("Dashboard").Cells(125, 1) = "Brand"
    Sheets("Dashboard").Cells(126, 1) = "Total"
End If

If HaveVision = True Then
    Sheets("Dashboard").Cells(130, 1) = "Miscellaneous"
    Sheets("Dashboard").Cells(131, 1) = "Lenses"
    Sheets("Dashboard").Cells(132, 1) = "Frames"
    Sheets("Dashboard").Cells(133, 1) = "Exam"
    Sheets("Dashboard").Cells(134, 1) = "Contact Lenses"
    Sheets("Dashboard").Cells(135, 1) = "Total"
End If

'Graph Set 10 - Unit Cost by Type Complete Years and Roll-12

e = 0

If HaveMedical = True Then
    
    Sheets("Dashboard").Cells(137, 1) = "Total Paid in Medical Claims Unit Cost by Type - Complete Years Only"
    Sheets("Dashboard").Cells(139, 1) = "Out-Patient"
    Sheets("Dashboard").Cells(140, 1) = "Other Equipment"
    Sheets("Dashboard").Cells(141, 1) = "Medical/Surgical"
    Sheets("Dashboard").Cells(142, 1) = "Lab/X-Ray"
    Sheets("Dashboard").Cells(143, 1) = "In-Patient"
    
    YearCounter = CompleteYears

    Do While YearCounter > 0
        If StartMonth = 1 Then
            Sheets("Dashboard").Cells(138, YearCounter + 1) = FinalYear - NYears + YearCounter
        Else
            Sheets("Dashboard").Cells(138, YearCounter + 1) = FinalYear - NYears + YearCounter + 1
        End If
        YearCounter = YearCounter - 1
    Loop
    
End If

If HaveVision = True Then
    Sheets("Dashboard").Cells(145, 1) = "Total Paid in Vision Claims Unit Cost by Type - Complete Years Only"
    Sheets("Dashboard").Cells(147, 1) = "Miscellaneous"
    Sheets("Dashboard").Cells(148, 1) = "Lenses"
    Sheets("Dashboard").Cells(149, 1) = "Frames"
    Sheets("Dashboard").Cells(150, 1) = "Exam"
    Sheets("Dashboard").Cells(151, 1) = "Contact Lenses"
    
    YearCounter = CompleteYears

    Do While YearCounter > 0
        If StartMonth = 1 Then
            Sheets("Dashboard").Cells(146, YearCounter + 1) = FinalYear - NYears + YearCounter
        Else
            Sheets("Dashboard").Cells(146, YearCounter + 1) = FinalYear - NYears + YearCounter + 1
        End If
        YearCounter = YearCounter - 1
    Loop
    
End If

For e = 0 To 1

    If HaveMedical = False Then
        e = 1
    End If

    YearCounter = CompleteYears
    
    ycolumn = YearCounter * 8 - 7
    
    Do While YearCounter > 0
        
            backcolumn = 1
            
            If e = 0 Then
                xrow = 0
            End If

            If e = 1 Then
                xrow = 8
            End If
            
            Do While backcolumn < 6
                Sheets("Dashboard").Cells(139 + xrow, YearCounter + 1) = Sheets(UnitArray(e)).Cells(15, ycolumn + 6 - backcolumn)
                backcolumn = backcolumn + 1
                xrow = xrow + 1
            Loop
    
        YearCounter = YearCounter - 1
        ycolumn = (YearCounter) * 8 - 7
    Loop

    If HaveVision = False Then
        e = 1
    End If

Next e

YearCounter = CompleteYears

Do While YearCounter > 0
    If StartMonth = 1 Then
        Sheets("Dashboard").Cells(138, YearCounter + 1) = FinalYear - NYears + YearCounter
    Else
        Sheets("Dashboard").Cells(138, YearCounter + 1) = FinalYear - NYears + YearCounter + 1
    End If
    YearCounter = YearCounter - 1
Loop

    'Rolling 12-months
    'Count Numbers
    
e = 0

For e = 0 To 1

    If HaveMedical = False Then
        e = 1
    End If

        backcolumn = 1

    Do While backcolumn < 6

        rollingsum = 0
        rollingcounter = 0
        uprow = 0
        LLoop = 0

        YearCounter = NYears

        ycolumn = YearCounter * 8 - 7

        Do While FinalMonth - rollingcounter > 0
            rollingsum = rollingsum + Sheets(CountArray(e)).Cells(2 + FinalMonth - uprow, ycolumn + 6 - backcolumn).Value
            rollingcounter = rollingcounter + 1
            uprow = uprow + 1
        Loop

        uprow = 0
        YearCounter = YearCounter - 1
        ycolumn = YearCounter * 8 - 7

        Do While ycolumn > -2

            Do While uprow < 12

                Do While rollingcounter < 12 And uprow < 12
                    rollingsum = rollingsum + Sheets(CountArray(e)).Cells(14 - uprow, ycolumn + 6 - backcolumn)
                    uprow = uprow + 1
                    rollingcounter = rollingcounter + 1
                Loop

                If rollingcounter = 12 Then
                    If e = 0 Then
                        Sheets("Dashboard").Cells(138 + backcolumn, CompleteYears + 3 + Rolltwelveyears - LLoop) = rollingsum
                    End If
                    If e = 1 Then
                        Sheets("Dashboard").Cells(146 + backcolumn, CompleteYears + 3 + Rolltwelveyears - LLoop) = rollingsum
                    End If
                    rollingcounter = 0
                    rollingsum = 0
                    LLoop = LLoop + 1
                End If

            Loop

            YearCounter = YearCounter - 1
            ycolumn = YearCounter * 8 - 7
            uprow = 0

        Loop

        backcolumn = backcolumn + 1

    Loop

    If HaveVision = False Then
        e = 1
    End If

Next e

    'Rolling 12-months
    'Costs Numbers

e = 0

For e = 0 To 2

    If HaveMedical = False Then
        e = 2
    End If

        backcolumn = 1

    Do While backcolumn < 6

        rollingsum = 0
        rollingcounter = 0
        uprow = 0
        LLoop = 0

        YearCounter = NYears

        ycolumn = YearCounter * 8 - 7

        Do While FinalMonth - rollingcounter > 0
            rollingsum = rollingsum + Sheets(ClaimsArray(e)).Cells(2 + FinalMonth - uprow, ycolumn + 6 - backcolumn).Value
            rollingcounter = rollingcounter + 1
            uprow = uprow + 1
        Loop

        uprow = 0
        YearCounter = YearCounter - 1
        ycolumn = YearCounter * 8 - 7

        Do While ycolumn > -2

            Do While uprow < 12

                Do While rollingcounter < 12 And uprow < 12
                    rollingsum = rollingsum + Sheets(ClaimsArray(e)).Cells(14 - uprow, ycolumn + 6 - backcolumn)
                    uprow = uprow + 1
                    rollingcounter = rollingcounter + 1
                Loop

                If rollingcounter = 12 Then
                    If e = 0 Then
                        Sheets("Dashboard").Cells(138 + backcolumn, CompleteYears + 5 + Rolltwelveyears * 2 - LLoop) = rollingsum
                    End If
                    If e = 2 Then
                        Sheets("Dashboard").Cells(146 + backcolumn, CompleteYears + 5 + Rolltwelveyears * 2 - LLoop) = rollingsum
                    End If
                    rollingcounter = 0
                    rollingsum = 0
                    LLoop = LLoop + 1
                End If

            Loop

            YearCounter = YearCounter - 1
            ycolumn = YearCounter * 8 - 7
            uprow = 0

        Loop

        backcolumn = backcolumn + 1

    Loop

    If HaveVision = False Then
        e = 2
    End If
    
    e = e + 1

Next e

Range("F2:G2").Select
Selection.Copy
Cells(138, CompleteYears + 4).PasteSpecial
Cells(146, CompleteYears + 4).PasteSpecial
Cells(138, CompleteYears + Rolltwelveyears + 6).PasteSpecial
Cells(146, CompleteYears + Rolltwelveyears + 6).PasteSpecial
Cells(138, CompleteYears + Rolltwelveyears * 2 + 8).PasteSpecial
Cells(146, CompleteYears + Rolltwelveyears * 2 + 8).PasteSpecial
Range(Cells(139, 1), Cells(143, 1)).Select
Selection.Copy
Cells(139, CompleteYears + 3).PasteSpecial
Cells(139, CompleteYears + Rolltwelveyears + 5).PasteSpecial
Cells(139, CompleteYears + Rolltwelveyears * 2 + 7).PasteSpecial
Range(Cells(147, 1), Cells(151, 1)).Select
Selection.Copy
Cells(147, CompleteYears + 3).PasteSpecial
Cells(147, CompleteYears + Rolltwelveyears + 5).PasteSpecial
Cells(147, CompleteYears + Rolltwelveyears * 2 + 7).PasteSpecial

'Calculating Rolling-12 Unit costs

xrow = 0
ycolumn = 0

Do While ycolumn < Rolltwelveyears

    xrow = 0
    
    Do While xrow <= 4
        
        If Sheets("Dashboard").Cells(139 + xrow, CompleteYears + 4 + ycolumn) <> 0 Then
            Sheets("Dashboard").Cells(139 + xrow, CompleteYears + Rolltwelveyears * 2 + 8 + ycolumn) = Sheets("Dashboard").Cells(139 + xrow, CompleteYears + Rolltwelveyears * 2 + 4 + ycolumn) / Sheets("Dashboard").Cells(139 + xrow, CompleteYears + 4 + ycolumn)
        Else
            Sheets("Dashboard").Cells(139 + xrow, CompleteYears + Rolltwelveyears * 2 + 8 + ycolumn) = 0
        End If
        
        xrow = xrow + 1
        
    Loop
    
    xrow = 0
    
    Do While xrow <= 4
        
        If Sheets("Dashboard").Cells(147 + xrow, CompleteYears + 4 + ycolumn) <> 0 Then
            Sheets("Dashboard").Cells(147 + xrow, CompleteYears + Rolltwelveyears * 2 + 8 + ycolumn) = Sheets("Dashboard").Cells(147 + xrow, CompleteYears + Rolltwelveyears * 2 + 4 + ycolumn) / Sheets("Dashboard").Cells(147 + xrow, CompleteYears + 4 + ycolumn)
        Else
            Sheets("Dashboard").Cells(147 + xrow, CompleteYears + Rolltwelveyears * 2 + 8 + ycolumn) = 0
        End If
        xrow = xrow + 1
        
    Loop

    ycolumn = ycolumn + 1
    
Loop


Sheets("Dashboard").Cells(137, CompleteYears + 3) = "Total Number of Medical Claims by Type - Rolling-12"
Sheets("Dashboard").Cells(137, CompleteYears + Rolltwelveyears + 5) = "Total Paid in Medical Claims by Type - Rolling-12"
Sheets("Dashboard").Cells(137, CompleteYears + Rolltwelveyears * 2 + 7) = "Total Paid in Medical Claims Unit Cost by Type - Rolling-12"
Sheets("Dashboard").Cells(145, CompleteYears + 3) = "Total Number of Vision Claims by Type - Rolling-12"
Sheets("Dashboard").Cells(145, CompleteYears + Rolltwelveyears + 5) = "Total Paid Vision Claims by Type - Rolling-12"
Sheets("Dashboard").Cells(145, CompleteYears + Rolltwelveyears * 2 + 7) = "Total Paid in Vision Claims Unit Cost by Type - Rolling-12"

Rows("138:138").Select
Selection.NumberFormat = "General"
Rows("146:146").Select
Selection.NumberFormat = "General"

Cells.Select
Selection.NumberFormat = "$#,##0"
Range(Cells(2, 2), Cells(2, 1 + CompleteYears)).Select
Selection.NumberFormat = "General"
Range(Cells(35, 2), Cells(35, 1 + CompleteYears)).Select
Selection.NumberFormat = "General"
Range(Cells(41, 2), Cells(41, 1 + CompleteYears)).Select
Selection.NumberFormat = "General"
Range(Cells(49, 2), Cells(49, 1 + CompleteYears)).Select
Selection.NumberFormat = "General"
Range(Cells(55, 2), Cells(55, 1 + CompleteYears)).Select
Selection.NumberFormat = "General"
Range(Cells(63, 2), Cells(63, 1 + CompleteYears)).Select
Selection.NumberFormat = "General"
Range(Cells(71, 2), Cells(71, 1 + CompleteYears)).Select
Selection.NumberFormat = "General"
Range(Cells(77, 2), Cells(77, 1 + CompleteYears)).Select
Selection.NumberFormat = "General"
Sheets("Dashboard").Range(Cells(35, CompleteYears + Rolltwelveyears + 8), Cells(35, CompleteYears * 2 + Rolltwelveyears + 8)).Select
Selection.NumberFormat = "General"
Sheets("Dashboard").Range(Cells(2, CompleteYears + Rolltwelveyears + 8), Cells(2, CompleteYears * 2 + Rolltwelveyears + 8)).Select
Selection.NumberFormat = "General"
Columns("A:Z").ColumnWidth = 8.43

Range(Cells(139, CompleteYears + 4), Cells(143, CompleteYears + Rolltwelveyears + 4)).Select
Selection.NumberFormat = "General"
Range(Cells(147, CompleteYears + 4), Cells(151, CompleteYears + Rolltwelveyears + 4)).Select
Selection.NumberFormat = "General"

If CompleteYears = 0 Then
    GoTo NoGraphs
End If

Call GraphMaker

GoTo EndNoError

NoGraphs:
If CompleteYears = 0 Then
    MsgBox ("You must have at least one year of data to make graphs.")
End If

Handler:
If HaveMedical = True Or HaveVision = True Then
    MsgBox ("There was an error.")
Else
    MsgBox ("You must have medical or vision coverage.")
End If

EndNoError:
End Sub

Sub GraphMaker()

Dim cht As ChartObject, ChartCount As Integer, Brightness As Double, HaveMedical As Boolean, HaveVision As Boolean, CompleteYears As Integer, SLoop As Integer
Dim Pollsheet As String, PollNCells As Integer, StartYearStr As String, StartYear As Integer, FinalYearStr As String, FinalYear As Integer, StartMonth As Integer, StartMonthStr As String, FinalMonth As Integer, FinalMonthStr As String, NYears As Integer

On Error GoTo Handler

HaveMedical = False
HaveVision = False

For Each sh In Application.Worksheets
    If sh.Name = "MDCLMS" Then
        HaveMedical = True
    End If
    If sh.Name = "VSCLMS" Then
        HaveVision = True
    End If
Next sh

If HaveMedical = True Then
    Pollsheet = "MDCLMS"
Else
    Pollsheet = "VSCLMS"
End If

PollNCells = Application.WorksheetFunction.CountA(Sheets(Pollsheet).Range("A:A")) + 1

StartYearStr = Left(Sheets(Pollsheet).Cells(5, 1), 4)
StartYear = StartYearStr

FinalYearStr = Left(Sheets(Pollsheet).Cells(PollNCells, 1), 4)
FinalYear = FinalYearStr

If Mid(Sheets(Pollsheet).Cells(5, 1), Len(Sheets(Pollsheet).Cells(5, 1)) - 1, 1) = "/" Then
    StartMonthStr = Right(Sheets(Pollsheet).Cells(5, 1), 1)
Else
    StartMonthStr = Right(Sheets(Pollsheet).Cells(5, 1), 2)
End If
StartMonth = StartMonthStr

If Mid(Sheets(Pollsheet).Cells(5, 1), Len(Sheets(Pollsheet).Cells(PollNCells, 1)) - 1, 1) = "/" Then
    FinalMonthStr = Right(Sheets(Pollsheet).Cells(5, 1), 1)
Else
    FinalMonthStr = Right(Sheets(Pollsheet).Cells(5, 1), 2)
End If
FinalMonth = FinalMonthStr

NYears = FinalYear - StartYear + 1

CompleteYears = NYears

If StartMonth <> 1 Then
    CompleteYears = CompleteYears - 1
End If
If FinalMonth <> 12 Then
    CompleteYears = CompleteYears - 1
End If

Sheets("Dashboard").Activate

    Sheets.Add After:=ActiveSheet
    ActiveSheet.Select
    ActiveSheet.Name = "Dashboard Graphs"

If HaveMedical = True Then

    Sheets("Dashboard Graphs").Activate
    
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Select
        ActiveSheet.Name = "Medical Graphs"
    
    Sheets("Medical Graphs").Activate
    
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Select
        ActiveSheet.Name = "Rx Graphs"
        
End If

If HaveVision = True Then

    Sheets("Rx Graphs").Activate
    
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Select
        ActiveSheet.Name = "Vision Graphs"
        
End If

Sheets("Dashboard").Activate

ChartCount = ActiveSheet.ChartObjects.Count

If ChartCount > 0 Then
    ActiveSheet.ChartObjects.Delete
End If

Set co = Sheets("Dashboard").ChartObjects.Add(1, 1, 1, 1)
ActiveSheet.ChartObjects(1).Activate
co.Chart.SetSourceData Source:=Sheets("Dashboard").Range(Cells(35, 1), Cells(38, CompleteYears + 1))
co.Chart.ChartType = xlColumnStacked
co.Chart.HasTitle = True
ActiveChart.ChartTitle.Select
ActiveChart.ChartTitle.Text = "Total Paid in Claims - Complete Years Only"
ActiveChart.PlotBy = xlRows
Set cht = co.Chart.Parent
With cht
.Left = Sheets("Dashboard").Cells(1, 1).Left
.Top = Sheets("Dashboard").Cells(1, 1).Top
.Height = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Height
.Width = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Width
End With
ActiveChart.FullSeriesCollection(1).ApplyDataLabels
ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(59, 110, 172)
        .Transparency = 0
        .Solid
    End With
ActiveChart.FullSeriesCollection(2).ApplyDataLabels
ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(233, 131, 0)
        .Transparency = 0
        .Solid
    End With
ActiveChart.Legend.Select
Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
ActiveChart.FullSeriesCollection(3).ApplyDataLabels
ActiveChart.FullSeriesCollection(3).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(188, 211, 67)
        .Transparency = 0
        .Solid
    End With
SLoop = 1
Do While SLoop <= 3
    ActiveChart.FullSeriesCollection(SLoop).DataLabels.Select
    Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
    SLoop = SLoop + 1
Loop
ActiveChart.Parent.Cut
Sheets("Dashboard Graphs").Select
ActiveSheet.Paste


Sheets("Dashboard").Activate

Set co = Sheets("Dashboard").ChartObjects.Add(1, 1, 1, 1)
ActiveSheet.ChartObjects(1).Activate
co.Chart.SetSourceData Source:=Sheets("Dashboard").Range(Cells(2, 1), Cells(5, CompleteYears + 1))
co.Chart.ChartType = xlColumnStacked
co.Chart.HasTitle = True
ActiveChart.ChartTitle.Select
ActiveChart.ChartTitle.Text = "Total Paid in Claims PEPY - Complete Years Only"
ActiveChart.PlotBy = xlRows
Set cht = co.Chart.Parent
With cht
.Left = Sheets("Dashboard").Cells(1, 1).Left
.Top = Sheets("Dashboard").Cells(1, 1).Top
.Height = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Height
.Width = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Width
End With
ActiveChart.FullSeriesCollection(1).ApplyDataLabels
ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(59, 110, 172)
        .Transparency = 0
        .Solid
    End With
ActiveChart.FullSeriesCollection(2).ApplyDataLabels
ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(233, 131, 0)
        .Transparency = 0
        .Solid
    End With
ActiveChart.FullSeriesCollection(3).ApplyDataLabels
ActiveChart.FullSeriesCollection(3).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(188, 211, 67)
        .Transparency = 0
        .Solid
    End With
ActiveChart.Legend.Select
Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
SLoop = 1
Do While SLoop <= 3
    ActiveChart.FullSeriesCollection(SLoop).DataLabels.Select
    Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
    SLoop = SLoop + 1
Loop
ActiveChart.Parent.Cut
Sheets("Dashboard Graphs").Select
Cells(1, 9).Select
ActiveSheet.Paste

Sheets("Dashboard").Activate

Set co = Sheets("Dashboard").ChartObjects.Add(1, 1, 1, 1)
ActiveSheet.ChartObjects(1).Activate
co.Chart.SetSourceData Source:=Sheets("Dashboard").Range(Cells(8, 1), Cells(11, 13))
co.Chart.ChartType = xlLine
co.Chart.HasTitle = True
ActiveChart.ChartTitle.Select
ActiveChart.ChartTitle.Text = "Total Paid in Claims PMPM - Past 12 Months"
ActiveChart.PlotBy = xlRows
Set cht = co.Chart.Parent
With cht
.Left = Sheets("Dashboard").Cells(1, 1).Left
.Top = Sheets("Dashboard").Cells(1, 1).Top
.Height = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Height
.Width = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Width
End With
ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(59, 110, 172)
        .Transparency = 0
    End With
ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(233, 131, 0)
        .Transparency = 0
    End With
ActiveChart.FullSeriesCollection(3).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(188, 211, 67)
        .Transparency = 0
    End With
ActiveChart.Legend.Select
Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
ActiveChart.Parent.Cut
Sheets("Dashboard Graphs").Select
Cells(17, 1).Select
ActiveSheet.Paste

Sheets("Dashboard").Activate

Set co = Sheets("Dashboard").ChartObjects.Add(1, 1, 1, 1)
ActiveSheet.ChartObjects(1).Activate
co.Chart.SetSourceData Source:=Sheets("Dashboard").Range(Cells(14, 1), Cells(16, 13))
co.Chart.ChartType = xlLine
co.Chart.HasTitle = True
ActiveChart.ChartTitle.Select
ActiveChart.ChartTitle.Text = "Total Paid in Claims Unit Cost - Past 12 Months"
ActiveChart.PlotBy = xlRows
Set cht = co.Chart.Parent
With cht
.Left = Sheets("Dashboard").Cells(1, 1).Left
.Top = Sheets("Dashboard").Cells(1, 1).Top
.Height = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Height
.Width = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Width
End With
ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(59, 110, 172)
        .Transparency = 0
    End With
ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(188, 211, 67)
        .Transparency = 0
    End With
ActiveChart.Legend.Select
Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
SLoop = 1
ActiveChart.Parent.Cut
Sheets("Dashboard Graphs").Select
Cells(17, 9).Select
ActiveSheet.Paste

Sheets("Dashboard").Activate

If HaveMedical = True Then
    
    Set co = Sheets("Dashboard").ChartObjects.Add(1, 1, 1, 1)
    ActiveSheet.ChartObjects(1).Activate
    co.Chart.SetSourceData Source:=Sheets("Dashboard").Range(Cells(41, 1), Cells(46, 1 + CompleteYears))
    co.Chart.ChartType = xlColumnStacked
    co.Chart.HasTitle = True
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Total Paid in Medical Claims by Type - Complete Years Only"
    ActiveChart.PlotBy = xlRows
    Set cht = co.Chart.Parent
    With cht
    .Left = Sheets("Dashboard").Cells(1, 1).Left
    .Top = Sheets("Dashboard").Cells(1, 1).Top
    .Height = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Height
    .Width = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Width
    End With
    ActiveChart.Legend.Select
    Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
    SLoop = 1
    Brightness = 0.8
    Do While SLoop < 6
        ActiveChart.FullSeriesCollection(SLoop).ApplyDataLabels
        ActiveChart.FullSeriesCollection(SLoop).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(59, 110, 172)
                .ForeColor.Brightness = Brightness
            End With
        Brightness = Brightness - 0.2625
        SLoop = SLoop + 1
    Loop
    SLoop = 1
    Do While SLoop <= 5
        ActiveChart.FullSeriesCollection(SLoop).DataLabels.Select
        Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
        SLoop = SLoop + 1
    Loop
    ActiveChart.Parent.Cut
    Sheets("Medical Graphs").Select
    Cells(1, 1).Select
    ActiveSheet.Paste
    
    Sheets("Dashboard").Activate
    
    Set co = Sheets("Dashboard").ChartObjects.Add(1, 1, 1, 1)
    ActiveSheet.ChartObjects(1).Activate
    co.Chart.SetSourceData Source:=Sheets("Dashboard").Range(Cells(63, 1), Cells(68, 1 + CompleteYears))
    co.Chart.ChartType = xlColumnStacked
    co.Chart.HasTitle = True
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Total Paid in Medical Claims PEPY by Type - Complete Years Only"
    ActiveChart.PlotBy = xlRows
    Set cht = co.Chart.Parent
    With cht
    .Left = Sheets("Dashboard").Cells(1, 1).Left
    .Top = Sheets("Dashboard").Cells(1, 1).Top
    .Height = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Height
    .Width = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Width
    End With
    ActiveChart.Legend.Select
    Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
    SLoop = 1
    Brightness = 0.8
    Do While SLoop < 6
        ActiveChart.FullSeriesCollection(SLoop).ApplyDataLabels
        ActiveChart.FullSeriesCollection(SLoop).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(59, 110, 172)
                .ForeColor.Brightness = Brightness
            End With
        Brightness = Brightness - 0.2625
        SLoop = SLoop + 1
    Loop
    SLoop = 1
    Do While SLoop <= 5
        ActiveChart.FullSeriesCollection(SLoop).DataLabels.Select
        Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
        SLoop = SLoop + 1
    Loop
    ActiveChart.Parent.Cut
    Sheets("Medical Graphs").Select
    Cells(1, 9).Select
    ActiveSheet.Paste
    
    Sheets("Dashboard").Activate
    
    Set co = Sheets("Dashboard").ChartObjects.Add(1, 1, 1, 1)
    ActiveSheet.ChartObjects(1).Activate
    co.Chart.SetSourceData Source:=Sheets("Dashboard").Range(Cells(85, 1), Cells(90, 13))
    co.Chart.ChartType = xlLine
    co.Chart.HasTitle = True
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Total Paid in Medical Claims PMPM - Past 12 Months"
    ActiveChart.PlotBy = xlRows
    Set cht = co.Chart.Parent
    With cht
    .Left = Sheets("Dashboard").Cells(1, 1).Left
    .Top = Sheets("Dashboard").Cells(1, 1).Top
    .Height = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Height
    .Width = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Width
    End With
    ActiveChart.Legend.Select
    Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
    SLoop = 1
    Brightness = 0.8
    Do While SLoop < 6
        ActiveChart.FullSeriesCollection(SLoop).Select
            With Selection.Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(59, 110, 172)
                .ForeColor.Brightness = Brightness
            End With
        Brightness = Brightness - 0.2625
        SLoop = SLoop + 1
    Loop
    SLoop = 1
    ActiveChart.Parent.Cut
    Sheets("Medical Graphs").Select
    Cells(17, 1).Select
    ActiveSheet.Paste
        
    Sheets("Dashboard").Activate
    
    Set co = Sheets("Dashboard").ChartObjects.Add(1, 1, 1, 1)
    ActiveSheet.ChartObjects(1).Activate
    co.Chart.SetSourceData Source:=Sheets("Dashboard").Range(Cells(19, 1), Cells(24, 13))
    co.Chart.ChartType = xlLine
    co.Chart.HasTitle = True
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Total Paid in Medical Claims Unit Cost - Past 12 Months"
    ActiveChart.PlotBy = xlRows
    Set cht = co.Chart.Parent
    With cht
    .Left = Sheets("Dashboard").Cells(1, 1).Left
    .Top = Sheets("Dashboard").Cells(1, 1).Top
    .Height = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Height
    .Width = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Width
    End With
    ActiveChart.Legend.Select
    Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
    SLoop = 1
    Brightness = 0.8
    Do While SLoop < 6
        ActiveChart.FullSeriesCollection(SLoop).Select
            With Selection.Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(59, 110, 172)
                .ForeColor.Brightness = Brightness
            End With
        Brightness = Brightness - 0.2625
        SLoop = SLoop + 1
    Loop
    ActiveChart.Parent.Cut
    Sheets("Medical Graphs").Select
    Cells(17, 9).Select
    ActiveSheet.Paste
    
    Sheets("Dashboard").Activate
    
    Set co = Sheets("Dashboard").ChartObjects.Add(1, 1, 1, 1)
    ActiveSheet.ChartObjects(1).Activate
    co.Chart.SetSourceData Source:=Sheets("Dashboard").Range(Cells(49, 1), Cells(52, 1 + CompleteYears))
    co.Chart.ChartType = xlColumnStacked
    co.Chart.HasTitle = True
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Total Paid in Rx Claims by Type - Complete Years Only"
    ActiveChart.PlotBy = xlRows
    Set cht = co.Chart.Parent
    With cht
    .Left = Sheets("Dashboard").Cells(1, 1).Left
    .Top = Sheets("Dashboard").Cells(1, 1).Top
    .Height = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Height
    .Width = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Width
    End With
    ActiveChart.Legend.Select
    Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
    SLoop = 1
    Brightness = 0.6
    Do While SLoop < 4
        ActiveChart.FullSeriesCollection(SLoop).ApplyDataLabels
        ActiveChart.FullSeriesCollection(SLoop).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(233, 131, 0)
                .ForeColor.Brightness = Brightness
            End With
        Brightness = Brightness - 0.3
        SLoop = SLoop + 1
    Loop
    SLoop = 1
    Do While SLoop <= 3
        ActiveChart.FullSeriesCollection(SLoop).DataLabels.Select
        Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
        SLoop = SLoop + 1
    Loop
    ActiveChart.Parent.Cut
    Sheets("Rx Graphs").Select
    Cells(1, 1).Select
    ActiveSheet.Paste
    
    Sheets("Dashboard").Activate
    
    Set co = Sheets("Dashboard").ChartObjects.Add(1, 1, 1, 1)
    ActiveSheet.ChartObjects(1).Activate
    co.Chart.SetSourceData Source:=Sheets("Dashboard").Range(Cells(71, 1), Cells(74, 1 + CompleteYears))
    co.Chart.ChartType = xlColumnStacked
    co.Chart.HasTitle = True
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Total Paid in Rx Claims PEPY by Type - Complete Years Only"
    ActiveChart.PlotBy = xlRows
    Set cht = co.Chart.Parent
    With cht
    .Left = Sheets("Dashboard").Cells(1, 1).Left
    .Top = Sheets("Dashboard").Cells(1, 1).Top
    .Height = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Height
    .Width = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Width
    End With
    ActiveChart.Legend.Select
    Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
    SLoop = 1
    Brightness = 0.6
    Do While SLoop < 4
        ActiveChart.FullSeriesCollection(SLoop).ApplyDataLabels
        ActiveChart.FullSeriesCollection(SLoop).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(233, 131, 0)
                .ForeColor.Brightness = Brightness
            End With
        Brightness = Brightness - 0.3
        SLoop = SLoop + 1
    Loop
    SLoop = 1
    Do While SLoop <= 3
        ActiveChart.FullSeriesCollection(SLoop).DataLabels.Select
        Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
        SLoop = SLoop + 1
    Loop
    ActiveChart.Parent.Cut
    Sheets("Rx Graphs").Select
    Cells(1, 9).Select
    ActiveSheet.Paste
    
    Sheets("Dashboard").Activate
    
    Set co = Sheets("Dashboard").ChartObjects.Add(1, 1, 1, 1)
    ActiveSheet.ChartObjects(1).Activate
    co.Chart.SetSourceData Source:=Sheets("Dashboard").Range(Cells(93, 1), Cells(96, 13))
    co.Chart.ChartType = xlLine
    co.Chart.HasTitle = True
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Total Paid in Rx Claims PMPM - Past 12 Months"
    ActiveChart.PlotBy = xlRows
    Set cht = co.Chart.Parent
    With cht
    .Left = Sheets("Dashboard").Cells(1, 1).Left
    .Top = Sheets("Dashboard").Cells(1, 1).Top
    .Height = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Height
    .Width = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Width
    End With
    ActiveChart.Legend.Select
    Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
    SLoop = 1
    Brightness = 0.6
    Do While SLoop < 4
        ActiveChart.FullSeriesCollection(SLoop).Select
            With Selection.Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(233, 131, 0)
                .ForeColor.Brightness = Brightness
            End With
        Brightness = Brightness - 0.3
        SLoop = SLoop + 1
    Loop
    ActiveChart.Parent.Cut
    Sheets("Rx Graphs").Select
    Cells(17, 1).Select
    ActiveSheet.Paste
    
End If

If HaveVision = True Then

    Sheets("Dashboard").Activate
    
    Set co = Sheets("Dashboard").ChartObjects.Add(1, 1, 1, 1)
    ActiveSheet.ChartObjects(1).Activate
    co.Chart.SetSourceData Source:=Sheets("Dashboard").Range(Cells(55, 1), Cells(60, 1 + CompleteYears))
    co.Chart.ChartType = xlColumnStacked
    co.Chart.HasTitle = True
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Total Paid in Vision Claims by Type - Complete Years Only"
    ActiveChart.PlotBy = xlRows
    Set cht = co.Chart.Parent
    With cht
    .Left = Sheets("Dashboard").Cells(1, 1).Left
    .Top = Sheets("Dashboard").Cells(1, 1).Top
    .Height = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Height
    .Width = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Width
    End With
    SLoop = 1
    Brightness = 0.6
    Do While SLoop < 6
        ActiveChart.FullSeriesCollection(SLoop).ApplyDataLabels
        ActiveChart.FullSeriesCollection(SLoop).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(188, 211, 67)
                .ForeColor.Brightness = Brightness
            End With
        Brightness = Brightness - 0.3
        SLoop = SLoop + 1
    Loop
    ActiveChart.Legend.Select
    Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
    SLoop = 1
    Do While SLoop <= 5
        ActiveChart.FullSeriesCollection(SLoop).DataLabels.Select
        Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
        SLoop = SLoop + 1
    Loop
    ActiveChart.Parent.Cut
    Sheets("Vision Graphs").Select
    Cells(1, 1).Select
    ActiveSheet.Paste
    
    Sheets("Dashboard").Activate
    
    Set co = Sheets("Dashboard").ChartObjects.Add(1, 1, 1, 1)
    ActiveSheet.ChartObjects(1).Activate
    co.Chart.SetSourceData Source:=Sheets("Dashboard").Range(Cells(77, 1), Cells(82, 1 + CompleteYears))
    co.Chart.ChartType = xlColumnStacked
    co.Chart.HasTitle = True
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Total Paid in Vision Claims PEPY by Type - Complete Years Only"
    ActiveChart.PlotBy = xlRows
    Set cht = co.Chart.Parent
    With cht
    .Left = Sheets("Dashboard").Cells(1, 1).Left
    .Top = Sheets("Dashboard").Cells(1, 1).Top
    .Height = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Height
    .Width = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Width
    End With
    SLoop = 1
    Brightness = 0.6
    Do While SLoop < 6
        ActiveChart.FullSeriesCollection(SLoop).ApplyDataLabels
        ActiveChart.FullSeriesCollection(SLoop).Select
            With Selection.Format.Fill
                .Visible = msoTrue
                .ForeColor.RGB = RGB(188, 211, 67)
                .ForeColor.Brightness = Brightness
            End With
        Brightness = Brightness - 0.3
        SLoop = SLoop + 1
    Loop
    ActiveChart.Legend.Select
    Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
    SLoop = 1
    Do While SLoop <= 5
        ActiveChart.FullSeriesCollection(SLoop).DataLabels.Select
        Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
        SLoop = SLoop + 1
    Loop
    ActiveChart.Parent.Cut
    Sheets("Vision Graphs").Select
    Cells(1, 9).Select
    ActiveSheet.Paste
    
    Sheets("Dashboard").Activate
    
    Set co = Sheets("Dashboard").ChartObjects.Add(1, 1, 1, 1)
    ActiveSheet.ChartObjects(1).Activate
    co.Chart.SetSourceData Source:=Sheets("Dashboard").Range(Cells(99, 1), Cells(104, 13))
    co.Chart.ChartType = xlLine
    co.Chart.HasTitle = True
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Total Paid in Vision Claims PMPM - Past 12 Months"
    ActiveChart.PlotBy = xlRows
    Set cht = co.Chart.Parent
    With cht
    .Left = Sheets("Dashboard").Cells(1, 1).Left
    .Top = Sheets("Dashboard").Cells(1, 1).Top
    .Height = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Height
    .Width = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Width
    End With
    ActiveChart.Legend.Select
    Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
    SLoop = 1
    Brightness = 0.6
    Do While SLoop < 6
        ActiveChart.FullSeriesCollection(SLoop).Select
            With Selection.Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(188, 211, 67)
                .ForeColor.Brightness = Brightness
            End With
        Brightness = Brightness - 0.3
        SLoop = SLoop + 1
    Loop
    ActiveChart.Parent.Cut
    Sheets("Vision Graphs").Select
    Cells(17, 1).Select
    ActiveSheet.Paste
        
    Sheets("Dashboard").Activate
    
    Set co = Sheets("Dashboard").ChartObjects.Add(1, 1, 1, 1)
    ActiveSheet.ChartObjects(1).Activate
    co.Chart.SetSourceData Source:=Sheets("Dashboard").Range(Cells(27, 1), Cells(32, 13))
    co.Chart.ChartType = xlLine
    co.Chart.HasTitle = True
    ActiveChart.ChartTitle.Select
    ActiveChart.ChartTitle.Text = "Total Paid in Vision Claims Unit Cost - Past 12 Months"
    ActiveChart.PlotBy = xlRows
    Set cht = co.Chart.Parent
    With cht
    .Left = Sheets("Dashboard").Cells(1, 1).Left
    .Top = Sheets("Dashboard").Cells(1, 1).Top
    .Height = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Height
    .Width = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Width
    End With
    ActiveChart.Legend.Select
    Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
    SLoop = 1
    Brightness = 0.6
    Do While SLoop < 6
        ActiveChart.FullSeriesCollection(SLoop).Select
            With Selection.Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = RGB(188, 211, 67)
                .ForeColor.Brightness = Brightness
            End With
        Brightness = Brightness - 0.3
        SLoop = SLoop + 1
    Loop
    ActiveChart.Parent.Cut
    Sheets("Vision Graphs").Select
    Cells(17, 9).Select
    ActiveSheet.Paste

End If

Dim NoError As Boolean

NoError = True
If NoError = True Then
    GoTo EndNoError
End If
Handler:
If HaveMedical = True Or HaveVision = True Then
    MsgBox ("There was an error.")
Else
    MsgBox ("You must have medical or vision coverage.")
End If

EndNoError:
End Sub

Sub Graphmaker2()

Dim cht As ChartObject, ChartCount As Integer, Brightness As Double, HaveMedical As Boolean, HaveVision As Boolean, CompleteYears As Integer, SLoop As Integer
Dim Pollsheet As String, PollNCells As Integer, StartYearStr As String, StartYear As Integer, FinalYearStr As String, FinalYear As Integer, StartMonth As Integer, StartMonthStr As String, FinalMonth As Integer, FinalMonthStr As String, NYears As Integer

On Error GoTo Handler

HaveMedical = False
HaveVision = False

For Each sh In Application.Worksheets
    If sh.Name = "MDCLMS" Then
        HaveMedical = True
    End If
    If sh.Name = "VSCLMS" Then
        HaveVision = True
    End If
Next sh

If HaveMedical = True Then
    Pollsheet = "MDCLMS"
Else
    Pollsheet = "VSCLMS"
End If

PollNCells = Application.WorksheetFunction.CountA(Sheets(Pollsheet).Range("A:A")) + 1

StartYearStr = Left(Sheets(Pollsheet).Cells(5, 1), 4)
StartYear = StartYearStr

FinalYearStr = Left(Sheets(Pollsheet).Cells(PollNCells, 1), 4)
FinalYear = FinalYearStr

If Mid(Sheets(Pollsheet).Cells(5, 1), Len(Sheets(Pollsheet).Cells(5, 1)) - 1, 1) = "/" Then
    StartMonthStr = Right(Sheets(Pollsheet).Cells(5, 1), 1)
Else
    StartMonthStr = Right(Sheets(Pollsheet).Cells(5, 1), 2)
End If
StartMonth = StartMonthStr

If Mid(Sheets(Pollsheet).Cells(5, 1), Len(Sheets(Pollsheet).Cells(PollNCells, 1)) - 1, 1) = "/" Then
    FinalMonthStr = Right(Sheets(Pollsheet).Cells(5, 1), 1)
Else
    FinalMonthStr = Right(Sheets(Pollsheet).Cells(5, 1), 2)
End If
FinalMonth = FinalMonthStr

NYears = FinalYear - StartYear + 1

CompleteYears = NYears

If StartMonth <> 1 Then
    CompleteYears = CompleteYears - 1
End If
If FinalMonth <> 12 Then
    CompleteYears = CompleteYears - 1
End If

Sheets("Dashboard").Activate

    Sheets.Add After:=ActiveSheet
    ActiveSheet.Select
    ActiveSheet.Name = "Dashboard Graphs"

If HaveMedical = True Then

    Sheets("Dashboard Graphs").Activate
    
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Select
        ActiveSheet.Name = "Medical Graphs"
    
    Sheets("Medical Graphs").Activate
    
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Select
        ActiveSheet.Name = "Rx Graphs"
        
End If

If HaveVision = True Then

    Sheets("Rx Graphs").Activate
    
        Sheets.Add After:=ActiveSheet
        ActiveSheet.Select
        ActiveSheet.Name = "Vision Graphs"
        
End If

Sheets("Dashboard").Activate

ChartCount = ActiveSheet.ChartObjects.Count

If ChartCount > 0 Then
    ActiveSheet.ChartObjects.Delete
End If

Set co = Sheets("Dashboard").ChartObjects.Add(1, 1, 1, 1)
ActiveSheet.ChartObjects(1).Activate
co.Chart.SetSourceData Source:=Sheets("Dashboard").Range(Cells(35, 1), Cells(38, CompleteYears + 1))
co.Chart.ChartType = xlColumnStacked
co.Chart.HasTitle = True
ActiveChart.ChartTitle.Select
ActiveChart.ChartTitle.Text = "Total Paid in Claims - Complete Years Only"
ActiveChart.PlotBy = xlRows
Set cht = co.Chart.Parent
With cht
.Left = Sheets("Dashboard").Cells(1, 1).Left
.Top = Sheets("Dashboard").Cells(1, 1).Top
.Height = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Height
.Width = Sheets("Dashboard").Range(Cells(1, 1), Cells(15, 7)).Width
End With
ActiveChart.FullSeriesCollection(1).ApplyDataLabels
ActiveChart.FullSeriesCollection(1).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(59, 110, 172)
        .Transparency = 0
        .Solid
    End With
ActiveChart.FullSeriesCollection(2).ApplyDataLabels
ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(233, 131, 0)
        .Transparency = 0
        .Solid
    End With
ActiveChart.Legend.Select
Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
ActiveChart.FullSeriesCollection(3).ApplyDataLabels
ActiveChart.FullSeriesCollection(3).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(188, 211, 67)
        .Transparency = 0
        .Solid
    End With
SLoop = 1
Do While SLoop <= 3
    ActiveChart.FullSeriesCollection(SLoop).DataLabels.Select
    Selection.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(50, 50, 50)
    SLoop = SLoop + 1
Loop
ActiveChart.Parent.Cut
Sheets("Dashboard Graphs").Select
ActiveSheet.Paste

End Sub
