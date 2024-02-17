Sub ErrorCheck()
    ' Constants
    Dim RADAR_SITES As Variant
    RADAR_SITES = Array("Site1", "Site2", "Site3", "Site4")
    Dim EXPECTED_VALUES As Variant
    EXPECTED_VALUES = Array("Query", "Id", "Trk", "Time", "MsgType", "Rng(nmi)", "Az(deg)", "Hgt(ft)", "MC(ft)", "MCV", "M3", "M3V", "M2", "M2V", "Lat", "Lon", "DecLat", "DecLon", "", "", "Af", "Date", "Disc", "Ident", "M2X", "M3X", "M4", "NonAf", "RL", "Tst", "TIS(sec)")

    ' Variables
    Dim col As Integer
    Dim row As Integer
    Dim messageType As String
    Dim siteNameAJ As String
    Dim Error1Int As Integer
    Dim Error1CumInt As Integer
    Dim Error2Int As Integer
    Dim MissingColumnsSTR As String

    ' Set error to 0
    Error1Int = 0
    Error1CumInt = 0
    Error2Int = 0
    
    ' Fill radar sites from RADAR_SITES to the AI column
    For row = 0 To UBound(RADAR_SITES)
        Cells(row + 3, "AI").Value = RADAR_SITES(row)
    Next row

    ' Copy unique values from column B to column AJ
    Columns("B:B").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("AJ2"), Unique:=True

    ' Sort range AJ3:AJ16 in ascending order
    With ActiveSheet.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("AJ3:AJ" & ActiveSheet.Cells(ActiveSheet.Rows.Count, "AJ").End(xlUp).row), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
        .SetRange Range("AJ3:AJ16")
        .header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' Initialize dictionary to store radar sites
    Dim radarSiteDict As Object
    Set radarSiteDict = CreateObject("Scripting.Dictionary")

    ' Add each expected radar site to dictionary
    For Each site In RADAR_SITES
        radarSiteDict(site) = False
    Next site

    ' Loop over cells in column AJ to check for radar site IDs
    Dim lastRowAJ As Integer
    lastRowAJ = ActiveSheet.Cells(ActiveSheet.Rows.Count, "AJ").End(xlUp).row

    ' Initialize flag for extra sites
    Dim extraSites As Boolean: extraSites = False

    For row = 3 To lastRowAJ
        siteNameAJ = Cells(row, "AJ").Value
        If radarSiteDict.Exists(siteNameAJ) Then
            radarSiteDict(siteNameAJ) = True
        Else
            extraSites = True
        End If
    Next row

    ' Loop over each item in dictionary to check if any expected radar site is missing
    For Each site In radarSiteDict.Keys
        If radarSiteDict(site) = False Then
            ' If an expected radar site is missing, mark cell in column AI with yellow fill
            With Cells(Application.WorksheetFunction.Match(site, RADAR_SITES, 0) + 2, "AI").Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Error1CumInt = 1
        End If
    Next site

    ' Update Error1Int based on whether any expected radar sites are missing and whether any extra sites are found
    If Error1CumInt = 1 And extraSites Then
        Error1Int = 3
    ElseIf Error1CumInt = 1 Then
        Error1Int = 1
    ElseIf extraSites Then
        Error1Int = 2
    End If

    ' Display message box for Error1Int
    If Error1Int = 1 Then
        MsgBox ("There is a mismatch between the expected radar site IDs and the radar site IDs in the CSV file")
    ElseIf Error1Int = 2 Then
        MsgBox ("The CSV file contains more radar sites than expected")
    ElseIf Error1Int = 3 Then
        MsgBox ("There is a mismatch between the expected radar site IDs and the radar site IDs in the CSV file and there are more radar sites than expected")
    End If

    ' Copy unique values from column E to column AL
    Columns("E:E").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Range("AL2"), Unique:=True

    ' Loop over rows 3 to 5 in column AL
    For row = 3 To 5
        messageType = Cells(row, "AL").Value
        If messageType = "SRTQC" Then Error2Int = Error2Int + 1
        If messageType = "BRTQC" Then Error2Int = Error2Int + 2
    Next row

    ' Display message box for Error2Int
    If Error2Int = 1 Then
        MsgBox ("File is missing BRTQC messages")
    ElseIf Error2Int = 2 Then
        MsgBox ("File is missing SRTQC messages")
    ElseIf Error2Int = 3 Then
        MsgBox ("File has only BRTQC and SRTQC Messages")
    ElseIf Error2Int >= 4 Then
        MsgBox ("Check file - it appears to contain too many different message types.")
    End If

    ' Loop over columns A to AE
    For col = 1 To 31
        ' Check if value of current cell matches expected value
        If Cells(1, col).Value <> EXPECTED_VALUES(col - 1) Then
            ' If value does not match, add column letter to MissingColumnsSTR
            MissingColumnsSTR = MissingColumnsSTR & Split(Cells(1, col).Address, "$")(1) & ", "
        End If
    Next col

    ' Display message box indicating whether any columns are missing or not
    If MissingColumnsSTR = "" Then
        MsgBox ("No Missing Columns")
    Else
        MsgBox ("Missing Columns: " & Left(MissingColumnsSTR, Len(MissingColumnsSTR) - 2))
    End If
End Sub

Sub CreateNewWorkbook()
    ' Variables and Constants
    Dim FirstSheetName As String
    Dim activeWorkbookName As String
    Dim activeWorkbookPath As String
    Dim fileNameLength As Integer
    Dim newFileNameLength As Integer
    Dim newWorkbookName As String
    Const fileExtensionLength As Integer = 4

    ' Array of sheet names
    Dim sheetNames As Variant
    sheetNames = Array("Site1", "Site2", "Site3", "Site4", "Results")

    ' Array of cell values
    Dim cellLocations As Variant
    cellLocations = Array(Array("C1", "SiteName"), Array("H1", "SiteName"), Array("M1", "SiteName"), Array("R1", "SiteName"))
    
    ' Array of cell ranges to merge
    Dim cellLocationRanges As Variant
    cellLocationRanges = Array("A1:E1", "F1:J1", "K1:O1", "P1:T1")
   
    ' Array of cell values
    Dim cellValues As Variant
    cellValues = Array("Start Date", "Start Time", "End Date", "End Time", "Outage Duration")
   
    ' Array of cell ranges
    Dim cellRanges As Variant
    cellRanges = Array("A2:E2", "F2:J2", "K2:O2", "P2:T2")

    ' Array of column letters and their corresponding number formats
    Dim colFormats As Variant
    colFormats = Array(Array("A", "mm/dd/yyyy"), Array("B", "hh:mm:ss.0"), Array("C", "mm/dd/yyyy"), _
                       Array("D", "hh:mm:ss.0"), Array("E", "hh:mm:ss.0"), Array("F", "mm/dd/yyyy"), _
                       Array("G", "hh:mm:ss.0"), Array("H", "mm/dd/yyyy"), Array("I", "hh:mm:ss.0"), _
                       Array("J", "hh:mm:ss.0"), Array("K", "mm/dd/yyyy"), Array("L", "hh:mm:ss.0"), _
                       Array("M", "mm/dd/yyyy"), Array("N", "hh:mm:ss.0"), Array("O", "hh:mm:ss.0"), _
                       Array("P", "mm/dd/yyyy"), Array("Q", "hh:mm:ss.0"), Array("R", "mm/dd/yyyy"), _
                       Array("S", "hh:mm:ss.0"), Array("T", "hh:mm:ss.0"), Array("U", "mm/dd/yyyy"), _
                       Array("V", "hh:mm:ss.0"), Array("W", "mm/dd/yyyy"), Array("Y", "hh:mm:ss.0"))

    ' Get the name of the first sheet
    FirstSheetName = ActiveSheet.Name

    ' Get the name and path of the active workbook
    activeWorkbookName = ActiveWorkbook.Name
    activeWorkbookPath = ActiveWorkbook.Path

    ' Get the length of the workbook name
    fileNameLength = Len(activeWorkbookName)

    ' Calculate the length of the new workbook name
    newFileNameLength = fileNameLength - fileExtensionLength

    ' Get the new workbook name by removing the last 4 characters (file extension)
    newWorkbookName = Left(activeWorkbookName, newFileNameLength)
   
    ' Add the new file extension to the workbook name
    newWorkbookName = newWorkbookName + ".xlsx"

    ' Change the current directory to the path of the active workbook
    ChDir activeWorkbookPath
       
    ' Save the active workbook with the new name and format
    ActiveWorkbook.SaveAs Filename:=newWorkbookName, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    ' Format column D as time
    Columns("D:D").NumberFormat = "hh:mm:ss.0"
   
    ' Insert a new column to the right of column E
    Columns("E:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
   
    ' Format the new column as time
    Columns("E:E").NumberFormat = "hh:mm:ss.0"
   
    ' Add the text "Elapsed Time" to cell E1
    Range("E1").Value = "Elapsed Time"
   
    ' Loop through the sheetNames array
    For Each sheetName In sheetNames
        ' Add a new sheet after the active sheet
        Sheets.Add After:=ActiveSheet
        ' Rename the new sheet to the current value in the sheetNames array
        ActiveSheet.Name = sheetName
    Next sheetName
   
    ' Loop over each cell value in the array
    For Each cellLocation In cellLocations
        ' Set the value of the current cell
        Range(cellLocation(0)).Value = cellLocation(1)
    Next cellLocation
   
    ' Loop over each cell range in the array
    For Each cellLocationRange In cellLocationRanges
        ' Merge the cells in the current range
        With Range(cellLocationRange)
            .HorizontalAlignment = xlCenter ' Center the horizontal alignment
            .VerticalAlignment = xlBottom ' Align the vertical alignment to the bottom
            .WrapText = False ' Do not wrap text
            .Orientation = 0 ' Set the orientation to 0 degrees
            .AddIndent = False ' Do not add an indent
            .IndentLevel = 0 ' Set the indent level to 0
            .ShrinkToFit = False ' Do not shrink to fit
            .ReadingOrder = xlContext ' Set the reading order to context
            .MergeCells = True ' Merge the cells
            .Merge
        End With
    Next cellLocationRange
 
    ' Loop over each cell range in the array
    For Each cellRange In cellRanges
        ' Set the values of cells in the current range
        Range(cellRange).Value = cellValues
    Next cellRange

    ' Set the column width of columns A to Y to 15
    Columns("A:Y").ColumnWidth = 15

    ' Loop over each column format in the array
    For Each colFormat In colFormats
        ' Set the number format for the current column
        Columns(colFormat(0) & ":" & colFormat(0)).NumberFormat = colFormat(1)
    Next colFormat

    ' Loop over each cell range in the array
    For Each cellRange In cellRanges
        ' Merge the cells in the current cell range
        With Range(cellRange)
            .HorizontalAlignment = xlCenter ' Center the horizontal alignment
            .VerticalAlignment = xlBottom ' Align the vertical alignment to the bottom
            .WrapText = False ' Do not wrap text
            .Orientation = 0 ' Set the orientation to 0 degrees
            .AddIndent = False ' Do not add an indent
            .IndentLevel = 0 ' Set the indent level to 0
            .ShrinkToFit = False ' Do not shrink to fit
            .ReadingOrder = xlContext ' Set the reading order to context
        End With
    Next cellRange
   
    ' Select cell A3 on the active sheet
    Range("A3").Select

    ' Select the first sheet
    Sheets(FirstSheetName).Select
   
    ' Display a message box
    MsgBox "Successfully Created New Workbook"
End Sub

Sub CopyData2NewWorkbook()
    ' Variables
    Dim FirstSheetNameSTR As String
    FirstSheetNameSTR = ActiveSheet.Name
    Dim siteNames As Variant
    siteNames = Array("Site1", "Site2", "Site3", "Site4")
    Dim siteName As Variant
    Dim Error1Int As Integer
    Error1Int = 0
   
    ' Set column width and turn on AutoFilter
    Columns("D:E").ColumnWidth = 15
    Rows("1:1").AutoFilter
   
    ' Loop through each site name in the siteNames array
    For Each siteName In siteNames
        ' Filter data in columns B and F based on the current siteName
        Sheets(FirstSheetNameSTR).Range("$A$1:$AE$113510").AutoFilter Field:=2, Criteria1:=siteName
       
        ' If the current siteName is "Site1", filter column F for "BRTQC"
        If siteName = "Site1" Then ' Change as necessary
            Sheets(FirstSheetNameSTR).Range("$A$1:$AE$113510").AutoFilter Field:=6, Criteria1:="BRTQC"
        Else ' Otherwise, filter column F for "SRTQC"
            Sheets(FirstSheetNameSTR).Range("$A$1:$AE$113510").AutoFilter Field:=6, Criteria1:="SRTQC"
        End If
       
        ' Check if data has already been copied to the sheet with the same name as the current siteName
        If Sheets(siteName).Range("A2").Value = "" Then
            ' Copy filtered data from the active sheet
            Range("A1:AF1").CurrentRegion.Copy
           
            ' Paste data into sheet with the same name as the current siteName
            Sheets(siteName).Paste

            ' Turn on AutoFilter and set column width in the sheet with the same name as the current siteName
            Sheets(siteName).Rows("1:1").AutoFilter
            Sheets(siteName).Columns("D:E").ColumnWidth = 15
       
            ' Add text "End" to a cell in column E in the sheet with the same name as the current siteName
            Sheets(siteName).Range("B2").End(xlDown).Offset(1, 1).Value = "End"
        Else ' If data has already been copied to the sheet with the same name as the current siteName, set Error1Int to 1
            Error1Int = 1
            Sheets(siteName).Range("B2").End(xlDown).Offset(1, 1).Value = "End"
        End If
    Next siteName
   
    ' Turn off AutoFilter in active sheet for columns B and F
    Sheets(FirstSheetNameSTR).Range("$A$1:$AE$113510").AutoFilter Field:=2
    Sheets(FirstSheetNameSTR).Range("$A$1:$AE$113510").AutoFilter Field:=6
   
    ' Check value of Error1Int and display appropriate message box to user
    If Error1Int = 1 Then
        MsgBox "Theres already data copied into the New worksheet therefore nothing was copied over; check to make sure you didnt run this sub twice."
    Else
        MsgBox "Succesfully Copied Data to New Worksheet"
    End If
End Sub

Sub OutageCheck()
    ' Turn off screen updating to speed up the code
    Application.ScreenUpdating = False
   
    ' Variables
    Dim FirstSheetNameSTR As String
    Dim N As Integer
    Dim TestSTR As String
    Dim Time1SGL As Single
    Dim Time2SGL As Single
    Dim Time3SGL As Single
    Dim TimeDeltaSGL As Single
    Dim StartDateSGL As Single
    Dim EndDateSGL As Single
    Dim cellRef As String
    ' Variables to store the current row and column
    Dim currentRow As Long
    Dim currentColumn As Long
   
    ' Loop through the values of N from 1 to 4
    For N = 1 To 4 ' Adjust as needed
        ' Use a Select Case statement to assign values to FirstSheetNameSTR and cellRef based on the value of N
        Select Case N
            Case 1
                FirstSheetNameSTR = "Site1"
                cellRef = "A3"
            Case 2
                FirstSheetNameSTR = "Site2"
                cellRef = "F3"
            Case 3
                FirstSheetNameSTR = "Site3"
                cellRef = "K3"
            Case 4
                FirstSheetNameSTR = "Site4"
                cellRef = "P3"
        End Select
       
        'Initialize currentRow and currentColumn with the starting row and column (D2)
        currentRow = 2 'Replace 2 with the starting row as needed
        currentColumn = 4 'Replace 4 with the starting column as needed
       
        ' Begin loop
        Do While strData <> "End"
       
            ' Direct references to cells
            TestSTR = Sheets(FirstSheetNameSTR).Cells(currentRow, currentColumn - 1).Value
            If TestSTR = "End" Then
                GoTo MyQuit
            End If
        
            ' Direct references to cells
            Time1SGL = Sheets(FirstSheetNameSTR).Cells(currentRow, currentColumn).Value 'Cell D2
            Time2SGL = Sheets(FirstSheetNameSTR).Cells(currentRow + 1, currentColumn).Value 'Cell D3
            TimeDeltaSGL = Time2SGL - Time1SGL
        
            If TimeDeltaSGL < 0 Then
                Time3SGL = Time2SGL + 1
                TimeDeltaSGL = Time3SGL - Time1SGL
            End If

            ' 1.47731481481484E-04
            If (TimeDeltaSGL >= 0.0006 And TimeDeltaSGL > 0) Then
                ' Direct references to cells
                EndDateSGL = Sheets(FirstSheetNameSTR).Cells(currentRow, currentColumn + 19).Value
                StartDateSGL = Sheets(FirstSheetNameSTR).Cells(currentRow + 1, currentColumn + 19).Value
            
                With Sheets("Results")
                    .Range(cellRef).Value = StartDateSGL
                    .Range(cellRef).Offset(0, 1).Value = Time1SGL
                    .Range(cellRef).Offset(0, 2).Value = EndDateSGL
                    .Range(cellRef).Offset(0, 3).Value = Time2SGL
                    .Range(cellRef).Offset(0, 4).Value = TimeDeltaSGL
                
                    ' Move down one row in the "Results" sheet for next iteration of loop.
                    cellRef = .Range(cellRef).Offset(1, 0).Address(False, False)
                End With
            End If

            ' Direct reference to a cell
            Sheets(FirstSheetNameSTR).Cells(currentRow, currentColumn + 1).Value = TimeDeltaSGL
            ' Update currentRow and currentColumn as needed
            currentRow = currentRow + 1 ' Move right one column
        Loop
        MyQuit:
            MsgBox ("Finished Checking: " & FirstSheetNameSTR)
    Next N

    Sheets("Results").Select
    Application.ScreenUpdating = True
    MsgBox "Succesfully Completed Outage Check"
End Sub
