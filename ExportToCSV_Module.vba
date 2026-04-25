Attribute VB_Name = "ExportToCSV_Module"
Option Explicit

Sub ExportFilteredTableToCSV()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim i As Long, j As Long
    Dim filePath As String
    Dim outputFolder As String
    Dim rowData As String
    Dim totalRows As Long
    Dim processedRows As Long
    Dim exportCount As Long
    Dim fileNum As Long
    Const MAX_ROWS As Long = 100
   
    On Error GoTo ErrorHandler
   
    Set ws = ActiveSheet
   
    ' Try to get table from active cell, otherwise use first table on sheet
    On Error Resume Next
    Set tbl = ActiveCell.ListObject
    On Error GoTo 0
   
    If tbl Is Nothing Then
        If ws.ListObjects.Count = 0 Then
            MsgBox "No table found on the active sheet.", vbExclamation
            Exit Sub
        End If
        Set tbl = ws.ListObjects(1)
    End If
   
    ' ================== Folder Picker ==================
    Dim fDialog As FileDialog
    Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
    With fDialog
        .Title = "Select Folder to Save QuickBooks Export"
        .AllowMultiSelect = False
        If .Show = -1 Then
            outputFolder = .SelectedItems(1)
        Else
            MsgBox "No folder selected. Process cancelled.", vbExclamation
            Exit Sub
        End If
    End With
   
    ' ================== Build Header ==================
    Dim headerRow As String
    headerRow = ""
    For j = 1 To tbl.HeaderRowRange.Columns.Count
        headerRow = headerRow & FormatCSVCell(tbl.HeaderRowRange.Cells(1, j).Value)
        If j < tbl.HeaderRowRange.Columns.Count Then headerRow = headerRow & ","
    Next j
   
    ' Count visible rows
    Set rng = tbl.DataBodyRange
    If Not rng Is Nothing Then
        totalRows = 0
        For i = 1 To rng.Rows.Count
            If Not rng.Rows(i).Hidden Then totalRows = totalRows + 1
        Next i
    End If
   
    If totalRows = 0 Then
        MsgBox "No visible data found to export.", vbExclamation
        Exit Sub
    End If
   
    ' Filename with timestamp (prevents constant overwrite prompts)
    filePath = outputFolder & "\quickbooks_export_" & Format(Now, "yyyymmdd_hhmmss") & ".csv"
   
    If Dir(filePath) <> "" Then
        If MsgBox("File already exists:" & vbCrLf & filePath & vbCrLf & vbCrLf & "Overwrite?", vbYesNo + vbQuestion) = vbNo Then
            Exit Sub
        End If
    End If
   
    ' ================== Write CSV ==================
    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, headerRow
   
    processedRows = 0
    exportCount = 0
    Application.StatusBar = "Processing table data..."
   
    If Not rng Is Nothing Then
        For i = 1 To rng.Rows.Count
            If Not rng.Rows(i).Hidden Then
                processedRows = processedRows + 1
               
                If processedRows Mod 10 = 0 Then
                    Application.StatusBar = "Exporting row " & exportCount & " of " & Application.Min(MAX_ROWS, totalRows) & "..."
                    DoEvents
                End If
               
                rowData = ""
                For j = 1 To rng.Columns.Count
                    Dim cellVal As Variant
                    cellVal = rng.Cells(i, j).Value
                    
                    ' Improved date handling
                    If Not IsEmpty(cellVal) And IsDate(cellVal) Then
                        rowData = rowData & FormatCSVCell(Format(cellVal, "mm/dd/yyyy"))
                    Else
                        rowData = rowData & FormatCSVCell(cellVal)
                    End If
                    
                    If j < rng.Columns.Count Then rowData = rowData & ","
                Next j
               
                Print #fileNum, rowData
                exportCount = exportCount + 1
                
                If exportCount >= MAX_ROWS Then Exit For
            End If
        Next i
    End If
   
    Close #fileNum
    Application.StatusBar = False
   
    ' ================== Results ==================
    Dim resultMsg As String
    resultMsg = "Export Complete!" & vbCrLf & vbCrLf & _
                "Rows exported: " & exportCount
   
    If exportCount >= MAX_ROWS And totalRows > MAX_ROWS Then
        resultMsg = resultMsg & vbCrLf & vbCrLf & _
                    "(Note: Reached QuickBooks limit of 100 rows. Additional visible rows were skipped)."
    End If
   
    resultMsg = resultMsg & vbCrLf & vbCrLf & "Saved to:" & vbCrLf & filePath
   
    MsgBox resultMsg, vbInformation, "Export Complete"
   
    ' Open folder
    Shell "explorer.exe """ & outputFolder & """", vbNormalFocus
   
    Exit Sub
   
ErrorHandler:
    Application.StatusBar = False
    On Error Resume Next
    Close #fileNum
    On Error GoTo 0
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Export Error"
End Sub

Function FormatCSVCell(cellValue As Variant) As String
    Dim strValue As String
    If IsNull(cellValue) Or IsEmpty(cellValue) Then
        FormatCSVCell = ""
        Exit Function
    End If
    strValue = CStr(cellValue)
    If InStr(strValue, ",") > 0 Or InStr(strValue, """") > 0 Or InStr(strValue, vbCrLf) > 0 Or InStr(strValue, vbLf) > 0 Then
        strValue = Replace(strValue, """", """""")
        FormatCSVCell = """" & strValue & """"
    Else
        FormatCSVCell = strValue
    End If
End Function