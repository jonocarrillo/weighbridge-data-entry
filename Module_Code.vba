' ========================================
' MODULE CODE - WEIGHMASTER CERTIFICATE
' Maps certificate fields to Excel table
' RESPECTS EXISTING FORMULAS - DOES NOT OVERWRITE CALCULATED COLUMNS
' ========================================

Option Explicit

Public GlobalWeighmaster As String

Sub ShowWeighmasterForm()
    UserForm1.Show
End Sub

' ========================================
' PAYMENT AUTO-DETERMINATION
' Replicates the PAYMENT column LET formula:
'   If AMOUNT < 0  -> EXPENSE
'   If Account in AR vendor list -> AR
'   Otherwise -> ""
' ========================================

Function DeterminePaymentMethod(acct As String, amt As Variant) As String
    DeterminePaymentMethod = ""
    
    ' AR vendor list (matches Excel formula exactly, case-insensitive)
    Dim arVendors As Variant
    arVendors = Array("ATHENS", "CR&R", "EDCO", "EDCO RECYCLING", "EDCO DISPOSAL", _
                      "MIDWEST ROOFING", "VALLEY VISTA SERVICES", "COVANTA", _
                      "GENESIS TRUCKING", "INTERIOR DEMOLITION INC.", "PINNACLE DISPOSAL")
    
    ' Check for negative amount -> EXPENSE
    If IsNumeric(amt) And amt <> "" Then
        If CDbl(amt) < 0 Then
            DeterminePaymentMethod = "EXPENSE"
            Exit Function
        End If
    End If
    
    ' Check account against AR vendor list (UPPER + TRIM to match formula)
    Dim acctUp As String
    acctUp = UCase(Trim(acct))
    Dim i As Long
    For i = LBound(arVendors) To UBound(arVendors)
        If acctUp = UCase(arVendors(i)) Then
            DeterminePaymentMethod = "AR"
            Exit Function
        End If
    Next i
End Function

' ========================================
' DIAGNOSTIC — run once from the VBA Immediate window:  DumpFormulas
' Prints every calculated-column formula to the Immediate pane so they
' can be ported into VBA. After porting, the formulas can be deleted
' from the table entirely and VBA will write the values directly.
' ========================================
Public Sub DumpFormulas()
    Dim ws As Worksheet, tbl As ListObject
    If Not TryGetTicketTable(ws, tbl) Then
        Debug.Print "ERROR: ticket table not found"
        Exit Sub
    End If

    Debug.Print "=== Formulas in '" & tbl.Name & "' on '" & ws.Name & "' ==="
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        Dim r As Long
        For r = 1 To col.DataBodyRange.Rows.count
            Dim cell As Range
            Set cell = col.DataBodyRange.Cells(r, 1)
            If cell.HasFormula Then
                Debug.Print "[" & col.Name & "]  row " & r & ":  " & cell.Formula
                Exit For
            End If
        Next r
    Next col
    Debug.Print "=== end ==="
End Sub

' ========================================
' FORMULA INHERITANCE
' Columns VBA writes to directly. Formula inheritance MUST skip these,
' otherwise the inherited formula would overwrite the value VBA wrote.
' Keep in sync with the clearCols list inside SaveWeighmasterCertificate.
' ========================================
Public Function VBAManagedColumns() As Variant
    VBAManagedColumns = Array("Invoice Number", "WASTE DATA", "TIME IN", "TIME OUT", "Memo", _
                              "T1", "T2", "TRUCKID", "PO# / ORIGIN", "Origin", "TONAGE", "DATE", _
                              "Account", "Weighmaster", "Product/Service", "RATE", "PAYMENT", "AMOUNT")
End Function

' Pulls formulas from the most recent row above `targetRow` that contains
' a formula in each column, into `targetRow`. Skips any column listed in
' skipCols (case-insensitive) so VBA-written values are preserved.
'
' Use this whenever you want a row to "look like" the rows above it without
' having to recalculate every single row in the workbook.
Public Sub InheritFormulasFromAbove(tbl As ListObject, targetRow As ListRow, skipCols As Variant)
    If tbl Is Nothing Or targetRow Is Nothing Then Exit Sub
    If tbl.ListRows.count <= 1 Then Exit Sub

    Dim skipDict As Object
    Set skipDict = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = LBound(skipCols) To UBound(skipCols)
        skipDict(LCase(Trim(CStr(skipCols(i))))) = True
    Next i

    Dim targetIdx As Long
    targetIdx = targetRow.Index

    Dim col As Long
    For col = 1 To tbl.ListColumns.count
        Dim colName As String
        colName = LCase(Trim(tbl.ListColumns(col).Name))
        If Not skipDict.Exists(colName) Then
            ' Walk upward from the row above until we find one with a formula
            Dim r As Long
            For r = targetIdx - 1 To 1 Step -1
                Dim srcCell As Range
                Set srcCell = tbl.ListRows(r).Range.Cells(1, col)
                If srcCell.HasFormula Then
                    On Error Resume Next
                    targetRow.Range.Cells(1, col).Formula = srcCell.Formula
                    On Error GoTo 0
                    Exit For
                End If
            Next r
        End If
    Next col
End Sub

' Public helper: pull formulas into an arbitrary row by 1-based table row index.
' Run this from the Immediate window or a button:  PullFormulasIntoRow 42
Public Sub PullFormulasIntoRow(rowIndex As Long)
    Dim ws As Worksheet, tbl As ListObject
    If Not TryGetTicketTable(ws, tbl) Then
        MsgBox "Could not find the live ticket table.", vbCritical
        Exit Sub
    End If
    If rowIndex < 1 Or rowIndex > tbl.ListRows.count Then
        MsgBox "Row " & rowIndex & " is outside the table (1.." & tbl.ListRows.count & ").", vbExclamation
        Exit Sub
    End If
    InheritFormulasFromAbove tbl, tbl.ListRows(rowIndex), VBAManagedColumns()
End Sub

Private Function TableHasColumn(ByVal tbl As ListObject, ByVal colName As String) As Boolean
    If tbl Is Nothing Then Exit Function
    
    On Error Resume Next
    TableHasColumn = Not tbl.ListColumns(colName) Is Nothing
    On Error GoTo 0
End Function

Private Function IsTicketDataTable(ByVal tbl As ListObject) As Boolean
    If tbl Is Nothing Then Exit Function
    
    IsTicketDataTable = TableHasColumn(tbl, "Weighmaster") And _
                        TableHasColumn(tbl, "Account") And _
                        TableHasColumn(tbl, "Product/Service")
End Function

Private Function FindTicketTableOnSheet(ByVal ws As Worksheet) As ListObject
    Dim tbl As ListObject
    
    If ws Is Nothing Then Exit Function
    
    On Error Resume Next
    Set tbl = ws.ListObjects("Supercluster")
    On Error GoTo 0
    If IsTicketDataTable(tbl) Then
        Set FindTicketTableOnSheet = tbl
        Exit Function
    End If
    
    For Each tbl In ws.ListObjects
        If IsTicketDataTable(tbl) Then
            Set FindTicketTableOnSheet = tbl
            Exit Function
        End If
    Next tbl
End Function

Public Function TryGetTicketTable(ByRef ws As Worksheet, ByRef tbl As ListObject) As Boolean
    Dim preferredSheets As Variant
    Dim i As Long
    Dim wsLoop As Worksheet
    
    preferredSheets = Array("Supercluster", "Sheet1")
    Set ws = Nothing
    Set tbl = Nothing
    
    ' First preference: the exact live ticket table name, regardless of which sheet it sits on.
    For Each wsLoop In ThisWorkbook.Worksheets
        On Error Resume Next
        Set tbl = wsLoop.ListObjects("Supercluster")
        On Error GoTo 0
        
        If IsTicketDataTable(tbl) Then
            Set ws = wsLoop
            TryGetTicketTable = True
            Exit Function
        End If
        
        Set tbl = Nothing
    Next wsLoop
    
    For i = LBound(preferredSheets) To UBound(preferredSheets)
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(preferredSheets(i)))
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            Set tbl = FindTicketTableOnSheet(ws)
            If Not tbl Is Nothing Then
                TryGetTicketTable = True
                Exit Function
            End If
        End If
        
        Set ws = Nothing
    Next i
    
    For Each wsLoop In ThisWorkbook.Worksheets
        Set tbl = FindTicketTableOnSheet(wsLoop)
        If Not tbl Is Nothing Then
            Set ws = wsLoop
            TryGetTicketTable = True
            Exit Function
        End If
    Next wsLoop
End Function



Sub SaveWeighmasterCertificate( _
        ticketNumber As String, weighedAt As String, _
        weighedForSeller As String, companyNameHauler As String, _
        wasteOrigin As String, productCode As String, _
        truckNumber As String, ticketNumber2 As String, _
        specialNotes As String, paymentMethod As String, _
        paidAmount As Variant, date1 As Variant, _
        time1 As Variant, date2 As Variant, _
        time2 As Variant, truckID1 As String, _
        truckID2 As String, outboundWeight As Variant, _
        inboundWeight As Variant, netWeightLbs As Variant, _
        netWeightTons As Variant, weighmaster As String, _
        Optional wasteData As String = "", Optional wasteCity As String = "", _
        Optional invoiceNumber As String = "", Optional byField As String = "", _
        Optional keepLocked As Boolean = False)
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim saveStep As String   ' tracks progress so ErrorHandler can report exactly where it failed
    saveStep = "initialising"
    
    ' === PERFORMANCE: Disable screen updates and auto-calc ===
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    saveStep = "looking up the live ticket worksheet/table"
    If Not TryGetTicketTable(ws, tbl) Then
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        MsgBox "Error: Could not find the live ticket table. Expected the table named 'Supercluster' with the Weighmaster columns.", vbCritical
        Exit Sub
    End If
    
    On Error GoTo ErrorHandler
    
    ' Find Weighmaster column
    Dim wmCol As ListColumn
    
    saveStep = "looking up 'Weighmaster' column"
    On Error Resume Next
    Set wmCol = tbl.ListColumns("Weighmaster")
    On Error GoTo ErrorHandler
    
    If wmCol Is Nothing Then
        Set newRow = tbl.ListRows.Add
    Else
        ' FIRST: Check if this ticket number already exists (edit mode)
        Dim existingRow As Long
        existingRow = 0
        
        If tbl.ListRows.count > 0 Then
            Dim R As Long
            Dim cellVal As Variant
            For R = 1 To wmCol.DataBodyRange.Rows.count
                cellVal = Trim(CStr(wmCol.DataBodyRange.Cells(R, 1).Value2))
                If cellVal = Trim(ticketNumber) Then
                    existingRow = R
                    Exit For
                End If
            Next R
        End If
        
        If existingRow > 0 Then
            ' Ticket exists — update that row
            Set newRow = tbl.ListRows(existingRow)
        Else
            ' Ticket is new — always append a fresh row so we never
            ' accidentally overwrite an existing row that happens to
            ' sit after the last filled Weighmaster cell.
            Set newRow = tbl.ListRows.Add
            
            ' Copy number formats from the row above so the new row inherits
            ' cell formatting (dates, currency, etc.) without carrying over any
            ' of the previous ticket's data values.
            If tbl.ListRows.count > 1 Then
                Dim fmtRow As ListRow
                Set fmtRow = tbl.ListRows(tbl.ListRows.count - 1)
                fmtRow.Range.Copy
                newRow.Range.PasteSpecial Paste:=xlPasteFormats
                Application.CutCopyMode = False
            End If

            ' Pull formulas from the row(s) above into this new row so calculated
            ' columns (PAYMENT, CASH, AR, CARD, Quantity, Week, Biweekly, Class,
            ' Month, etc.) work immediately without having to recalculate the
            ' whole table. VBA-managed columns are skipped so values aren't
            ' overwritten.
            InheritFormulasFromAbove tbl, newRow, VBAManagedColumns()
        End If
    End If

    saveStep = "locating / creating target row for ticket #" & ticketNumber
    ' VERIFICATION POPUP: Uncomment to confirm save location
    ' MsgBox "Saving Ticket #" & ticketNumber & " to Row " & newRow.Index & " of Table '" & tbl.Name & "' on Sheet '" & ws.Name & "'", vbInformation
    
    ' ==================================================
    ' FAST SAVING: Pre-cache column indices, write all at once
    ' ==================================================
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    
    ' Build column index map ONCE (instead of looking up each time)
    Dim lc As ListColumn
    For Each lc In tbl.ListColumns
        colMap(lc.Name) = lc.Index
    Next lc
    
    ' ==================================================
    ' CLEAR ALL FORM-MANAGED COLUMNS FIRST
    ' Only columns VBA directly manages (verified against actual Table1 schema).
    ' Formula columns (AMOUNT, PAYMENT, CASH, AR, CARD, Quantity, RATE,
    ' Week, Biweekly, Class, Month) are excluded — Excel recalculates them.
    ' ==================================================
    saveStep = "clearing form-managed columns before writing"
    Dim clearCols As Variant
    clearCols = Array("Invoice Number", "WASTE DATA", "TIME IN", "TIME OUT", "Memo", _
                      "T1", "T2", "TRUCKID", "PO# / ORIGIN", "Origin", "TONAGE", "DATE", _
                      "Account", "Weighmaster", "Product/Service", "RATE")
    Dim ci As Long
    For ci = LBound(clearCols) To UBound(clearCols)
        If colMap.Exists(clearCols(ci)) Then
            newRow.Range(1, colMap(clearCols(ci))).Value = ""
        End If
    Next ci

    saveStep = "writing Weighmaster / ticket number"
    If colMap.Exists("Weighmaster") Then newRow.Range(1, colMap("Weighmaster")).Value = ticketNumber
    
    saveStep = "writing DATE"
    If colMap.Exists("DATE") Then
        If Len(Trim(CStr(date1))) > 0 And IsDate(date1) Then
            newRow.Range(1, colMap("DATE")).Value = CDate(date1)
        Else
            newRow.Range(1, colMap("DATE")).Value = ""
        End If
    End If
    
    saveStep = "writing PO/Origin (street address) and Origin (city)"
    ' PO# / ORIGIN (col I) = street address only
    ' Origin (col O)       = city only (from cmbCity dropdown)
    If colMap.Exists("PO# / ORIGIN") Then newRow.Range(1, colMap("PO# / ORIGIN")).Value = wasteOrigin
    If colMap.Exists("Origin") Then newRow.Range(1, colMap("Origin")).Value = wasteCity
    
    saveStep = "building TIME IN value"
    Dim fullTime1 As Variant
    fullTime1 = ""
    If Len(Trim(CStr(time1))) > 0 Then
        If Len(Trim(CStr(date1))) > 0 And IsDate(CStr(date1)) And IsDate(CStr(time1)) Then
            fullTime1 = CDate(CStr(date1)) + CDate(CStr(time1))
        ElseIf IsDate(CStr(time1)) Then
            fullTime1 = CDate(CStr(time1))
        Else
            fullTime1 = time1
        End If
    ElseIf Len(Trim(CStr(date1))) > 0 And IsDate(CStr(date1)) Then
        fullTime1 = CDate(CStr(date1))
    End If
    saveStep = "writing TIME IN"
    If colMap.Exists("TIME IN") Then
        If fullTime1 <> "" Then
            newRow.Range(1, colMap("TIME IN")).Value = fullTime1
            newRow.Range(1, colMap("TIME IN")).NumberFormat = "m/d/yyyy h:mm AM/PM"
        Else
            newRow.Range(1, colMap("TIME IN")).Value = ""
            newRow.Range(1, colMap("TIME IN")).NumberFormat = "General"
        End If
    End If
    
    saveStep = "building TIME OUT value"
    Dim fullTime2 As Variant
    fullTime2 = ""
    If Len(Trim(CStr(time2))) > 0 Then
        If Len(Trim(CStr(date2))) > 0 And IsDate(CStr(date2)) And IsDate(CStr(time2)) Then
            fullTime2 = CDate(CStr(date2)) + CDate(CStr(time2))
        ElseIf Len(Trim(CStr(date1))) > 0 And IsDate(CStr(date1)) And IsDate(CStr(time2)) Then
            fullTime2 = CDate(CStr(date1)) + CDate(CStr(time2))
        ElseIf IsDate(CStr(time2)) Then
            fullTime2 = CDate(CStr(time2))
        Else
            fullTime2 = time2
        End If
    End If
    
    saveStep = "writing TIME OUT"
    If colMap.Exists("TIME OUT") Then
        If fullTime2 <> "" Then
            newRow.Range(1, colMap("TIME OUT")).Value = fullTime2
            newRow.Range(1, colMap("TIME OUT")).NumberFormat = "m/d/yyyy h:mm AM/PM"
        Else
            newRow.Range(1, colMap("TIME OUT")).Value = ""
            newRow.Range(1, colMap("TIME OUT")).NumberFormat = "General"
        End If
    End If
    
    saveStep = "writing Memo"
    ' Always write — even blank — so empty Memo overwrites any seeded value
    If colMap.Exists("Memo") Then newRow.Range(1, colMap("Memo")).Value = weighedForSeller
    
    saveStep = "writing T1 (Inbound Weight)"
    If colMap.Exists("T1") Then
        If IsNumeric(inboundWeight) And Len(Trim(CStr(inboundWeight))) > 0 And CDbl(inboundWeight) <> 0 Then
            newRow.Range(1, colMap("T1")).Value = CDbl(inboundWeight)
        Else
            newRow.Range(1, colMap("T1")).Value = ""
        End If
    End If
    
    saveStep = "writing TRUCKID"
    ' Always write — even blank
    If colMap.Exists("TRUCKID") Then newRow.Range(1, colMap("TRUCKID")).Value = truckNumber
    
    saveStep = "writing Account"
    If colMap.Exists("Account") Then
        If Len(Trim(companyNameHauler)) > 0 Then
            newRow.Range(1, colMap("Account")).Value = companyNameHauler
        ElseIf Len(Trim(weighedForSeller)) > 0 Then
            newRow.Range(1, colMap("Account")).Value = weighedForSeller
        Else
            newRow.Range(1, colMap("Account")).Value = ""
        End If
    End If
    
    saveStep = "writing Product/Service"
    ' Always write — even blank
    If colMap.Exists("Product/Service") Then newRow.Range(1, colMap("Product/Service")).Value = productCode
    
    saveStep = "writing WASTE DATA"
    ' Always write — even blank
    If colMap.Exists("WASTE DATA") Then newRow.Range(1, colMap("WASTE DATA")).Value = wasteData
    
    ' Normalize missing outbound/tare values so incomplete tickets still save
    ' cleanly instead of blanking dependent columns.
    Dim hasInboundWeight As Boolean
    Dim hasOutboundWeight As Boolean
    Dim normalizedInbound As Double
    Dim normalizedOutbound As Double
    Dim normalizedTons As Double
    Dim normalizedAmount As Variant
    Dim normalizedPayment As String
    
    hasInboundWeight = IsNumeric(inboundWeight) And Len(Trim(CStr(inboundWeight))) > 0
    hasOutboundWeight = IsNumeric(outboundWeight) And Len(Trim(CStr(outboundWeight))) > 0
    
    If hasInboundWeight Then normalizedInbound = CDbl(inboundWeight)
    If hasOutboundWeight Then
        normalizedOutbound = CDbl(outboundWeight)
    Else
        normalizedOutbound = 0
    End If
    
    ' If outbound (T2/tare) is missing, leave T2/TONAGE/AMOUNT/RATE blank
    ' rather than fabricating values from inbound alone — fabricated values
    ' propagate into dependent rate/amount formulas and trigger #REF errors.
    saveStep = "writing T2 (Outbound Weight)"
    If colMap.Exists("T2") Then
        If hasOutboundWeight Then
            newRow.Range(1, colMap("T2")).Value = normalizedOutbound
        Else
            newRow.Range(1, colMap("T2")).Value = ""
        End If
    End If

    saveStep = "writing TONAGE"
    Dim tonStr As String
    tonStr = Trim(Replace(Replace(Replace(CStr(netWeightTons), ",", ""), "$", ""), " ", ""))
    normalizedTons = 0
    If IsNumeric(tonStr) And CDbl(tonStr) > 0 Then
        normalizedTons = CDbl(tonStr)
    ElseIf hasInboundWeight And hasOutboundWeight Then
        normalizedTons = Abs(normalizedInbound - normalizedOutbound) / 2000#
        If normalizedTons > 0 And normalizedTons < 1 Then normalizedTons = 1
    End If

    If colMap.Exists("TONAGE") Then
        If normalizedTons > 0 And hasOutboundWeight Then
            newRow.Range(1, colMap("TONAGE")).Value = normalizedTons
        Else
            newRow.Range(1, colMap("TONAGE")).Value = ""
        End If
    End If

    saveStep = "writing AMOUNT"
    Dim amtStr As String
    ' Strip $, commas, spaces — IsNumeric() returns True for "$150" but CDbl("$150") = Error #13
    amtStr = Trim(Replace(Replace(Replace(CStr(paidAmount), ",", ""), "$", ""), " ", ""))
    normalizedAmount = ""
    If IsNumeric(amtStr) And CDbl(amtStr) <> 0 Then
        normalizedAmount = CDbl(amtStr)
    ElseIf normalizedTons > 0 And hasOutboundWeight Then
        normalizedAmount = CalculateAmountFromPricing(productCode, companyNameHauler, date1, normalizedTons)
        If IsNumeric(normalizedAmount) Then
            If UCase(Trim(paymentMethod)) = "CREDIT" Then
                normalizedAmount = CDbl(normalizedAmount) * 1.023
            Else
                normalizedAmount = CDbl(normalizedAmount)
            End If
        Else
            normalizedAmount = ""
        End If
    End If
    If colMap.Exists("AMOUNT") Then
        If IsNumeric(normalizedAmount) And CDbl(normalizedAmount) <> 0 Then
            newRow.Range(1, colMap("AMOUNT")).Value = CDbl(normalizedAmount)
        Else
            newRow.Range(1, colMap("AMOUNT")).Value = ""
        End If
    End If
    
    saveStep = "writing RATE"
    If colMap.Exists("RATE") Then
        If LCase(Trim(productCode)) = "weight" Then
            If IsNumeric(normalizedAmount) And CDbl(normalizedAmount) <> 0 Then
                newRow.Range(1, colMap("RATE")).Value = CDbl(normalizedAmount)
            Else
                newRow.Range(1, colMap("RATE")).Value = ""
            End If
        Else
            If IsNumeric(normalizedAmount) And CDbl(normalizedAmount) <> 0 And normalizedTons > 0 Then
                newRow.Range(1, colMap("RATE")).Value = Round(CDbl(normalizedAmount) / normalizedTons, 2)
            Else
                newRow.Range(1, colMap("RATE")).Value = ""
            End If
        End If
    End If
    
    saveStep = "writing PAYMENT"
    ' Always write — even blank — so clearing payment type is saved correctly
    normalizedPayment = paymentMethod
    If Len(Trim(normalizedPayment)) = 0 And IsNumeric(normalizedAmount) Then
        normalizedPayment = DeterminePaymentMethod(companyNameHauler, normalizedAmount)
    End If
    If colMap.Exists("PAYMENT") Then newRow.Range(1, colMap("PAYMENT")).Value = normalizedPayment
    
    saveStep = "writing Invoice Number"
    ' Always write — even blank — so invoice number is cleared when a new blank ticket is saved
    If colMap.Exists("Invoice Number") Then newRow.Range(1, colMap("Invoice Number")).Value = invoiceNumber
    
    ' NOTE: txtWeighmaster (operator/person name) and txtBy have no dedicated writable columns
    ' in the current Table1 schema. They are used for printing only (DrawCurrentTicketToSheet).
    
    ' EnableEvents and Calculation MUST be restored before DoEvents/Calculate/FileSave
    ' so that ExecuteMso "FileSave" can complete without Error #13 (Type Mismatch).
    ' ScreenUpdating is only restored if the caller is NOT managing it (keepLocked=False).
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    If Not keepLocked Then Application.ScreenUpdating = True
    
    ' Flush all pending Excel work before saving
    DoEvents
    
    ' Force formula columns (AMOUNT etc.) to recalculate with the new row data
    Application.Calculate
    DoEvents
    
    ' Use Ctrl+S equivalent — this correctly integrates with OneDrive sync.
    ' ThisWorkbook.Save can silently conflict with AutoSave on OneDrive files.
    On Error Resume Next
    Application.CommandBars.ExecuteMso "FileSave"
    DoEvents  ' Give OneDrive a moment to acknowledge the save
    On Error GoTo 0
    
    Exit Sub

ErrorHandler:
    Dim errWsName As String
    Dim errTblName As String
    
    errWsName = "(not found)"
    errTblName = "(not found)"
    If Not ws Is Nothing Then errWsName = "'" & ws.Name & "'"
    If Not tbl Is Nothing Then errTblName = "'" & tbl.Name & "'"
    
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Error saving certificate." & vbCrLf & vbCrLf & _
           "Step: " & saveStep & vbCrLf & _
           "Error: " & Err.Description & " (Error #" & Err.Number & ")" & vbCrLf & vbCrLf & _
           "Sheet: " & errWsName & _
           "   Table: " & errTblName, _
           vbCritical, "Save Error"
End Sub

' Helper Sub to safely write to a column by NAME
Private Sub WriteToColumn(tbl As ListObject, row As ListRow, colName As String, val As Variant)
    Dim colIndex As Long
    On Error Resume Next
    colIndex = tbl.ListColumns(colName).Index
    On Error GoTo 0
    
    If colIndex > 0 Then
        If IsNumeric(val) And val <> "" Then
            ' Check if it looks like a number but might actually be text (like "10-20")
            If VarType(val) = vbString And InStr(val, "-") > 0 And Not IsDate(val) Then
                 row.Range(1, colIndex).Value = val
            Else
                 row.Range(1, colIndex).Value = CDbl(val)
            End If
        Else
            row.Range(1, colIndex).Value = val
        End If
    Else
        ' Column mismatch debug
        ' Debug.Print "Warning: Column '" & colName & "' not found in table."
    End If
End Sub

Function GetNextTicketNumber() As String
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim ticketNum As Long
    Dim col As ListColumn
    
    If Not TryGetTicketTable(ws, tbl) Then
        GetNextTicketNumber = "103092"
        Exit Function
    End If
    
    ' Find Weighmaster column
    Set col = Nothing
    On Error Resume Next
    Set col = tbl.ListColumns("Weighmaster")
    On Error GoTo 0
    
    If col Is Nothing Then
        GetNextTicketNumber = "103092"
        Exit Function
    End If
    
    If tbl.ListRows.count > 0 Then
        ' Scan from bottom up to find the last non-empty Weighmaster entry
        Dim R As Long
        Dim cellVal As Variant
        Dim lastVal As Long
        lastVal = 0
        
        For R = col.DataBodyRange.Rows.count To 1 Step -1
            cellVal = col.DataBodyRange.Cells(R, 1).Value2
            If Not IsEmpty(cellVal) And Trim(CStr(cellVal)) <> "" Then
                lastVal = CLng(val(CStr(cellVal)))
                Exit For
            End If
        Next R
        
        If lastVal > 0 Then
            ticketNum = lastVal + 1
        Else
            ticketNum = 103092
        End If
    Else
        ticketNum = 103092
    End If
    
    GetNextTicketNumber = CStr(ticketNum)
End Function

Function GetTicketList() As Variant
    ' Returns an array of all non-empty ticket numbers from the Weighmaster column
    ' Returns Empty if no tickets found
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn
    
    If Not TryGetTicketTable(ws, tbl) Then
        GetTicketList = Empty
        Exit Function
    End If
    
    On Error Resume Next
    Set col = tbl.ListColumns("Weighmaster")
    On Error GoTo 0
    
    If col Is Nothing Or tbl.ListRows.count = 0 Then
        GetTicketList = Empty
        Exit Function
    End If
    
    ' Collect all non-empty ticket numbers
    Dim tickets() As String
    Dim cnt As Long
    Dim R As Long
    Dim cellVal As Variant
    cnt = 0
    
    ReDim tickets(1 To col.DataBodyRange.Rows.count)
    
    For R = 1 To col.DataBodyRange.Rows.count
        cellVal = Trim(CStr(col.DataBodyRange.Cells(R, 1).Value2))
        If Len(cellVal) > 0 Then
            cnt = cnt + 1
            tickets(cnt) = cellVal
        End If
    Next R
    
    If cnt = 0 Then
        GetTicketList = Empty
    Else
        ReDim Preserve tickets(1 To cnt)
        GetTicketList = tickets
    End If
End Function

Function LoadTicketFromTable(ticketNumber As String) As Object
    ' Finds an existing ticket by number and returns its data as a Dictionary
    ' Returns Nothing if not found
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn
    
    If Not TryGetTicketTable(ws, tbl) Then
        Set LoadTicketFromTable = Nothing
        Exit Function
    End If
    
    On Error Resume Next
    Set col = tbl.ListColumns("Weighmaster")
    On Error GoTo 0
    
    If col Is Nothing Or tbl.ListRows.count = 0 Then
        Set LoadTicketFromTable = Nothing
        Exit Function
    End If
    
    ' Find the row with this ticket number
    Dim R As Long
    Dim foundRow As Long
    foundRow = 0
    
    For R = 1 To col.DataBodyRange.Rows.count
        If Trim(CStr(col.DataBodyRange.Cells(R, 1).Value2)) = Trim(ticketNumber) Then
            foundRow = R
            Exit For
        End If
    Next R
    
    If foundRow = 0 Then
        Set LoadTicketFromTable = Nothing
        Exit Function
    End If
    
    ' Build column map
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    Dim lc As ListColumn
    For Each lc In tbl.ListColumns
        colMap(lc.Name) = lc.Index
    Next lc
    
    ' Read values into result dictionary
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    Dim row As ListRow
    Set row = tbl.ListRows(foundRow)
    
    If colMap.Exists("PO# / ORIGIN") Then result("WasteOrigin") = CStr(row.Range(1, colMap("PO# / ORIGIN")).Value)
    If colMap.Exists("Origin") Then result("City") = CStr(row.Range(1, colMap("Origin")).Value)
    If colMap.Exists("WASTE DATA") Then result("WasteData") = CStr(row.Range(1, colMap("WASTE DATA")).Value)
    If colMap.Exists("Account") Then result("CompanyNameHauler") = CStr(row.Range(1, colMap("Account")).Value)
    If colMap.Exists("Product/Service") Then result("ProductCode") = CStr(row.Range(1, colMap("Product/Service")).Value)
    If colMap.Exists("TRUCKID") Then result("TruckNumber") = CStr(row.Range(1, colMap("TRUCKID")).Value)
    If colMap.Exists("Memo") Then result("WeighedForSeller") = CStr(row.Range(1, colMap("Memo")).Value)
    If colMap.Exists("PAYMENT") Then result("PaymentMethod") = CStr(row.Range(1, colMap("PAYMENT")).Value)
    
    ' Read DATE column (only "DATE" uppercase exists in Table1)
    If colMap.Exists("DATE") Then
        Dim dVal As Variant
        dVal = row.Range(1, colMap("DATE")).Value
        If dVal <> "" And (IsNumeric(dVal) Or IsDate(dVal)) Then
            result("Date1") = Format(CDate(dVal), "mm/dd/yyyy")
        ElseIf dVal <> "" Then
            result("Date1") = CStr(dVal)
        End If
    End If

    If colMap.Exists("TIME IN") Then
        Dim timeIn As Variant
        timeIn = row.Range(1, colMap("TIME IN")).Value
        If timeIn <> "" And (IsNumeric(timeIn) Or IsDate(timeIn)) Then
            ' Only use TIME IN's date part if Date1 is still empty
            If Not result.Exists("Date1") Or result("Date1") = "" Then
                result("Date1") = Format(CDate(timeIn), "mm/dd/yyyy")
            End If
            result("Time1") = Format(CDate(timeIn), "hh:mm AM/PM")
        Else
            result("Time1") = CStr(timeIn)
        End If
    End If
    
    If colMap.Exists("TIME OUT") Then
        Dim timeOut As Variant
        timeOut = row.Range(1, colMap("TIME OUT")).Value
        If timeOut <> "" And (IsNumeric(timeOut) Or IsDate(timeOut)) Then
            result("Time2") = Format(CDate(timeOut), "hh:mm AM/PM")
        Else
            result("Time2") = CStr(timeOut)
        End If
    End If
    
    If colMap.Exists("T1") Then result("InboundWeight") = CStr(row.Range(1, colMap("T1")).Value)
    If colMap.Exists("T2") Then result("OutboundWeight") = CStr(row.Range(1, colMap("T2")).Value)
    If colMap.Exists("Invoice Number") Then result("InvoiceNum") = CStr(row.Range(1, colMap("Invoice Number")).Value)
    If colMap.Exists("AMOUNT") Then result("PaidAmount") = CStr(row.Range(1, colMap("AMOUNT")).Value)
    If colMap.Exists("TONAGE") Then result("Tonnage") = CStr(row.Range(1, colMap("TONAGE")).Value)
    
    Set LoadTicketFromTable = result
End Function

' ========================================
' HAUL PRODUCTS — single editable list. To add/remove a haul product:
'   1. Edit this array
'   2. Make sure the second value matches a column name in HaulingPricing
'   3. Make sure "Allowed Tons" entries are comma-separated in the SAME order
' Each entry is Array(<product code typed in form>, <HaulingPricing column name>).
' ========================================
Public Function HaulProducts() As Variant
    HaulProducts = Array( _
        Array("LOWBOY", "Lowboy"), _
        Array("10YD",   "10YD"), _
        Array("40YD",   "40YD") _
    )
End Function

Function CalculateAmountFromPricing(product As String, account As String, _
                                     ticketDate As Variant, tonnage As Double) As Variant
    ' Replicates the Excel AMOUNT LET formula using ScalePricing and HaulingPricing tables.
    ' Memo is NOT used in pricing — amounts come from table lookups only.
    ' Returns the calculated amount or "" if no match / result is zero.
    
    Dim prod As String, acct As String
    prod = UCase(Trim(product))
    acct = UCase(Trim(account))
    
    ' Parse the date
    Dim d As Date
    If Not IsDate(ticketDate) Or Len(Trim(CStr(ticketDate))) = 0 Then
        d = Date
    Else
        d = CDate(ticketDate)
    End If
    
    ' Check if hauling product — driven by HaulProducts() so the list is
    ' edited in one place. haulPos is 1-based and lines up with the order
    ' of comma-separated values in HaulingPricing.[Allowed Tons].
    Dim isHaul As Boolean
    Dim haulPos As Long
    Dim haulColName As String
    Dim hpList As Variant
    isHaul = False
    haulPos = 0
    haulColName = ""
    hpList = HaulProducts()
    Dim hpI As Long
    For hpI = LBound(hpList) To UBound(hpList)
        If UCase(Trim(CStr(hpList(hpI)(0)))) = prod Then
            isHaul = True
            haulPos = hpI - LBound(hpList) + 1
            haulColName = CStr(hpList(hpI)(1))
            Exit For
        End If
    Next hpI
    
    ' --- SCALE PRICING LOOKUP ---
    Dim chargeS As Variant
    chargeS = ""
    
    Dim spTbl As ListObject
    Set spTbl = Nothing
    On Error Resume Next
    Dim wsLoop As Worksheet
    For Each wsLoop In ThisWorkbook.Worksheets
        Set spTbl = wsLoop.ListObjects("ScalePricing")
        If Not spTbl Is Nothing Then Exit For
    Next wsLoop
    On Error GoTo 0
    
    If Not spTbl Is Nothing Then
        Dim spRows As Long
        spRows = spTbl.ListRows.count
        
        If spRows > 0 Then
            Dim spProd As Variant, spAcct As Variant, spStart As Variant, spEnd As Variant
            Dim spRate As Variant, spType As Variant
            
            ' Get column indices
            Dim ciProd As Long, ciAcct As Long, ciStart As Long, ciEnd As Long
            Dim ciRate As Long, ciType As Long
            On Error Resume Next
            ciProd = spTbl.ListColumns("PRODUCT/SERVICE").Index
            ciAcct = spTbl.ListColumns("Account").Index
            ciStart = spTbl.ListColumns("Start").Index
            ciEnd = spTbl.ListColumns("END").Index
            ciRate = spTbl.ListColumns("Rate").Index
            ciType = spTbl.ListColumns("Type").Index
            On Error GoTo 0
            
            If ciProd > 0 And ciRate > 0 Then
                Dim idxExact As Long, idxPublic As Long
                idxExact = 0: idxPublic = 0
                
                Dim R As Long
                For R = 1 To spRows
                    Dim rowProd As String, rowAcct As String
                    Dim rowStart As Date, rowEnd As Date
                    
                    rowProd = UCase(Trim(CStr(spTbl.DataBodyRange.Cells(R, ciProd).Value)))
                    If rowProd <> prod Then GoTo NextSP
                    
                    ' Date normalization — LET: IF(YEAR(Start+0)>=2020, Start, DATE(YEAR(_d),MONTH(Start),DAY(Start)))
                    ' Dates stored with a real year (>=2020) are used as-is.
                    ' Template-style dates (e.g. "Jan 1" with no real year) are re-anchored to the ticket year.
                    If ciStart > 0 And ciEnd > 0 Then
                        On Error Resume Next
                        Dim spRawStart As Variant, spRawEnd As Variant
                        spRawStart = spTbl.DataBodyRange.Cells(R, ciStart).Value
                        spRawEnd   = spTbl.DataBodyRange.Cells(R, ciEnd).Value
                        If IsDate(spRawStart) Then
                            If Year(CDate(spRawStart)) >= 2020 Then
                                rowStart = CDate(spRawStart)
                            Else
                                rowStart = DateSerial(Year(d), Month(CDate(spRawStart)), Day(CDate(spRawStart)))
                            End If
                        End If
                        If IsDate(spRawEnd) Then
                            If Year(CDate(spRawEnd)) >= 2020 Then
                                rowEnd = CDate(spRawEnd)
                            Else
                                rowEnd = DateSerial(Year(d), Month(CDate(spRawEnd)), Day(CDate(spRawEnd)))
                            End If
                        End If
                        On Error GoTo 0
                        If d < rowStart Or d > rowEnd Then GoTo NextSP
                    End If
                    
                    ' Check account
                    If ciAcct > 0 Then
                        rowAcct = UCase(Trim(CStr(spTbl.DataBodyRange.Cells(R, ciAcct).Value)))
                    Else
                        rowAcct = ""
                    End If
                    
                    If rowAcct = acct And idxExact = 0 Then
                        idxExact = R
                    ElseIf (rowAcct = "" Or rowAcct = "PUBLIC") And idxPublic = 0 Then
                        idxPublic = R
                    End If
NextSP:
                Next R
                
                Dim spIdx As Long
                If idxExact > 0 Then
                    spIdx = idxExact
                ElseIf idxPublic > 0 Then
                    spIdx = idxPublic
                Else
                    spIdx = 0
                End If
                
                If spIdx > 0 Then
                    Dim rate As Double, rateType As String
                    rate = CDbl(spTbl.DataBodyRange.Cells(spIdx, ciRate).Value)
                    If ciType > 0 Then
                        rateType = LCase(Trim(CStr(spTbl.DataBodyRange.Cells(spIdx, ciType).Value)))
                    Else
                        rateType = ""
                    End If
                    
                    If rateType = "run" Then
                        chargeS = rate
                    Else
                        chargeS = tonnage * rate
                    End If
                End If
            End If
        End If
    End If
    
    ' --- HAULING PRICING LOOKUP ---
    Dim chargeH As Variant
    chargeH = ""
    
    If isHaul Then
        Dim hpTbl As ListObject
        Set hpTbl = Nothing
        On Error Resume Next
        For Each wsLoop In ThisWorkbook.Worksheets
            Set hpTbl = wsLoop.ListObjects("HaulingPricing")
            If Not hpTbl Is Nothing Then Exit For
        Next wsLoop
        On Error GoTo 0
        
        If Not hpTbl Is Nothing Then
            Dim hpRows As Long
            hpRows = hpTbl.ListRows.count
            
            If hpRows > 0 Then
                Dim hiAcct As Long, hiStart As Long, hiEnd As Long
                Dim hiAllowed As Long, hiOverage As Long, hiMiles As Long
                Dim hiDistRate As Long, hiSurcharge As Long
                Dim hiBaseRate As Long  ' resolved by haulColName from HaulProducts()

                On Error Resume Next
                hiAcct      = hpTbl.ListColumns("Account").Index
                hiStart     = hpTbl.ListColumns("Start").Index
                hiEnd       = hpTbl.ListColumns("End").Index
                hiAllowed   = hpTbl.ListColumns("Allowed Tons").Index
                hiOverage   = hpTbl.ListColumns("Overage").Index
                hiMiles     = hpTbl.ListColumns("Miles").Index
                hiDistRate  = hpTbl.ListColumns("Dist. Rate").Index
                hiSurcharge = hpTbl.ListColumns("Surcharge").Index
                If Len(haulColName) > 0 Then hiBaseRate = hpTbl.ListColumns(haulColName).Index
                On Error GoTo 0
                
                Dim hIdxExact As Long, hIdxPublic As Long
                hIdxExact = 0: hIdxPublic = 0
                
                For R = 1 To hpRows
                    ' Date normalization — same LET logic as ScalePricing
                    If hiStart > 0 And hiEnd > 0 Then
                        On Error Resume Next
                        Dim hpRawStart As Variant, hpRawEnd As Variant
                        Dim hpRowStart As Date, hpRowEnd As Date
                        hpRawStart = hpTbl.DataBodyRange.Cells(R, hiStart).Value
                        hpRawEnd   = hpTbl.DataBodyRange.Cells(R, hiEnd).Value
                        If IsDate(hpRawStart) Then
                            If Year(CDate(hpRawStart)) >= 2020 Then
                                hpRowStart = CDate(hpRawStart)
                            Else
                                hpRowStart = DateSerial(Year(d), Month(CDate(hpRawStart)), Day(CDate(hpRawStart)))
                            End If
                        End If
                        If IsDate(hpRawEnd) Then
                            If Year(CDate(hpRawEnd)) >= 2020 Then
                                hpRowEnd = CDate(hpRawEnd)
                            Else
                                hpRowEnd = DateSerial(Year(d), Month(CDate(hpRawEnd)), Day(CDate(hpRawEnd)))
                            End If
                        End If
                        On Error GoTo 0
                        If d < hpRowStart Or d > hpRowEnd Then GoTo NextHP
                    End If
                    
                    If hiAcct > 0 Then
                        rowAcct = UCase(Trim(CStr(hpTbl.DataBodyRange.Cells(R, hiAcct).Value)))
                    Else
                        rowAcct = ""
                    End If
                    
                    If rowAcct = acct And hIdxExact = 0 Then
                        hIdxExact = R
                    ElseIf (rowAcct = "" Or rowAcct = "PUBLIC") And hIdxPublic = 0 Then
                        hIdxPublic = R
                    End If
NextHP:
                Next R
                
                Dim hIdx As Long
                If hIdxExact > 0 Then
                    hIdx = hIdxExact
                ElseIf hIdxPublic > 0 Then
                    hIdx = hIdxPublic
                Else
                    hIdx = 0
                End If
                
                If hIdx > 0 Then
                    ' Base rate is read from the column resolved via HaulProducts() — no per-product Select Case.
                    Dim baseRate As Double
                    baseRate = 0
                    On Error Resume Next
                    If hiBaseRate > 0 Then
                        baseRate = CDbl(hpTbl.DataBodyRange.Cells(hIdx, hiBaseRate).Value)
                    End If
                    
                    ' Allowed Tons: comma-separated string "Lowboy,CONCRE,10YD,40YD"
                    ' e.g. "0,0,0,4" -> 40YD (haulPos=4) -> parts(3) = "4" -> allowVal = 4
                    Dim allowStr As String, allowVal As Double
                    allowVal = 0
                    If hiAllowed > 0 Then
                        allowStr = Trim(CStr(hpTbl.DataBodyRange.Cells(hIdx, hiAllowed).Value))
                        If Len(allowStr) > 0 Then
                            Dim parts() As String
                            parts = Split(allowStr, ",")
                            ' haulPos is 1-based; parts() is 0-based; so index = haulPos - 1
                            If haulPos - 1 <= UBound(parts) Then
                                allowVal = Val(Trim(parts(haulPos - 1)))
                            End If
                        End If
                    End If
                    
                    Dim overageRate As Double, miles As Double, distRate As Double, surcharge As Double
                    overageRate = 0: miles = 0: distRate = 0: surcharge = 0
                    If hiOverage > 0 Then overageRate = CDbl(hpTbl.DataBodyRange.Cells(hIdx, hiOverage).Value)
                    If hiMiles > 0 Then miles = CDbl(hpTbl.DataBodyRange.Cells(hIdx, hiMiles).Value)
                    If hiDistRate > 0 Then distRate = CDbl(hpTbl.DataBodyRange.Cells(hIdx, hiDistRate).Value)
                    If hiSurcharge > 0 Then surcharge = CDbl(hpTbl.DataBodyRange.Cells(hIdx, hiSurcharge).Value)
                    On Error GoTo 0
                    
                    Dim factor As Double
                    factor = 1 + surcharge / 100
                    
                    Dim overageCharge As Double, distCharge As Double
                    overageCharge = 0
                    If tonnage > allowVal Then overageCharge = (tonnage - allowVal) * overageRate
                    distCharge = miles * distRate
                    
                    chargeH = (baseRate + overageCharge + distCharge) * factor
                End If
            End If
        End If
    End If
    
    ' Final pick — replicates LET:
    '   _pick = IF(_isHaul, _chargeH, _chargeS)
    '   Result = IFERROR(IF(_pick=0, "", _pick), "")   [cash override handled at top]
    Dim pick As Variant
    If isHaul Then
        pick = chargeH
    Else
        pick = chargeS
    End If
    
    If IsNumeric(pick) Then
        If CDbl(pick) = 0 Then
            CalculateAmountFromPricing = ""
        Else
            CalculateAmountFromPricing = pick
        End If
    ElseIf IsEmpty(pick) Or CStr(pick) = "" Then
        CalculateAmountFromPricing = ""
    Else
        CalculateAmountFromPricing = pick
    End If
End Function

' ========================================
' TARE WEIGHT LOOKUP (replicates T2 formula)
' Checks Tonage[EDCO RECYCLING] -> Tonage[TARE]
' Fallback: Tonage[GENESIS TRUCKING] -> Tonage[TARE2]
' Returns 0 if not found
' ========================================

Function LookupTareWeight(truckID As String) As Double
    LookupTareWeight = 0
    If Len(Trim(truckID)) = 0 Then Exit Function
    
    Dim tonTbl As ListObject
    Dim wsLoop As Worksheet
    
    ' Find Tonage table across all sheets
    On Error Resume Next
    For Each wsLoop In ThisWorkbook.Worksheets
        Set tonTbl = wsLoop.ListObjects("Tonage")
        If Not tonTbl Is Nothing Then Exit For
    Next wsLoop
    On Error GoTo 0
    
    If tonTbl Is Nothing Then Exit Function
    
    Dim R As Long
    Dim colEdco As Long, colTare As Long
    Dim colGen As Long, colTare2 As Long
    
    ' Get column indices
    On Error Resume Next
    colEdco = tonTbl.ListColumns("EDCO RECYCLING").Index
    colTare = tonTbl.ListColumns("TARE").Index
    colGen = tonTbl.ListColumns("GENESIS TRUCKING").Index
    colTare2 = tonTbl.ListColumns("TARE2").Index
    On Error GoTo 0
    
    Dim target As String
    target = UCase(Trim(truckID))
    
    ' First pass: EDCO RECYCLING -> TARE
    If colEdco > 0 And colTare > 0 Then
        For R = 1 To tonTbl.ListRows.count
            If UCase(Trim(CStr(tonTbl.DataBodyRange.Cells(R, colEdco).Value))) = target Then
                Dim tareVal As Variant
                tareVal = tonTbl.DataBodyRange.Cells(R, colTare).Value
                If IsNumeric(tareVal) And tareVal <> "" Then
                    LookupTareWeight = CDbl(tareVal)
                    Exit Function
                End If
            End If
        Next R
    End If
    
    ' Second pass: GENESIS TRUCKING -> TARE2
    If colGen > 0 And colTare2 > 0 Then
        For R = 1 To tonTbl.ListRows.count
            If UCase(Trim(CStr(tonTbl.DataBodyRange.Cells(R, colGen).Value))) = target Then
                Dim tareVal2 As Variant
                tareVal2 = tonTbl.DataBodyRange.Cells(R, colTare2).Value
                If IsNumeric(tareVal2) And tareVal2 <> "" Then
                    LookupTareWeight = CDbl(tareVal2)
                    Exit Function
                End If
            End If
        Next R
    End If
End Function

' ========================================
' ONE-TIME CLEANUP (Run manually if needed)
' ========================================
Public Sub CleanupOldPrintSheets()
    Dim ws As Worksheet
    Dim count As Long
    count = 0
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Worksheets
        If Left(ws.Name, 12) = "PrintTicket_" And ws.Name <> "PrintTicket_Temp" Then
            On Error Resume Next
            ws.Visible = xlSheetVisible ' Unhide first to avoid 1004 delete errors
            ws.Delete
            If Err.Number = 0 Then count = count + 1
            On Error GoTo 0
        End If
    Next ws
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Cleanup Complete! Deleted " & count & " old print sheets.", vbInformation
End Sub

Public Sub ExportFilteredTicketsToPDF()
    ' This macro can be attached to a button on the Supercluster sheet.
    ' It delegates the actual loop and PDF generation to the UserForm where the layout logic lives.
    UserForm1.ExportVisibleTicketsToPDF
End Sub




