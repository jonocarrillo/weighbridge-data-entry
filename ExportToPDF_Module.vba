Attribute VB_Name = "ExportToPDF_Module"
Option Explicit

' HOW TO USE:
'   1. FIRST TIME ONLY: Alt+F8 -> SetupTicketTemplate  (creates the TicketTemplate sheet)
'   2. Run ExportFilteredTicketsToPDF  (Alt+F8, or assign to a button)
'
'   To edit the ticket layout:
'     Alt+F8 -> ShowTicketTemplate  (unhides the sheet for editing)
'     Alt+F8 -> HideTicketTemplate  (hides it again)
'
' Each visible ticket = 1 page, ticket fills the TOP HALF only.
' Bottom half is blank. Layout/formatting lives in the "TicketTemplate" sheet.

Public Sub ExportFilteredTicketsToPDF()

    ' 1. Locate table
    Dim ws As Worksheet, tbl As ListObject
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Supercluster")
    On Error GoTo 0
    If ws Is Nothing Then MsgBox "Worksheet 'Supercluster' not found.", vbCritical: Exit Sub
    On Error Resume Next
    Set tbl = ws.ListObjects("Supercluster")
    If tbl Is Nothing Then Set tbl = ws.ListObjects("Table1")
    On Error GoTo 0
    If tbl Is Nothing Then MsgBox "Table not found.", vbCritical: Exit Sub

    ' 2. Verify template exists
    Dim wsTpl As Worksheet
    On Error Resume Next
    Set wsTpl = ThisWorkbook.Worksheets(TEMPLATE_SHEET_NAME)
    On Error GoTo 0
    If wsTpl Is Nothing Then
        Dim ans As Integer
        ans = MsgBox("TicketTemplate sheet not found." & vbCrLf & _
                     "Do you want to create it now?", vbQuestion + vbYesNo)
        If ans = vbYes Then
            SetupTicketTemplate
            Set wsTpl = ThisWorkbook.Worksheets(TEMPLATE_SHEET_NAME)
        Else
            Exit Sub
        End If
    End If

    ' 3. Save-as dialog
    Dim pdfPath As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    With fd
        .Title = "Save Filtered Tickets PDF As"
        .InitialFileName = Environ("USERPROFILE") & "\Desktop\Tickets_" & Format(Date, "mm-dd-yyyy") & ".pdf"
        .FilterIndex = 2
        If .Show = -1 Then
            pdfPath = .SelectedItems(1)
            If LCase(Right(pdfPath, 4)) <> ".pdf" Then pdfPath = pdfPath & ".pdf"
        Else
            Exit Sub
        End If
    End With

    ' 4. Column map
    Dim colMap As Object
    Set colMap = CreateObject("Scripting.Dictionary")
    Dim lc As ListColumn
    For Each lc In tbl.ListColumns
        colMap(lc.Name) = lc.Index
    Next lc

    ' 5. Suppress flicker + disable AutoSave
    Dim wasAutoSave As Boolean
    wasAutoSave = ThisWorkbook.AutoSaveOn
    If wasAutoSave Then ThisWorkbook.AutoSaveOn = False
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 6. Cache logo once
    Dim assetWs As Worksheet, srcLogo As Shape, tmpShp As Shape
    Dim logoAvailable As Boolean: logoAvailable = False
    On Error Resume Next
    Set assetWs = ThisWorkbook.Worksheets("Assets")
    On Error GoTo 0
    If Not assetWs Is Nothing Then
        Dim wasHiddenA As Boolean
        wasHiddenA = (assetWs.Visible <> xlSheetVisible)
        If wasHiddenA Then assetWs.Visible = xlSheetVisible
        For Each tmpShp In assetWs.Shapes
            If tmpShp.Type = msoPicture Or tmpShp.Type = msoLinkedPicture Then
                If Not Intersect(tmpShp.TopLeftCell, assetWs.Range("A1")) Is Nothing Then
                    Set srcLogo = tmpShp: logoAvailable = True: Exit For
                End If
            End If
        Next tmpShp
        If Not logoAvailable Then
            For Each tmpShp In assetWs.Shapes
                If tmpShp.Type = msoPicture Or tmpShp.Type = msoLinkedPicture Then
                    Set srcLogo = tmpShp: logoAvailable = True: Exit For
                End If
            Next tmpShp
        End If
        If wasHiddenA Then assetWs.Visible = xlSheetVeryHidden
    End If

    ' 7. Make template temporarily visible for copying
    Dim tplWasHidden As Boolean
    tplWasHidden = (wsTpl.Visible <> xlSheetVisible)
    If tplWasHidden Then wsTpl.Visible = xlSheetVisible

    ' 8. Temp workbook
    Dim wbTemp As Workbook
    Set wbTemp = Workbooks.Add
    Dim pageCount As Long: pageCount = 0

    ' 9. Loop visible rows (max 50)
    Const MAX_ROWS As Long = 50
    Dim exportedRows As Long: exportedRows = 0
    Dim tRow As ListRow

    For Each tRow In tbl.ListRows
        If tRow.Range.Rows(1).Hidden Then GoTo NextRow
        If exportedRows >= MAX_ROWS Then GoTo ExitLoop

        ' Read row data
        Dim sTicket As String, sHauler As String, sOrigin As String
        Dim sProd As String, sTruck As String, sPayment As String
        Dim sPaid As String, sDate1 As String, sTime1 As String
        Dim sTime2 As String, sInWt As String, sOutWt As String
        Dim sNetLb As String, sNetTn As String, sSeller As String
        Dim sWaste As String, sWM As String, sByF As String

        sTicket = "": sHauler = "": sOrigin = "": sProd = ""
        sTruck = "": sPayment = "": sPaid = "": sDate1 = ""
        sTime1 = "": sTime2 = "": sInWt = "": sOutWt = ""
        sNetLb = "": sNetTn = "": sSeller = "": sWaste = ""
        sWM = "": sByF = "Jesus Polanco"

        If colMap.Exists("Weighmaster")     Then sTicket  = Trim(CStr(tRow.Range(1, colMap("Weighmaster")).Value))
        If colMap.Exists("Account")         Then sHauler  = Trim(CStr(tRow.Range(1, colMap("Account")).Value))
        If colMap.Exists("PO# / ORIGIN")    Then sOrigin  = Trim(CStr(tRow.Range(1, colMap("PO# / ORIGIN")).Value))
        If colMap.Exists("TRUCKID")         Then sTruck   = Trim(CStr(tRow.Range(1, colMap("TRUCKID")).Value))
        If colMap.Exists("PAYMENT")         Then sPayment = Trim(CStr(tRow.Range(1, colMap("PAYMENT")).Value))
        If colMap.Exists("Memo")            Then sSeller  = Trim(CStr(tRow.Range(1, colMap("Memo")).Value))
        If colMap.Exists("WASTE DATA")      Then sWaste   = Trim(CStr(tRow.Range(1, colMap("WASTE DATA")).Value))
        If colMap.Exists("Scalemaster")     Then sWM      = Trim(CStr(tRow.Range(1, colMap("Scalemaster")).Value))
        If colMap.Exists("Product/Service") Then
            sProd = Trim(CStr(tRow.Range(1, colMap("Product/Service")).Value))
            If UCase(sProd) = "MIXED" Then sProd = "Mixed C&D"
        End If

        Dim dcn As String: dcn = ""
        If colMap.Exists("Date") Then dcn = "Date"
        If colMap.Exists("DATE") Then dcn = "DATE"
        If dcn <> "" Then
            Dim dv As Variant: dv = tRow.Range(1, colMap(dcn)).Value
            If CStr(dv) <> "" And (IsNumeric(dv) Or IsDate(dv)) Then sDate1 = Format(CDate(dv), "mm/dd/yyyy") Else sDate1 = CStr(dv)
        End If
        If colMap.Exists("TIME IN") Then
            Dim ti As Variant: ti = tRow.Range(1, colMap("TIME IN")).Value
            If CStr(ti) <> "" And (IsNumeric(ti) Or IsDate(ti)) Then
                If sDate1 = "" Then sDate1 = Format(CDate(ti), "mm/dd/yyyy")
                sTime1 = Format(CDate(ti), "hh:mm AM/PM")
            End If
        End If
        If colMap.Exists("TIME OUT") Then
            Dim to_ As Variant: to_ = tRow.Range(1, colMap("TIME OUT")).Value
            If CStr(to_) <> "" And (IsNumeric(to_) Or IsDate(to_)) Then sTime2 = Format(CDate(to_), "hh:mm AM/PM")
        End If
        If colMap.Exists("T1") Then
            Dim v1 As Variant: v1 = tRow.Range(1, colMap("T1")).Value
            If IsNumeric(v1) And CDbl(v1) <> 0 Then sInWt = CStr(CLng(v1))
        End If
        If colMap.Exists("T2") Then
            Dim v2 As Variant: v2 = tRow.Range(1, colMap("T2")).Value
            If IsNumeric(v2) And CDbl(v2) <> 0 Then sOutWt = CStr(CLng(v2))
        End If

        Dim inL As Double, outL As Double, netL As Double, netT As Double
        inL = 0: outL = 0: netL = 0: netT = 0
        If IsNumeric(sInWt)  And sInWt  <> "" Then inL  = CDbl(sInWt)
        If IsNumeric(sOutWt) And sOutWt <> "" Then outL = CDbl(sOutWt)
        If inL > 0 And outL > 0 Then
            netL = Abs(inL - outL): netT = netL / 2000
            If netT < 1 And netT > 0 Then netT = 1
        End If
        If netL > 0 Then sNetLb = Format(netL, "#,##0") & " lb"
        If netT > 0 Then sNetTn = Format(netT, "0.00")

        If UCase(sPayment) = "AR" Or UCase(sPayment) = "EXPENSE" Then sPaid = ""
        If UCase(sPayment) = "EXPENSE" Then sPayment = ""
        If sTicket = "" Then GoTo NextRow

        ' ── Copy template into temp workbook ──
        wsTpl.Copy After:=wbTemp.Sheets(wbTemp.Sheets.Count)
        Dim wsDest As Worksheet
        Set wsDest = wbTemp.Sheets(wbTemp.Sheets.Count)

        ' ── Fill placeholder tokens with real data ──
        Call FillTicketData(wsDest, sTicket, sHauler, sOrigin, sProd, sTruck, sPayment, _
                            sPaid, sDate1, sTime1, sTime2, sInWt, sOutWt, _
                            sNetLb, sNetTn, sSeller, sWaste, sWM, sByF)

        ' ── Paste logo ──
        Const R As Long = 2
        If logoAvailable Then
            Dim wasH As Boolean
            wasH = (assetWs.Visible <> xlSheetVisible)
            If wasH Then assetWs.Visible = xlSheetVisible
            srcLogo.CopyPicture xlScreen, xlPicture
            If wasH Then assetWs.Visible = xlSheetVeryHidden
            DoEvents
            On Error Resume Next
            wsDest.Paste Destination:=wsDest.Range("C" & R)
            On Error GoTo 0
            If wsDest.Shapes.Count > 0 Then
                Dim ps As Shape: Set ps = wsDest.Shapes(wsDest.Shapes.Count)
                Dim lr As Range: Set lr = wsDest.Range(wsDest.Cells(R, 3), wsDest.Cells(R + 2, 7))
                ps.LockAspectRatio = msoTrue
                Dim tW As Double, tH As Double
                tW = lr.Width - 4: tH = lr.Height - 4
                If (ps.Width / ps.Height) > (tW / tH) Then ps.Width = tW Else ps.Height = tH
                ps.Left = lr.Left + 2
                ps.Top  = lr.Top + (lr.Height - ps.Height) / 2
            End If
        End If

        pageCount = pageCount + 1
        exportedRows = exportedRows + 1
        DoEvents

NextRow:
    Next tRow

ExitLoop:

    If tplWasHidden Then wsTpl.Visible = xlSheetVeryHidden

    If pageCount = 0 Then
        wbTemp.Close SaveChanges:=False
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        If wasAutoSave Then ThisWorkbook.AutoSaveOn = True
        MsgBox "No visible tickets found.", vbInformation
        Exit Sub
    End If

    ' Remove default blank Sheet1
    Application.DisplayAlerts = False
    If wbTemp.Sheets.Count > pageCount Then wbTemp.Sheets(1).Delete
    Application.DisplayAlerts = True

    ' Export all tabs as one PDF
    Dim errNum As Long
    On Error Resume Next
    wbTemp.Sheets.Select
    wbTemp.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, _
        Quality:=xlQualityStandard, IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, OpenAfterPublish:=False
    errNum = Err.Number
    On Error GoTo 0

    wbTemp.Close SaveChanges:=False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If wasAutoSave Then ThisWorkbook.AutoSaveOn = True

    If errNum = 0 Then
        Dim msg As String
        msg = "Done! " & pageCount & " ticket(s) exported to:" & vbCrLf & pdfPath
        If exportedRows >= MAX_ROWS Then _
            msg = msg & vbCrLf & vbCrLf & "(Export limited to " & MAX_ROWS & " rows — additional rows were skipped.)"
        MsgBox msg, vbInformation
    Else
        MsgBox "Export failed (#" & errNum & "). File may already be open.", vbCritical
    End If

End Sub
