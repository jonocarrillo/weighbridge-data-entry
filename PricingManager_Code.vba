        ' ========================================
        ' PRICING MANAGER FORM CODE
        ' Manages ScalePricing and HaulingPricing tables
        ' Paste this into PricingManagerForm's code view (F7)
        ' ========================================

        Private isLoading As Boolean
        Private currentTable As String   ' "SCALE" or "HAUL"
        Private selectedRow As Long      ' 1-based table row, 0 = none

        ' ========================================
        ' INITIALIZE
        ' ========================================
        Private Sub UserForm_Initialize()
            isLoading = True
            currentTable = "SCALE"
            selectedRow = 0
            Me.Controls("cmdScaleTab").BackColor = RGB(0, 120, 215)
            Me.Controls("cmdScaleTab").ForeColor = RGB(255, 255, 255)
            Me.Controls("cmdHaulTab").BackColor = RGB(220, 220, 220)
            Me.Controls("cmdHaulTab").ForeColor = RGB(50, 50, 50)
            ShowScaleFields True
            ShowHaulFields False
            LoadScaleList
            ClearScaleFields
            isLoading = False
        End Sub

        ' ========================================
        ' TAB SWITCHING
        ' ========================================
        Private Sub cmdScaleTab_Click()
            currentTable = "SCALE"
            selectedRow = 0
            Me.Controls("cmdScaleTab").BackColor = RGB(0, 120, 215)
            Me.Controls("cmdScaleTab").ForeColor = RGB(255, 255, 255)
            Me.Controls("cmdHaulTab").BackColor = RGB(220, 220, 220)
            Me.Controls("cmdHaulTab").ForeColor = RGB(50, 50, 50)
            ShowScaleFields True
            ShowHaulFields False
            LoadScaleList
            ClearScaleFields
            UpdateStatusBar "Scale Pricing table loaded."
        End Sub

        Private Sub cmdHaulTab_Click()
            currentTable = "HAUL"
            selectedRow = 0
            Me.Controls("cmdHaulTab").BackColor = RGB(0, 120, 215)
            Me.Controls("cmdHaulTab").ForeColor = RGB(255, 255, 255)
            Me.Controls("cmdScaleTab").BackColor = RGB(220, 220, 220)
            Me.Controls("cmdScaleTab").ForeColor = RGB(50, 50, 50)
            ShowScaleFields False
            ShowHaulFields True
            LoadHaulList
            ClearHaulFields
            UpdateStatusBar "Hauling Pricing table loaded."
        End Sub

        Private Sub ShowScaleFields(vis As Boolean)
            Me.Controls("lstScale").Visible = vis
            Dim n As String
            For Each n In Array("lblSAccount","lblSProduct","lblSStart","lblSEnd","lblSRate","lblSType", _
                                "txtSAccount","txtSProduct","txtSStart","txtSEnd","txtSRate","cmbSType", _
                                "lblSHdr1","lblSHdr2","lblSHdr3","lblSHdr4","lblSHdr5","lblSHdr6")
                On Error Resume Next
                Me.Controls(n).Visible = vis
                On Error GoTo 0
            Next n
        End Sub

        Private Sub ShowHaulFields(vis As Boolean)
            Me.Controls("lstHaul").Visible = vis
            Dim n As String
            For Each n In Array("lblHAccount","lblHStart","lblHEnd","lblHLowboy","lblHConcrete", _
                                "lblH10YD","lblH40YD","lblHAllowed","lblHOverage","lblHMiles","lblHDistRate","lblHSurch", _
                                "txtHAccount","txtHStart","txtHEnd","txtHLowboy","txtHConcrete", _
                                "txtH10YD","txtH40YD","txtHAllowed","txtHOverage","txtHMiles","txtHDistRate","txtHSurch", _
                                "lblHHdr1","lblHHdr2","lblHHdr3","lblHHdr4","lblHHdr5","lblHHdr6")
                On Error Resume Next
                Me.Controls(n).Visible = vis
                On Error GoTo 0
            Next n
        End Sub

        ' ========================================
        ' LOAD LIST — SCALE
        ' ========================================
        Private Sub LoadScaleList()
            isLoading = True
            Me.Controls("lstScale").Clear

            Dim tbl As ListObject
            Set tbl = GetTable("ScalePricing")
            If tbl Is Nothing Then
                UpdateStatusBar "ERROR: ScalePricing table not found."
                isLoading = False
                Exit Sub
            End If

            Dim ciProd As Long, ciAcct As Long, ciStart As Long, ciEnd As Long
            Dim ciRate As Long, ciType As Long
            On Error Resume Next
            ciProd  = tbl.ListColumns("PRODUCT/SERVICE").Index
            ciAcct  = tbl.ListColumns("Account").Index
            ciStart = tbl.ListColumns("Start").Index
            ciEnd   = tbl.ListColumns("END").Index
            ciRate  = tbl.ListColumns("Rate").Index
            ciType  = tbl.ListColumns("Type").Index
            On Error GoTo 0

            Dim R As Long
            For R = 1 To tbl.ListRows.Count
                Dim prod As String, acct As String, sDate As String, eDate As String
                Dim rt As String, tp As String
                On Error Resume Next
                prod  = CStr(tbl.DataBodyRange.Cells(R, ciProd).Value)
                acct  = CStr(tbl.DataBodyRange.Cells(R, ciAcct).Value)
                sDate = Format(tbl.DataBodyRange.Cells(R, ciStart).Value, "mm/dd")
                eDate = Format(tbl.DataBodyRange.Cells(R, ciEnd).Value, "mm/dd")
                rt    = CStr(tbl.DataBodyRange.Cells(R, ciRate).Value)
                tp    = CStr(tbl.DataBodyRange.Cells(R, ciType).Value)
                On Error GoTo 0

                Dim display As String
                display = PadRight(UCase(prod), 12) & PadRight(IIf(acct = "", "(public)", acct), 20) & _
                          PadRight(sDate & "-" & eDate, 14) & PadRight("$" & Format(CDbl(rt), "0.00"), 10) & tp
                Me.Controls("lstScale").AddItem display
            Next R

            UpdateStatusBar "ScalePricing: " & tbl.ListRows.Count & " rows."
            isLoading = False
        End Sub

        ' ========================================
        ' LOAD LIST — HAUL
        ' ========================================
        Private Sub LoadHaulList()
            isLoading = True
            Me.Controls("lstHaul").Clear

            Dim tbl As ListObject
            Set tbl = GetTable("HaulingPricing")
            If tbl Is Nothing Then
                UpdateStatusBar "ERROR: HaulingPricing table not found."
                isLoading = False
                Exit Sub
            End If

            Dim hiAcct As Long, hiStart As Long, hiEnd As Long
            Dim hiLowboy As Long, hiConc As Long, hi10 As Long, hi40 As Long, hiAllow As Long
            On Error Resume Next
            hiAcct   = tbl.ListColumns("Account").Index
            hiStart  = tbl.ListColumns("Start").Index
            hiEnd    = tbl.ListColumns("End").Index
            hiLowboy = tbl.ListColumns("Lowboy").Index
            hiConc   = tbl.ListColumns("CONCRETE").Index
            hi10     = tbl.ListColumns("10YD").Index
            hi40     = tbl.ListColumns("40YD").Index
            hiAllow  = tbl.ListColumns("Allowed Tons").Index
            On Error GoTo 0

            Dim R As Long
            For R = 1 To tbl.ListRows.Count
                Dim acct As String, sDate As String, eDate As String
                Dim lb As String, cn As String, t10 As String, t40 As String, allow As String
                On Error Resume Next
                acct  = CStr(tbl.DataBodyRange.Cells(R, hiAcct).Value)
                sDate = Format(tbl.DataBodyRange.Cells(R, hiStart).Value, "mm/dd")
                eDate = Format(tbl.DataBodyRange.Cells(R, hiEnd).Value, "mm/dd")
                lb    = CStr(tbl.DataBodyRange.Cells(R, hiLowboy).Value)
                cn    = CStr(tbl.DataBodyRange.Cells(R, hiConc).Value)
                t10   = CStr(tbl.DataBodyRange.Cells(R, hi10).Value)
                t40   = CStr(tbl.DataBodyRange.Cells(R, hi40).Value)
                allow = CStr(tbl.DataBodyRange.Cells(R, hiAllow).Value)
                On Error GoTo 0

                Dim display As String
                display = PadRight(IIf(acct = "", "(public)", acct), 22) & _
                          PadRight(sDate & "-" & eDate, 13) & _
                          PadRight("Low:$" & lb, 10) & PadRight("Con:$" & cn, 10) & _
                          PadRight("10:$" & t10, 9) & PadRight("40:$" & t40, 9) & _
                          "Allow:" & allow
                Me.Controls("lstHaul").AddItem display
            Next R

            UpdateStatusBar "HaulingPricing: " & tbl.ListRows.Count & " rows."
            isLoading = False
        End Sub

        ' ========================================
        ' LIST SELECTION — SCALE
        ' ========================================
        Private Sub lstScale_Click()
            If isLoading Then Exit Sub
            selectedRow = Me.Controls("lstScale").ListIndex + 1
            If selectedRow < 1 Then Exit Sub
            LoadScaleRow selectedRow
        End Sub

        Private Sub LoadScaleRow(R As Long)
            Dim tbl As ListObject
            Set tbl = GetTable("ScalePricing")
            If tbl Is Nothing Or R < 1 Or R > tbl.ListRows.Count Then Exit Sub

            Dim ciProd As Long, ciAcct As Long, ciStart As Long, ciEnd As Long
            Dim ciRate As Long, ciType As Long
            On Error Resume Next
            ciProd  = tbl.ListColumns("PRODUCT/SERVICE").Index
            ciAcct  = tbl.ListColumns("Account").Index
            ciStart = tbl.ListColumns("Start").Index
            ciEnd   = tbl.ListColumns("END").Index
            ciRate  = tbl.ListColumns("Rate").Index
            ciType  = tbl.ListColumns("Type").Index
            On Error GoTo 0

            isLoading = True
            Me.Controls("txtSProduct").Value = CStr(tbl.DataBodyRange.Cells(R, ciProd).Value)
            Me.Controls("txtSAccount").Value = CStr(tbl.DataBodyRange.Cells(R, ciAcct).Value)
            Me.Controls("txtSStart").Value   = Format(tbl.DataBodyRange.Cells(R, ciStart).Value, "mm/dd/yyyy")
            Me.Controls("txtSEnd").Value     = Format(tbl.DataBodyRange.Cells(R, ciEnd).Value, "mm/dd/yyyy")
            Me.Controls("txtSRate").Value    = CStr(tbl.DataBodyRange.Cells(R, ciRate).Value)
            Me.Controls("cmbSType").Value    = CStr(tbl.DataBodyRange.Cells(R, ciType).Value)
            isLoading = False
            UpdateStatusBar "Row " & R & " loaded. Edit fields below and click Save Row."
        End Sub

        ' ========================================
        ' LIST SELECTION — HAUL
        ' ========================================
        Private Sub lstHaul_Click()
            If isLoading Then Exit Sub
            selectedRow = Me.Controls("lstHaul").ListIndex + 1
            If selectedRow < 1 Then Exit Sub
            LoadHaulRow selectedRow
        End Sub

        Private Sub LoadHaulRow(R As Long)
            Dim tbl As ListObject
            Set tbl = GetTable("HaulingPricing")
            If tbl Is Nothing Or R < 1 Or R > tbl.ListRows.Count Then Exit Sub

            Dim hiAcct As Long, hiStart As Long, hiEnd As Long
            Dim hiLowboy As Long, hiConc As Long, hi10 As Long, hi40 As Long
            Dim hiAllow As Long, hiOver As Long, hiMiles As Long, hiDist As Long, hiSurch As Long
            On Error Resume Next
            hiAcct   = tbl.ListColumns("Account").Index
            hiStart  = tbl.ListColumns("Start").Index
            hiEnd    = tbl.ListColumns("End").Index
            hiLowboy = tbl.ListColumns("Lowboy").Index
            hiConc   = tbl.ListColumns("CONCRETE").Index
            hi10     = tbl.ListColumns("10YD").Index
            hi40     = tbl.ListColumns("40YD").Index
            hiAllow  = tbl.ListColumns("Allowed Tons").Index
            hiOver   = tbl.ListColumns("Overage").Index
            hiMiles  = tbl.ListColumns("Miles").Index
            hiDist   = tbl.ListColumns("Dist. Rate").Index
            hiSurch  = tbl.ListColumns("Surcharge").Index
            On Error GoTo 0

            isLoading = True
            Me.Controls("txtHAccount").Value  = CStr(tbl.DataBodyRange.Cells(R, hiAcct).Value)
            Me.Controls("txtHStart").Value    = Format(tbl.DataBodyRange.Cells(R, hiStart).Value, "mm/dd/yyyy")
            Me.Controls("txtHEnd").Value      = Format(tbl.DataBodyRange.Cells(R, hiEnd).Value, "mm/dd/yyyy")
            Me.Controls("txtHLowboy").Value   = CStr(tbl.DataBodyRange.Cells(R, hiLowboy).Value)
            Me.Controls("txtHConcrete").Value = CStr(tbl.DataBodyRange.Cells(R, hiConc).Value)
            Me.Controls("txtH10YD").Value     = CStr(tbl.DataBodyRange.Cells(R, hi10).Value)
            Me.Controls("txtH40YD").Value     = CStr(tbl.DataBodyRange.Cells(R, hi40).Value)
            Me.Controls("txtHAllowed").Value  = CStr(tbl.DataBodyRange.Cells(R, hiAllow).Value)
            Me.Controls("txtHOverage").Value  = CStr(tbl.DataBodyRange.Cells(R, hiOver).Value)
            Me.Controls("txtHMiles").Value    = CStr(tbl.DataBodyRange.Cells(R, hiMiles).Value)
            Me.Controls("txtHDistRate").Value = CStr(tbl.DataBodyRange.Cells(R, hiDist).Value)
            Me.Controls("txtHSurch").Value    = CStr(tbl.DataBodyRange.Cells(R, hiSurch).Value)
            isLoading = False
            UpdateStatusBar "Row " & R & " loaded. Edit fields below and click Save Row."
        End Sub

        ' ========================================
        ' NEW ROW
        ' ========================================
        Private Sub cmdNew_Click()
            selectedRow = 0
            If currentTable = "SCALE" Then
                ClearScaleFields
                Me.Controls("lstScale").ListIndex = -1
                Me.Controls("txtSProduct").SetFocus
            Else
                ClearHaulFields
                Me.Controls("lstHaul").ListIndex = -1
                Me.Controls("txtHAccount").SetFocus
            End If
            UpdateStatusBar "New row — fill in fields and click Save Row."
        End Sub

        ' ========================================
        ' SAVE ROW
        ' ========================================
        Private Sub cmdSave_Click()
            If currentTable = "SCALE" Then
                SaveScaleRow
            Else
                SaveHaulRow
            End If
        End Sub

        Private Sub SaveScaleRow()
            ' Validate
            If Trim(Me.Controls("txtSProduct").Value) = "" Then
                MsgBox "Product/Service is required.", vbExclamation: Exit Sub
            End If
            If Trim(Me.Controls("txtSRate").Value) = "" Then
                MsgBox "Rate is required.", vbExclamation: Exit Sub
            End If

            Dim tbl As ListObject
            Set tbl = GetTable("ScalePricing")
            If tbl Is Nothing Then MsgBox "ScalePricing table not found.", vbCritical: Exit Sub

            Dim ciProd As Long, ciAcct As Long, ciStart As Long, ciEnd As Long
            Dim ciRate As Long, ciType As Long
            On Error Resume Next
            ciProd  = tbl.ListColumns("PRODUCT/SERVICE").Index
            ciAcct  = tbl.ListColumns("Account").Index
            ciStart = tbl.ListColumns("Start").Index
            ciEnd   = tbl.ListColumns("END").Index
            ciRate  = tbl.ListColumns("Rate").Index
            ciType  = tbl.ListColumns("Type").Index
            On Error GoTo 0

            Dim targetRow As ListRow
            If selectedRow > 0 And selectedRow <= tbl.ListRows.Count Then
                Set targetRow = tbl.ListRows(selectedRow)
            Else
                Set targetRow = tbl.ListRows.Add
            End If

            Application.EnableEvents = False
            On Error Resume Next
            targetRow.Range(1, ciProd).Value  = UCase(Trim(Me.Controls("txtSProduct").Value))
            targetRow.Range(1, ciAcct).Value  = Trim(Me.Controls("txtSAccount").Value)
            If IsDate(Me.Controls("txtSStart").Value) Then
                targetRow.Range(1, ciStart).Value = CDate(Me.Controls("txtSStart").Value)
                targetRow.Range(1, ciStart).NumberFormat = "mm/dd/yyyy"
            End If
            If IsDate(Me.Controls("txtSEnd").Value) Then
                targetRow.Range(1, ciEnd).Value = CDate(Me.Controls("txtSEnd").Value)
                targetRow.Range(1, ciEnd).NumberFormat = "mm/dd/yyyy"
            End If
            targetRow.Range(1, ciRate).Value = CDbl(Me.Controls("txtSRate").Value)
            targetRow.Range(1, ciType).Value = LCase(Trim(Me.Controls("cmbSType").Value))
            On Error GoTo 0
            Application.EnableEvents = True

            LoadScaleList
            UpdateStatusBar "Scale pricing row saved successfully."
        End Sub

        Private Sub SaveHaulRow()
            Dim tbl As ListObject
            Set tbl = GetTable("HaulingPricing")
            If tbl Is Nothing Then MsgBox "HaulingPricing table not found.", vbCritical: Exit Sub

            Dim hiAcct As Long, hiStart As Long, hiEnd As Long
            Dim hiLowboy As Long, hiConc As Long, hi10 As Long, hi40 As Long
            Dim hiAllow As Long, hiOver As Long, hiMiles As Long, hiDist As Long, hiSurch As Long
            On Error Resume Next
            hiAcct   = tbl.ListColumns("Account").Index
            hiStart  = tbl.ListColumns("Start").Index
            hiEnd    = tbl.ListColumns("End").Index
            hiLowboy = tbl.ListColumns("Lowboy").Index
            hiConc   = tbl.ListColumns("CONCRETE").Index
            hi10     = tbl.ListColumns("10YD").Index
            hi40     = tbl.ListColumns("40YD").Index
            hiAllow  = tbl.ListColumns("Allowed Tons").Index
            hiOver   = tbl.ListColumns("Overage").Index
            hiMiles  = tbl.ListColumns("Miles").Index
            hiDist   = tbl.ListColumns("Dist. Rate").Index
            hiSurch  = tbl.ListColumns("Surcharge").Index
            On Error GoTo 0

            Dim targetRow As ListRow
            If selectedRow > 0 And selectedRow <= tbl.ListRows.Count Then
                Set targetRow = tbl.ListRows(selectedRow)
            Else
                Set targetRow = tbl.ListRows.Add
            End If

            Application.EnableEvents = False
            On Error Resume Next
            targetRow.Range(1, hiAcct).Value  = Trim(Me.Controls("txtHAccount").Value)
            If IsDate(Me.Controls("txtHStart").Value) Then
                targetRow.Range(1, hiStart).Value = CDate(Me.Controls("txtHStart").Value)
                targetRow.Range(1, hiStart).NumberFormat = "mm/dd/yyyy"
            End If
            If IsDate(Me.Controls("txtHEnd").Value) Then
                targetRow.Range(1, hiEnd).Value = CDate(Me.Controls("txtHEnd").Value)
                targetRow.Range(1, hiEnd).NumberFormat = "mm/dd/yyyy"
            End If
            targetRow.Range(1, hiLowboy).Value = SafeNum(Me.Controls("txtHLowboy").Value)
            targetRow.Range(1, hiConc).Value   = SafeNum(Me.Controls("txtHConcrete").Value)
            targetRow.Range(1, hi10).Value     = SafeNum(Me.Controls("txtH10YD").Value)
            targetRow.Range(1, hi40).Value     = SafeNum(Me.Controls("txtH40YD").Value)
            targetRow.Range(1, hiAllow).Value  = Trim(Me.Controls("txtHAllowed").Value)
            targetRow.Range(1, hiOver).Value   = SafeNum(Me.Controls("txtHOverage").Value)
            targetRow.Range(1, hiMiles).Value  = SafeNum(Me.Controls("txtHMiles").Value)
            targetRow.Range(1, hiDist).Value   = SafeNum(Me.Controls("txtHDistRate").Value)
            targetRow.Range(1, hiSurch).Value  = SafeNum(Me.Controls("txtHSurch").Value)
            On Error GoTo 0
            Application.EnableEvents = True

            LoadHaulList
            UpdateStatusBar "Hauling pricing row saved successfully."
        End Sub

        ' ========================================
        ' DELETE ROW
        ' ========================================
        Private Sub cmdDelete_Click()
            If selectedRow < 1 Then
                MsgBox "Select a row first.", vbExclamation: Exit Sub
            End If

            Dim tblName As String
            tblName = IIf(currentTable = "SCALE", "ScalePricing", "HaulingPricing")
            If MsgBox("Delete row " & selectedRow & " from " & tblName & "?", _
                      vbYesNo + vbExclamation, "Confirm Delete") <> vbYes Then Exit Sub

            Dim tbl As ListObject
            Set tbl = GetTable(tblName)
            If tbl Is Nothing Then Exit Sub

            Application.EnableEvents = False
            On Error Resume Next
            tbl.ListRows(selectedRow).Delete
            On Error GoTo 0
            Application.EnableEvents = True

            selectedRow = 0
            If currentTable = "SCALE" Then
                LoadScaleList: ClearScaleFields
            Else
                LoadHaulList: ClearHaulFields
            End If
            UpdateStatusBar "Row deleted."
        End Sub

        ' ========================================
        ' DUPLICATE ROW (copy selected as template for new)
        ' ========================================
        Private Sub cmdDuplicate_Click()
            If selectedRow < 1 Then
                MsgBox "Select a row to duplicate.", vbExclamation: Exit Sub
            End If
            ' Fields are already loaded from selection — just clear selectedRow so Save adds a new row
            selectedRow = 0
            UpdateStatusBar "Row duplicated as new — edit and click Save Row."
        End Sub

        ' ========================================
        ' CLEAR FIELDS
        ' ========================================
        Private Sub ClearScaleFields()
            Me.Controls("txtSProduct").Value  = ""
            Me.Controls("txtSAccount").Value  = ""
            Me.Controls("txtSStart").Value    = ""
            Me.Controls("txtSEnd").Value      = ""
            Me.Controls("txtSRate").Value     = ""
            Me.Controls("cmbSType").Value     = "ton"
        End Sub

        Private Sub ClearHaulFields()
            Dim n As String
            For Each n In Array("txtHAccount","txtHStart","txtHEnd","txtHLowboy","txtHConcrete", _
                                "txtH10YD","txtH40YD","txtHAllowed","txtHOverage","txtHMiles","txtHDistRate","txtHSurch")
                Me.Controls(n).Value = ""
            Next n
        End Sub

        ' ========================================
        ' CLOSE
        ' ========================================
        Private Sub cmdClose_Click()
            Unload Me
        End Sub

        ' ========================================
        ' HELPERS
        ' ========================================
        Private Function GetTable(tblName As String) As ListObject
            Dim ws As Worksheet
            Dim tbl As ListObject
            On Error Resume Next
            For Each ws In ThisWorkbook.Worksheets
                Set tbl = ws.ListObjects(tblName)
                If Not tbl Is Nothing Then
                    Set GetTable = tbl
                    Exit Function
                End If
            Next ws
            On Error GoTo 0
            Set GetTable = Nothing
        End Function

        Private Function SafeNum(v As String) As Double
            If IsNumeric(Trim(v)) Then SafeNum = CDbl(Trim(v)) Else SafeNum = 0
        End Function

        Private Function PadRight(s As String, n As Integer) As String
            PadRight = Left(s & Space(n), n)
        End Function

        Private Sub UpdateStatusBar(msg As String)
            On Error Resume Next
            Me.Controls("lblStatus").Caption = msg
            On Error GoTo 0
        End Sub
