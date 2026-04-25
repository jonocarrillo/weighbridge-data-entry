        ' ========================================
        ' USERFORM CODE - WEIGHMASTER CERTIFICATE
        ' Multi-ticket navigation support
        ' ========================================

        ' --- TICKET STORAGE (using 2D array) ---
        Private ticketData() As String
        Private ticketCount As Long
        Private currentTicket As Long
        Private isInitializing As Boolean
        ' True once the user has manually typed into txtPaidAmount.
        ' Set in txtPaidAmount_Change. RefreshAmount honors this and won't
        ' overwrite the user's value. Reset on ClearFormFields / new ticket /
        ' loading an existing ticket.
        Private userOverrodeAmount As Boolean
        Private strWeighmaster As String     ' Set once at login, reused each ticket

        ' ========================================
        ' INITIALIZATION
        ' ========================================

        Private Sub UserForm_Initialize()
            isInitializing = True
            
            ticketCount = 1
            currentTicket = 1
            ReDim ticketData(1 To 1, 1 To 22)
            
            PopulateDropdowns
            
            ' --- Weighmaster login removed based on user request ---
            ' The fields will be blank for a new ticket, or you can type them directly into the form
            strWeighmaster = ""
            
            ClearFormFields
            Me.txtTicketNumber.Value = GetNextTicketNumber()
            Me.txtDate1.Value = Format(Date, "mm/dd/yyyy")
            Me.txtWeighmaster.Value = strWeighmaster
            Me.txtBy.Value = "Jesus Polanco"
            
            UpdateTicketCounter
            
            ' Pre-create the permanent print sheet NOW (while idle) so it is
            ' already baked into the saved workbook before any print action runs.
            ' This prevents GetOrCreatePrintSheet() from needing Worksheets.Add()
            ' mid-print, which would trigger an OneDrive sync at the worst moment.
            EnsurePrintSheetExists
            
            isInitializing = False
            Me.txtWeighedForSeller.SetFocus
        End Sub

        Private Function UserFormTableHasColumn(ByVal tbl As ListObject, ByVal colName As String) As Boolean
            If tbl Is Nothing Then Exit Function
            
            On Error Resume Next
            UserFormTableHasColumn = Not tbl.ListColumns(colName) Is Nothing
            On Error GoTo 0
        End Function

        Private Function IsUserFormTicketTable(ByVal tbl As ListObject) As Boolean
            If tbl Is Nothing Then Exit Function
            
            IsUserFormTicketTable = UserFormTableHasColumn(tbl, "Weighmaster") And _
                                    UserFormTableHasColumn(tbl, "Account") And _
                                    UserFormTableHasColumn(tbl, "Product/Service")
        End Function

        Private Function FindUserFormTicketTableOnSheet(ByVal ws As Worksheet) As ListObject
            Dim tbl As ListObject
            
            If ws Is Nothing Then Exit Function
            
            On Error Resume Next
            Set tbl = ws.ListObjects("Supercluster")
            On Error GoTo 0
            If IsUserFormTicketTable(tbl) Then
                Set FindUserFormTicketTableOnSheet = tbl
                Exit Function
            End If

            For Each tbl In ws.ListObjects
                If IsUserFormTicketTable(tbl) Then
                    Set FindUserFormTicketTableOnSheet = tbl
                    Exit Function
                End If
            Next tbl
        End Function

        Private Function TryGetUserFormTicketTable(ByRef ws As Worksheet, ByRef tbl As ListObject) As Boolean
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
                
                If IsUserFormTicketTable(tbl) Then
                    Set ws = wsLoop
                    TryGetUserFormTicketTable = True
                    Exit Function
                End If
                
                Set tbl = Nothing
            Next wsLoop
            
            For i = LBound(preferredSheets) To UBound(preferredSheets)
                On Error Resume Next
                Set ws = ThisWorkbook.Worksheets(CStr(preferredSheets(i)))
                On Error GoTo 0
                
                If Not ws Is Nothing Then
                    Set tbl = FindUserFormTicketTableOnSheet(ws)
                    If Not tbl Is Nothing Then
                        TryGetUserFormTicketTable = True
                        Exit Function
                    End If
                End If
                
                Set ws = Nothing
            Next i
            
            For Each wsLoop In ThisWorkbook.Worksheets
                Set tbl = FindUserFormTicketTableOnSheet(wsLoop)
                If Not tbl Is Nothing Then
                    Set ws = wsLoop
                    TryGetUserFormTicketTable = True
                    Exit Function
                End If
            Next wsLoop
        End Function

        Private Sub PopulateDropdowns()
            ' Populate comboboxes from the live ticket table
            Dim ws As Worksheet
            Dim tbl As ListObject
            Dim cell As Range
            Dim dict As Object
            Dim v As Variant
            
            If Not TryGetUserFormTicketTable(ws, tbl) Then Exit Sub
            ' ---------------------------------------------------------
            ' 1. Company Name / Hauler -> "Account" Column
            ' ---------------------------------------------------------
            Dim colAcct As ListColumn
            Set colAcct = Nothing
            On Error Resume Next
            Set colAcct = tbl.ListColumns("Account")
            If colAcct Is Nothing Then Set colAcct = tbl.ListColumns("Accounts") ' Try plural
            If colAcct Is Nothing Then Set colAcct = tbl.ListColumns(12) ' Fallback index
            On Error GoTo 0
            
            If Not colAcct Is Nothing Then
                Me.Controls("txtCompanyNameHauler").Clear
                
                ' Try Data Validation first (if standard list)
                Dim dvList As String
                Dim dvRange As Range
                Dim validationCell As Range
                
                ' Determine cell to check for Validation (First data row OR cell below header if empty)
                If tbl.ListRows.count > 0 Then
                    Set validationCell = colAcct.DataBodyRange.Cells(1, 1)
                Else
                    Set validationCell = colAcct.Range.Cells(2, 1)
                End If
                
                On Error Resume Next
                dvList = validationCell.Validation.Formula1
                On Error GoTo 0
                
                If dvList <> "" And Left(dvList, 1) = "=" Then
                    ' It's a range reference (e.g. =Sheet2!$A$1:$A$10)
                    Set dvRange = Range(dvList)
                    If Not dvRange Is Nothing Then
                        For Each cell In dvRange
                            Me.Controls("txtCompanyNameHauler").AddItem CStr(cell.Value)
                        Next cell
                    End If
                ElseIf dvList <> "" Then
                    ' It's a comma-separated list
                    v = Split(dvList, ",")
                    For i = LBound(v) To UBound(v)
                        Me.Controls("txtCompanyNameHauler").AddItem Trim(v(i))
                    Next i
                ElseIf tbl.ListRows.count > 0 Then
                    ' No validation found, fallback to unique values from column data (if any)
                    Set dict = CreateObject("Scripting.Dictionary")
                    If Not colAcct.DataBodyRange Is Nothing Then
                        For Each cell In colAcct.DataBodyRange
                            v = Trim(CStr(cell.Value))
                            If Len(v) > 0 And Not dict.Exists(v) Then dict.Add v, 1
                        Next cell
                        
                        Dim keys As Variant
                        keys = dict.keys
                        For i = LBound(keys) To UBound(keys)
                            Me.Controls("txtCompanyNameHauler").AddItem keys(i)
                        Next i
                    End If
                Else
                    ' Empty table, no validation found -> Can't populate anything
                    Me.Controls("txtCompanyNameHauler").AddItem "(Type or set Validation)"
                End If
            Else
                MsgBox "Debug: Column 'Account' (or 12) NOT FOUND in table.", vbExclamation
            End If
            
            ' Column N (14) = Product/Service -> txtProductCode
            ' ---------------------------------------------------------
            ' 2. Product/Service -> txtProductCode
            ' ---------------------------------------------------------
            Dim colProd As ListColumn
            Set colProd = Nothing
            On Error Resume Next
            Set colProd = tbl.ListColumns("Product/Service")
            If colProd Is Nothing Then Set colProd = tbl.ListColumns(14) ' Fallback index
            On Error GoTo 0
            
            If Not colProd Is Nothing Then
                Me.Controls("txtProductCode").Clear
                
                ' Try Data Validation first (if standard list)
                Dim dvListP As String
                Dim dvRangeP As Range
                On Error Resume Next
                dvListP = colProd.DataBodyRange.Cells(1, 1).Validation.Formula1
                On Error GoTo 0
                
                If dvListP <> "" And Left(dvListP, 1) = "=" Then
                    ' It's a range reference
                    Set dvRangeP = Range(dvListP)
                    If Not dvRangeP Is Nothing Then
                        For Each cell In dvRangeP
                            Me.Controls("txtProductCode").AddItem CStr(cell.Value)
                        Next cell
                    End If
                ElseIf dvListP <> "" Then
                    ' It's a comma-separated list
                    v = Split(dvListP, ",")
                    For i = LBound(v) To UBound(v)
                        Me.Controls("txtProductCode").AddItem Trim(v(i))
                    Next i
                Else
                    ' No validation found, use unique values from column
                    Set dict = CreateObject("Scripting.Dictionary")
                    For Each cell In colProd.DataBodyRange
                        v = Trim(CStr(cell.Value))
                        If Len(v) > 0 And Not dict.Exists(v) Then dict.Add v, 1
                    Next cell
                    For Each v In dict.keys
                        Me.Controls("txtProductCode").AddItem CStr(v)
                    Next v
                End If
            End If
            
            ' Column B (2) = Waste Data -> txtWasteOrigin (address portion)
            Set dict = CreateObject("Scripting.Dictionary")
            If Not tbl.ListColumns(2).DataBodyRange Is Nothing Then
                For Each cell In tbl.ListColumns(2).DataBodyRange
                    v = Trim(CStr(cell.Value))
                    If Len(v) > 0 And Not dict.Exists(v) Then dict.Add v, 1
                Next cell
            End If
            For Each v In dict.keys
                Me.Controls("txtWasteOrigin").AddItem CStr(v)
            Next v

            ' City dropdown (cmbCity) — hardcoded CA city list + unique values from ORIGIN column
            Dim dictCity As Object
            Set dictCity = CreateObject("Scripting.Dictionary")
            ' Seed with common Southern California cities
            Dim defaultCities As Variant
            defaultCities = Array( _
                "Long Beach", "Los Angeles", "Compton", "Carson", "Torrance", _
                "Hawthorne", "Inglewood", "Gardena", "Paramount", "Lynwood", _
                "South Gate", "Downey", "Norwalk", "Bellflower", "Lakewood", _
                "Signal Hill", "Harbor City", "Wilmington", "San Pedro", _
                "Rancho Dominguez", "Huntington Park", "Maywood", "Bell", _
                "Bell Gardens", "Commerce", "Vernon", "Montebello", _
                "Whittier", "La Mirada", "Santa Fe Springs", "Cerritos", _
                "Artesia", "Hawaiian Gardens", "Cypress", "Buena Park", _
                "Anaheim", "Orange", "Santa Ana", "Irvine", "Other")
            Dim ci As Integer
            Me.Controls("cmbCity").Clear
            For ci = 0 To UBound(defaultCities)
                If Not dictCity.Exists(defaultCities(ci)) Then
                    dictCity.Add defaultCities(ci), 1
                    Me.Controls("cmbCity").AddItem defaultCities(ci)
                End If
            Next ci
            ' Also pull any additional cities already stored in the ORIGIN column
            Dim colOrig As ListColumn
            On Error Resume Next
            Set colOrig = tbl.ListColumns("PO# / ORIGIN")
            If colOrig Is Nothing Then Set colOrig = tbl.ListColumns("ORIGIN")
            On Error GoTo 0
            If Not colOrig Is Nothing Then
                If Not colOrig.DataBodyRange Is Nothing Then
                    For Each cell In colOrig.DataBodyRange
                        Dim cellVal As String
                        cellVal = Trim(CStr(cell.Value))
                        If InStr(cellVal, ",") > 0 Then
                            Dim cityPart As String
                            cityPart = Trim(Left(cellVal, InStr(cellVal, ",") - 1))
                            If Len(cityPart) > 0 And Not dictCity.Exists(cityPart) Then
                                dictCity.Add cityPart, 1
                                Me.Controls("cmbCity").AddItem cityPart
                            End If
                        End If
                    Next cell
                End If
            End If
            
        End Sub

        Private Sub ClearFormFields()
            Me.Controls("txtWeighedAt").Caption = "AMERICAN INDUSTRIAL SERVICES"
            Me.txtTicketNumber.Value = ""
            Me.txtWeighedForSeller.Value = ""
            On Error Resume Next
            Me.Controls("cmbCity").Value = ""
            On Error GoTo 0
            Me.Controls("txtCompanyNameHauler").Value = ""
            Me.Controls("txtWasteOrigin").Value = ""
            Me.txtWasteData.Value = ""
            Me.Controls("txtProductCode").Value = ""
            Me.txtTruckNumber.Value = ""
            Me.optCredit.Value = False
            Me.optDebit.Value = False
            Me.optAR.Value = False
            Me.optCash.Value = False
            Me.optExpense.Value = False
            Me.txtPaidAmount.Value = ""
            userOverrodeAmount = False
            Me.txtDate1.Value = ""
            Me.txtTime1.Value = ""
            Me.txtTime2.Value = ""
            Me.Controls("lblInboundWt1").Caption = "Inbound Wt:"
            Me.Controls("lblTrk1").Caption = "Truck ID: ------"
            Me.Controls("lblTrk2").Caption = "Truck ID: ------"
            Me.txtInboundWeight.Value = ""
            Me.txtOutboundWeight.Value = ""
            Me.txtNetWeightLbs.Value = ""
            Me.txtNetWeightTons.Value = ""
            Me.txtWeighmaster.Value = ""
            Me.txtBy.Value = "Jesus Polanco"
            On Error Resume Next
            Me.txtInvoiceNum.Value = ""
            On Error GoTo 0
            
            On Error Resume Next
            Me.Controls("chkWeightCert").Value = False
            Me.Controls("chkRecyclingCert").Value = False
            On Error GoTo 0
        End Sub

        ' ========================================
        ' SAVE / LOAD TICKET STATE (in memory)
        ' ========================================

        Private Sub SaveCurrentState()
            If currentTicket < 1 Or currentTicket > ticketCount Then Exit Sub
            
            ticketData(currentTicket, 1) = Me.Controls("txtWeighedAt").Caption
            ticketData(currentTicket, 2) = Me.txtTicketNumber.Value
            ticketData(currentTicket, 3) = Me.txtWeighedForSeller.Value
            ticketData(currentTicket, 4) = Me.Controls("txtCompanyNameHauler").Value
            ticketData(currentTicket, 5) = Me.Controls("txtWasteOrigin").Value
            ticketData(currentTicket, 19) = Me.txtWasteData.Value
            On Error Resume Next
            ticketData(currentTicket, 22) = Me.Controls("cmbCity").Value
            On Error GoTo 0
            ticketData(currentTicket, 6) = Me.Controls("txtProductCode").Value
            ticketData(currentTicket, 7) = Me.txtTruckNumber.Value
            If Me.optCredit.Value Then
                ticketData(currentTicket, 10) = "CREDIT"
            ElseIf Me.optDebit.Value Then
                ticketData(currentTicket, 10) = "DEBIT"
            ElseIf Me.optAR.Value Then
                ticketData(currentTicket, 10) = "AR"
            ElseIf Me.optCash.Value Then
                ticketData(currentTicket, 10) = "CASH"
            ElseIf Me.optExpense.Value Then
                ticketData(currentTicket, 10) = "EXPENSE"
            Else
                ticketData(currentTicket, 10) = ""
            End If
            ticketData(currentTicket, 11) = Me.txtPaidAmount.Value
            ticketData(currentTicket, 12) = Me.txtDate1.Value
            ticketData(currentTicket, 13) = Me.txtTime1.Value
            ticketData(currentTicket, 14) = Me.txtTime2.Value
            ticketData(currentTicket, 15) = Me.txtInboundWeight.Value
            ticketData(currentTicket, 16) = Me.txtOutboundWeight.Value
            ticketData(currentTicket, 17) = Me.txtWeighmaster.Value
            ticketData(currentTicket, 18) = Me.txtBy.Value
            
            On Error Resume Next
            If CBool(Me.Controls("chkWeightCert").Value) Then
                ticketData(currentTicket, 20) = "1"
            Else
                ticketData(currentTicket, 20) = "0"
            End If
            If CBool(Me.Controls("chkRecyclingCert").Value) Then
                ticketData(currentTicket, 21) = "1"
            Else
                ticketData(currentTicket, 21) = "0"
            End If
            On Error GoTo 0
        End Sub

        Private Sub LoadTicketState(idx As Long)
            If idx < 1 Or idx > ticketCount Then Exit Sub
            
            isInitializing = True
            
            Me.Controls("txtWeighedAt").Caption = ticketData(idx, 1)
            Me.txtTicketNumber.Value = ticketData(idx, 2)
            Me.txtWeighedForSeller.Value = ticketData(idx, 3)
            Me.Controls("txtCompanyNameHauler").Value = ticketData(idx, 4)
            Me.Controls("txtWasteOrigin").Value = ticketData(idx, 5)
            Me.txtWasteData.Value = ticketData(idx, 19)
            On Error Resume Next
            Me.Controls("cmbCity").Value = ticketData(idx, 22)
            On Error GoTo 0
            Me.Controls("txtProductCode").Value = ticketData(idx, 6)
            Me.txtTruckNumber.Value = ticketData(idx, 7)
            
            Me.optCredit.Value = False
            Me.optDebit.Value = False
            Me.optAR.Value = False
            Me.optCash.Value = False
            Me.optExpense.Value = False
            Select Case ticketData(idx, 10)
                Case "CREDIT": Me.optCredit.Value = True
                Case "DEBIT": Me.optDebit.Value = True
                Case "AR": Me.optAR.Value = True
                Case "CASH": Me.optCash.Value = True
                Case "EXPENSE": Me.optExpense.Value = True
            End Select
            
            Me.txtPaidAmount.Value = ticketData(idx, 11)
            Me.txtDate1.Value = ticketData(idx, 12)
            Me.txtTime1.Value = FormatAsTime(ticketData(idx, 13))
            Me.txtTime2.Value = FormatAsTime(ticketData(idx, 14))
            Me.txtInboundWeight.Value = ticketData(idx, 15)
            Me.txtOutboundWeight.Value = ticketData(idx, 16)
            Me.txtWeighmaster.Value = ticketData(idx, 17)
            Me.txtBy.Value = ticketData(idx, 18)
            
            On Error Resume Next
            If ticketData(idx, 20) = "1" Then
                Me.Controls("chkWeightCert").Value = True
            ElseIf ticketData(idx, 20) = "0" Then
                Me.Controls("chkWeightCert").Value = False
            End If
            
            If ticketData(idx, 21) = "1" Then
                Me.Controls("chkRecyclingCert").Value = True
            ElseIf ticketData(idx, 21) = "0" Then
                Me.Controls("chkRecyclingCert").Value = False
            End If
            On Error GoTo 0
            
            ' Refresh labels
            If Me.txtTruckNumber.Value <> "" Then
                Me.Controls("lblTrk1").Caption = "Truck ID: " & Me.txtTruckNumber.Value
                Me.Controls("lblTrk2").Caption = "Truck ID: " & Me.txtTruckNumber.Value
            Else
                Me.Controls("lblTrk1").Caption = "Truck ID: ------"
                Me.Controls("lblTrk2").Caption = "Truck ID: ------"
            End If
            
            If IsNumeric(Me.txtInboundWeight.Value) And Me.txtInboundWeight.Value <> "" Then
                Me.Controls("lblInboundWt1").Caption = "Inbound Wt: " & Format(CDbl(Me.txtInboundWeight.Value), "#,##0") & " lb"
            Else
                Me.Controls("lblInboundWt1").Caption = "Inbound Wt: 0 lb"
            End If
            
            isInitializing = False
            CalculateNetWeight
        End Sub

        Private Sub UpdateTicketCounter()
            Dim tickets As Variant
            tickets = GetTicketList()
            
            Dim rowInfo As String
            rowInfo = GetTargetTableRow()
            
            If IsEmpty(tickets) Then
                Me.Controls("lblTicketCounter").Caption = "NEW  ·  " & rowInfo
            Else
                ' Find current ticket's position
                Dim pos As Long, i As Long
                pos = 0
                For i = 1 To UBound(tickets)
                    If Trim(tickets(i)) = Trim(Me.txtTicketNumber.Value) Then
                        pos = i
                        Exit For
                    End If
                Next i
                
                If pos > 0 Then
                    Me.Controls("lblTicketCounter").Caption = pos & " / " & UBound(tickets) & "  ·  " & rowInfo
                Else
                    Me.Controls("lblTicketCounter").Caption = "NEW  ·  " & rowInfo
                End If
            End If
        End Sub

        ' Returns a display string showing the actual Excel worksheet row number
        ' where this ticket will be written — not a table-relative index.
        '
        ' tbl.DataBodyRange.Row = the worksheet row of the FIRST data row in the table
        ' (i.e. the row directly below the header). All offsets are built from that.
        '
        ' Existing ticket  -> "Sheet row 312"       (found in DataBodyRange at index r)
        ' New ticket       -> "Sheet row 312 (new)" (will be appended after last row)
        Private Function GetTargetTableRow() As String
            Dim ws As Worksheet
            Dim tbl As ListObject
            Dim col As ListColumn
            Dim ticketNum As String
            
            ticketNum = Trim(Me.txtTicketNumber.Value)
            If ticketNum = "" Then
                GetTargetTableRow = "Sheet row -"
                Exit Function
            End If
            
            If Not TryGetUserFormTicketTable(ws, tbl) Then
                GetTargetTableRow = "Sheet row -"
                Exit Function
            End If

            On Error Resume Next
            Set col = tbl.ListColumns("Weighmaster")
            On Error GoTo 0
            
            ' firstDataRow = the actual Excel row number of row 1 of the table body
            Dim firstDataRow As Long
            firstDataRow = tbl.HeaderRowRange.row + 1   ' Works even when the table body is empty
            
            If col Is Nothing Or tbl.ListRows.count = 0 Then
                ' Table is empty — new row will land on firstDataRow
                GetTargetTableRow = "Sheet row " & firstDataRow & " (new)"
                Exit Function
            End If
            
            ' Scan for existing match
            Dim R As Long
            For R = 1 To col.DataBodyRange.Rows.count
                If Trim(CStr(col.DataBodyRange.Cells(R, 1).Value2)) = ticketNum Then
                    ' Actual sheet row = first data row + (r - 1)
                    GetTargetTableRow = "Sheet row " & (firstDataRow + R - 1)
                    Exit Function
                End If
            Next R
            
            ' Not found — ListRows.Add appends after the last existing row
            ' = firstDataRow + current count
            GetTargetTableRow = "Sheet row " & (firstDataRow + tbl.ListRows.count) & " (new)"
        End Function

        ' ========================================
        ' NAVIGATION - Browse existing table entries
        ' ========================================

        Private Sub cmdPrev_Click()
            Dim tickets As Variant
            tickets = GetTicketList()
            If IsEmpty(tickets) Then Exit Sub
            
            ' Find current position
            Dim pos As Long, i As Long
            pos = 0
            For i = 1 To UBound(tickets)
                If Trim(tickets(i)) = Trim(Me.txtTicketNumber.Value) Then
                    pos = i
                    Exit For
                End If
            Next i
            
            If pos > 1 Then
                LoadTableTicket tickets(pos - 1)
            ElseIf pos = 0 And UBound(tickets) >= 1 Then
                ' Currently on a new ticket, go to last existing
                LoadTableTicket tickets(UBound(tickets))
            End If
        End Sub

        Private Sub cmdNext_Click()
            Dim tickets As Variant
            tickets = GetTicketList()
            If IsEmpty(tickets) Then Exit Sub
            
            ' Find current position
            Dim pos As Long, i As Long
            pos = 0
            For i = 1 To UBound(tickets)
                If Trim(tickets(i)) = Trim(Me.txtTicketNumber.Value) Then
                    pos = i
                    Exit For
                End If
            Next i
            
            If pos > 0 And pos < UBound(tickets) Then
                LoadTableTicket tickets(pos + 1)
            End If
        End Sub

        Private Function FormatAsTime(ByVal val As Variant) As String
            If Trim(CStr(val)) = "" Then
                FormatAsTime = ""
            ElseIf IsNumeric(val) Or IsDate(val) Then
                On Error Resume Next
                FormatAsTime = Format(CDate(val), "hh:mm AM/PM")
                On Error GoTo 0
            Else
                FormatAsTime = CStr(val)
            End If
        End Function

        Private Sub LoadTableTicket(ByVal ticketNum As String)
            Dim data As Object
            Set data = LoadTicketFromTable(ticketNum)
            If data Is Nothing Then Exit Sub
            
            isInitializing = True
            ClearFormFields
            
            Me.txtTicketNumber.Value = ticketNum
            If data.Exists("CompanyNameHauler") Then Me.Controls("txtCompanyNameHauler").Value = data("CompanyNameHauler")
            ' City and address are now stored in separate columns — load them directly
            On Error Resume Next
            If data.Exists("City") Then Me.Controls("cmbCity").Value = data("City")
            On Error GoTo 0
            If data.Exists("WasteOrigin") Then Me.Controls("txtWasteOrigin").Value = data("WasteOrigin")
            If data.Exists("WasteData") Then Me.txtWasteData.Value = data("WasteData")
            If data.Exists("ProductCode") Then Me.Controls("txtProductCode").Value = data("ProductCode")
            If data.Exists("TruckNumber") Then Me.txtTruckNumber.Value = data("TruckNumber")
            If data.Exists("WeighedForSeller") Then Me.txtWeighedForSeller.Value = data("WeighedForSeller")
            If data.Exists("InvoiceNum") Then Me.txtInvoiceNum.Value = data("InvoiceNum")
            If data.Exists("ByField") And data("ByField") <> "" Then Me.txtBy.Value = data("ByField")
            If data.Exists("Date1") Then
                Dim d_val As Variant
                d_val = data("Date1")
                If IsNumeric(d_val) Or IsDate(d_val) Then
                    On Error Resume Next
                    Me.txtDate1.Value = Format(CDate(d_val), "mm/dd/yyyy")
                    On Error GoTo 0
                Else
                    Me.txtDate1.Value = d_val
                End If
            End If
            If data.Exists("Time1") Then Me.txtTime1.Value = FormatAsTime(data("Time1"))
            If data.Exists("Time2") Then Me.txtTime2.Value = FormatAsTime(data("Time2"))
            
            If data.Exists("InboundWeight") And data("InboundWeight") <> "0" And data("InboundWeight") <> "" Then
                Me.txtInboundWeight.Value = data("InboundWeight")
            End If
            If data.Exists("OutboundWeight") And data("OutboundWeight") <> "0" And data("OutboundWeight") <> "" Then
                Me.txtOutboundWeight.Value = data("OutboundWeight")
            End If
            If data.Exists("Tonnage") And data("Tonnage") <> "" And data("Tonnage") <> "0" Then
                Me.txtNetWeightTons.Value = Format(CDbl(data("Tonnage")), "0.00")
            End If
            
            Me.optCredit.Value = False: Me.optDebit.Value = False
            Me.optAR.Value = False: Me.optCash.Value = False
            If data.Exists("PaymentMethod") Then
                Select Case UCase(Trim(data("PaymentMethod")))
                    Case "CREDIT":  Me.optCredit.Value = True
                    Case "DEBIT":   Me.optDebit.Value = True
                    Case "AR":      Me.optAR.Value = True
                    Case "CASH":    Me.optCash.Value = True
                    Case "EXPENSE": Me.optExpense.Value = True
                End Select
            End If
            
            ' Restore saved amount and treat it as a user override so it
            ' isn't recalculated away when other fields are touched.
            If data.Exists("PaidAmount") And data("PaidAmount") <> "0" And data("PaidAmount") <> "" Then
                On Error Resume Next
                Me.txtPaidAmount.Value = Format(CDbl(data("PaidAmount")), "#,##0.00")
                On Error GoTo 0
                userOverrodeAmount = True
            End If
            
            If Me.txtTruckNumber.Value <> "" Then
                Me.Controls("lblTrk1").Caption = "Truck ID: " & Me.txtTruckNumber.Value
                Me.Controls("lblTrk2").Caption = "Truck ID: " & Me.txtTruckNumber.Value
            End If
            If IsNumeric(Me.txtInboundWeight.Value) And Me.txtInboundWeight.Value <> "" Then
                Me.Controls("lblInboundWt1").Caption = "Inbound Wt: " & Format(CDbl(Me.txtInboundWeight.Value), "#,##0") & " lb"
            End If
            
            CalculateNetWeight
            isInitializing = False
            UpdateTicketCounter
        End Sub

        Private Sub cmdNewTicket_Click()
            ClearFormFields
            Me.txtTicketNumber.Value = GetNextTicketNumber()
            Me.txtDate1.Value = Format(Date, "mm/dd/yyyy")
            Me.txtWeighmaster.Value = strWeighmaster
            Me.txtBy.Value = "Jesus Polanco"
            UpdateTicketCounter
            Me.txtWeighedForSeller.SetFocus
        End Sub

        ' ========================================
        ' AUTO-UPDATE EVENTS
        ' ========================================

        Private Sub txtTicketNumber_AfterUpdate()
            If isInitializing Then Exit Sub
            If Len(Trim(Me.txtTicketNumber.Value)) = 0 Then Exit Sub
            
            ' Try to load existing ticket from table
            Dim data As Object
            Set data = LoadTicketFromTable(Me.txtTicketNumber.Value)
            
            If data Is Nothing Then Exit Sub  ' Not found = new ticket, do nothing
            
            ' Ticket exists — populate form with its data
            isInitializing = True
            
            If data.Exists("CompanyNameHauler") Then Me.Controls("txtCompanyNameHauler").Value = data("CompanyNameHauler")
            ' City and address are now stored in separate columns — load them directly
            On Error Resume Next
            If data.Exists("City") Then Me.Controls("cmbCity").Value = data("City")
            On Error GoTo 0
            If data.Exists("WasteOrigin") Then Me.Controls("txtWasteOrigin").Value = data("WasteOrigin")
            If data.Exists("WasteData") Then Me.txtWasteData.Value = data("WasteData")
            If data.Exists("ProductCode") Then Me.Controls("txtProductCode").Value = data("ProductCode")
            If data.Exists("TruckNumber") Then Me.txtTruckNumber.Value = data("TruckNumber")
            If data.Exists("WeighedForSeller") Then Me.txtWeighedForSeller.Value = data("WeighedForSeller")
            If data.Exists("InvoiceNum") Then Me.txtInvoiceNum.Value = data("InvoiceNum")
            If data.Exists("ByField") And data("ByField") <> "" Then Me.txtBy.Value = data("ByField")
            If data.Exists("Date1") Then
                Dim d_auto As Variant
                d_auto = data("Date1")
                If IsNumeric(d_auto) Or IsDate(d_auto) Then
                    On Error Resume Next
                    Me.txtDate1.Value = Format(CDate(d_auto), "mm/dd/yyyy")
                    On Error GoTo 0
                Else
                    Me.txtDate1.Value = d_auto
                End If
            End If
            If data.Exists("Time1") Then Me.txtTime1.Value = FormatAsTime(data("Time1"))
            If data.Exists("Time2") Then Me.txtTime2.Value = FormatAsTime(data("Time2"))
            
            If data.Exists("InboundWeight") And data("InboundWeight") <> "0" And data("InboundWeight") <> "" Then
                Me.txtInboundWeight.Value = data("InboundWeight")
            End If
            If data.Exists("OutboundWeight") And data("OutboundWeight") <> "0" And data("OutboundWeight") <> "" Then
                Me.txtOutboundWeight.Value = data("OutboundWeight")
            End If
            If data.Exists("Tonnage") And data("Tonnage") <> "" And data("Tonnage") <> "0" Then
                Me.txtNetWeightTons.Value = Format(CDbl(data("Tonnage")), "0.00")
            End If
            
            ' Payment method
            Me.optCredit.Value = False: Me.optDebit.Value = False
            Me.optAR.Value = False: Me.optCash.Value = False
            If data.Exists("PaymentMethod") Then
                Select Case UCase(Trim(data("PaymentMethod")))
                    Case "CREDIT":  Me.optCredit.Value = True
                    Case "DEBIT":   Me.optDebit.Value = True
                    Case "AR":      Me.optAR.Value = True
                    Case "CASH":    Me.optCash.Value = True
                    Case "EXPENSE": Me.optExpense.Value = True
                End Select
            End If
            
            ' Restore saved amount
            If data.Exists("PaidAmount") And data("PaidAmount") <> "0" And data("PaidAmount") <> "" Then
                On Error Resume Next
                Me.txtPaidAmount.Value = Format(CDbl(data("PaidAmount")), "#,##0.00")
                On Error GoTo 0
            End If
            
            ' Update labels
            If Me.txtTruckNumber.Value <> "" Then
                Me.Controls("lblTrk1").Caption = "Truck ID: " & Me.txtTruckNumber.Value
                Me.Controls("lblTrk2").Caption = "Truck ID: " & Me.txtTruckNumber.Value
            End If
            If IsNumeric(Me.txtInboundWeight.Value) And Me.txtInboundWeight.Value <> "" Then
                Me.Controls("lblInboundWt1").Caption = "Inbound Wt: " & Format(CDbl(Me.txtInboundWeight.Value), "#,##0") & " lb"
            End If
            
            CalculateNetWeight
            isInitializing = False
        End Sub

        Private Sub txtTruckNumber_Change()
            If isInitializing Then Exit Sub
            If Me.txtTruckNumber.Value <> "" Then
                Me.Controls("lblTrk1").Caption = "Truck ID: " & Me.txtTruckNumber.Value
                Me.Controls("lblTrk2").Caption = "Truck ID: " & Me.txtTruckNumber.Value
                
                ' Auto-fill tare (outbound) weight from Tonage table
                Dim tare As Double
                tare = LookupTareWeight(Me.txtTruckNumber.Value)
                If tare > 0 Then
                    isInitializing = True
                    Me.txtOutboundWeight.Value = CStr(tare)
                    isInitializing = False
                    If Me.txtTime2.Value = "" Then
                        Me.txtTime2.Value = Format(Now, "hh:mm AM/PM")
                    End If
                    CalculateNetWeight
                End If
            Else
                Me.Controls("lblTrk1").Caption = "Truck ID: ------"
                Me.Controls("lblTrk2").Caption = "Truck ID: ------"
            End If
        End Sub

        Private Sub txtInboundWeight_Change()
            If isInitializing Then Exit Sub
            
            If Me.txtInboundWeight.Value <> "" And Me.txtTime1.Value = "" Then
                Me.txtTime1.Value = Format(Now, "hh:mm AM/PM")
                If Me.txtDate1.Value = "" Then
                    Me.txtDate1.Value = Format(Date, "mm/dd/yyyy")
                End If
            End If
            
            If IsNumeric(Me.txtInboundWeight.Value) And Me.txtInboundWeight.Value <> "" Then
                Me.Controls("lblInboundWt1").Caption = "Inbound Wt: " & Format(CDbl(Me.txtInboundWeight.Value), "#,##0") & " lb"
            Else
                Me.Controls("lblInboundWt1").Caption = "Inbound Wt:"
            End If
            
            CalculateNetWeight
        End Sub

        Private Sub txtOutboundWeight_Change()
            If isInitializing Then Exit Sub
            CalculateNetWeight
        End Sub

        Private Sub txtOutboundWeight_AfterUpdate()
            If isInitializing Then Exit Sub
            If Me.txtOutboundWeight.Value <> "" And Me.txtTime2.Value = "" Then
                Me.txtTime2.Value = Format(Now, "hh:mm AM/PM")
            End If
        End Sub

        ' ========================================
        ' NET WEIGHT = |Inbound - Outbound|, min 1 ton
        ' ========================================

        Private Sub CalculateNetWeight()
            Dim inbound As Double
            Dim outbound As Double
            Dim netLbs As Double
            Dim netTons As Double
            
            inbound = 0
            outbound = 0
            
            If IsNumeric(Me.txtInboundWeight.Value) And Me.txtInboundWeight.Value <> "" Then
                inbound = CDbl(Me.txtInboundWeight.Value)
            End If
            If IsNumeric(Me.txtOutboundWeight.Value) And Me.txtOutboundWeight.Value <> "" Then
                outbound = CDbl(Me.txtOutboundWeight.Value)
            End If
            
            ' Require BOTH inputs to calculate net
            If inbound > 0 And outbound > 0 Then
                netLbs = Abs(inbound - outbound)
                netTons = netLbs / 2000
                ' Business rule: minimum billable weight is 1 ton
                If netTons < 1 And netTons > 0 Then netTons = 1
                
                Me.txtNetWeightLbs.Value = Format(netLbs, "#,##0") & " lb"
                Me.txtNetWeightTons.Value = Format(netTons, "0.00")
                
                ' Both weights provided, refresh price
                RefreshAmount
            Else
                ' Missing one or both weights, clear the results ONLY if actively editing
                If Not isInitializing Then
                    Me.txtNetWeightLbs.Value = ""
                    Me.txtNetWeightTons.Value = ""
                    
                    ' Also safely clear the calculated amount without running the expensive lookup
                    isInitializing = True
                    Me.txtPaidAmount.Value = ""
                    isInitializing = False
                End If
            End If
        End Sub

        Private Sub RefreshAmount()
            If isInitializing Then Exit Sub
            ' If the user has manually typed a value into the Amount box,
            ' don't overwrite it. Auto-payment-method logic still runs below.
            If userOverrodeAmount Then
                Dim acctOv As String
                acctOv = Me.Controls("txtCompanyNameHauler").Value
                Dim autoPaymentOv As String
                autoPaymentOv = DeterminePaymentMethod(acctOv, Me.txtPaidAmount.Value)
                isInitializing = True
                If autoPaymentOv = "EXPENSE" Then
                    Me.optCredit.Value = False
                    Me.optDebit.Value = False
                    Me.optCash.Value = False
                    Me.optAR.Value = False
                    Me.optExpense.Value = True
                ElseIf autoPaymentOv = "AR" Then
                    If Not (Me.optCredit.Value Or Me.optDebit.Value Or Me.optCash.Value Or Me.optExpense.Value) Then
                        Me.optAR.Value = True
                    End If
                End If
                isInitializing = False
                Exit Sub
            End If

            Dim prod As String, acct As String
            Dim ticketDate As Variant, tons As Double
            
            prod = Me.Controls("txtProductCode").Value
            acct = Me.Controls("txtCompanyNameHauler").Value
            ticketDate = Me.txtDate1.Value
            
            ' Get tonnage from the calculated net weight
            tons = 0
            If IsNumeric(Me.txtNetWeightTons.Value) Then
                tons = CDbl(Me.txtNetWeightTons.Value)
            End If
            

            ' Only calculate if we have enough inputs
            If Len(Trim(prod)) = 0 Then
                Me.txtPaidAmount.Value = ""
                Exit Sub
            End If
            
            Dim result As Variant
            result = CalculateAmountFromPricing(prod, acct, ticketDate, tons)
            
            If IsEmpty(result) Or result = "" Then
                Me.txtPaidAmount.Value = ""
            Else
                ' --- APPLY CREDIT SURCHARGE IF OPTCREDIT IS SELECTED ---
                Dim finalAmount As Double
                finalAmount = CDbl(result)
                If Me.optCredit.Value = True Then
                    finalAmount = finalAmount * 1.023
                End If
                Me.txtPaidAmount.Value = Format(finalAmount, "#,##0.00")
            End If
            
            ' Auto-select payment method based on amount + account (replicates PAYMENT column formula)
            Dim autoPayment As String
            autoPayment = DeterminePaymentMethod(acct, Me.txtPaidAmount.Value)
            
            ' Guard against re-entrant click events when setting radio buttons
            isInitializing = True
            
            If autoPayment = "EXPENSE" Then
                Me.optCredit.Value = False
                Me.optDebit.Value = False
                Me.optCash.Value = False
                Me.optAR.Value = False
                Me.optExpense.Value = True
            ElseIf autoPayment = "AR" Then
                ' Respect manual choices for Credit/Debit/Cash/Expense. Otherwise default to AR.
                If Not (Me.optCredit.Value Or Me.optDebit.Value Or Me.optCash.Value Or Me.optExpense.Value) Then
                    Me.optAR.Value = True
                End If
            Else
                ' Keep all manual choices (including AR and Expense) intact if they don't apply automatically.
                ' This allows manual clicks for AR and loaded ticket states to persist.
            End If
            
            isInitializing = False
        End Sub

        ' ========================================
        ' PAYMENT METHOD LIVE UPDATES
        ' Re-calculate amount (e.g., adding 2.3% fee) when user manually clicks Credit
        ' ========================================
        Private Sub optCredit_Click()
            If Not isInitializing Then RefreshAmount
        End Sub

        Private Sub optDebit_Click()
            If Not isInitializing Then RefreshAmount
        End Sub

        Private Sub optCash_Click()
            If Not isInitializing Then RefreshAmount
        End Sub

        Private Sub optAR_Click()
            If Not isInitializing Then RefreshAmount
        End Sub

        Private Sub optExpense_Click()
            If Not isInitializing Then RefreshAmount
        End Sub

        ' Mark the Amount as user-overridden whenever the user types into it,
        ' so RefreshAmount won't overwrite their value when other fields change.
        Private Sub txtPaidAmount_Change()
            If isInitializing Then Exit Sub
            userOverrodeAmount = (Len(Trim(Me.txtPaidAmount.Value)) > 0)
        End Sub

        ' Fires only when user finishes typing and moves away from the field
        Private Sub txtProductCode_AfterUpdate()
            If isInitializing Then Exit Sub
            RefreshAmount
        End Sub

        Private Sub txtCompanyNameHauler_AfterUpdate()
            If isInitializing Then Exit Sub
            RefreshAmount
        End Sub

        ' ========================================
        ' SAVE TICKET (to Excel + remove from open list)
        ' ========================================

        Private Sub cmdSaveTicket_Click()
            Dim paymentMethod As String
            
            If Me.txtTicketNumber.Value = "" Then
                MsgBox "Weighmaster Number is required.", vbExclamation, "Required Field"
                Exit Sub
            End If
            If Me.Controls("txtCompanyNameHauler").Value = "" Then
                MsgBox "Company Name/Hauler is required.", vbExclamation, "Required Field"
                Me.Controls("txtCompanyNameHauler").SetFocus
                Exit Sub
            End If
            
            ' Waste Data required for container hauls
            Dim prodVal As String
            prodVal = UCase(Trim(Me.Controls("txtProductCode").Value))
            If (prodVal = "40YD" Or prodVal = "10YD") And Trim(Me.txtWasteData.Value) = "" Then
                MsgBox "Waste Data is required for " & prodVal & " material type.", vbExclamation, "Required Field"
                Me.txtWasteData.SetFocus
                Exit Sub
            End If
            
            ' Freeze screen for the entire save + form-reset sequence
            isInitializing = True
            Application.ScreenUpdating = False
            Application.EnableEvents = False
            Application.Calculation = xlCalculationManual
            
            paymentMethod = ""
            If Me.optCredit.Value Then paymentMethod = "CREDIT"
            If Me.optDebit.Value Then paymentMethod = "DEBIT"
            If Me.optAR.Value Then paymentMethod = "AR"
            If Me.optCash.Value Then paymentMethod = "CASH"
            If Me.optExpense.Value Then paymentMethod = "EXPENSE"
            
            ' Pre-stage args to stay under VBA's 25 line-continuation limit
            Dim sTicket As String, sWeighedAt As String, sSeller As String
            Dim sHauler As String, sOrigin As String, sProd As String
            Dim sTruck As String, sPaid As String
            Dim sDate1 As String, sTime1 As String, sTime2 As String
            Dim sInWt As String, sOutWt As String, sNetLb As String
            Dim sNetTn As String, sWM As String, sByF As String, sWD As String
            
            sTicket = Me.txtTicketNumber.Value
            sWeighedAt = Me.Controls("txtWeighedAt").Caption
            sSeller = Me.txtWeighedForSeller.Value
            sHauler = Me.Controls("txtCompanyNameHauler").Value
            ' Combine City + Address into ORIGIN column as "City, Address"
            Dim sCityVal As String, sAddrVal As String
            On Error Resume Next
            sCityVal = Trim(Me.Controls("cmbCity").Value)
            On Error GoTo 0
            sAddrVal = Trim(Me.Controls("txtWasteOrigin").Value)
            If sCityVal <> "" And sAddrVal <> "" Then
                sOrigin = sCityVal & ", " & sAddrVal
            ElseIf sCityVal <> "" Then
                sOrigin = sCityVal
            Else
                sOrigin = sAddrVal
            End If
            sProd = Me.Controls("txtProductCode").Value
            sTruck = Me.txtTruckNumber.Value
            sPaid = Me.txtPaidAmount.Value
            sDate1 = Me.txtDate1.Value
            sTime1 = Me.txtTime1.Value
            sTime2 = Me.txtTime2.Value
            sInWt = Me.txtInboundWeight.Value
            sOutWt = Me.txtOutboundWeight.Value
            sNetLb = Me.txtNetWeightLbs.Value
            sNetTn = Me.txtNetWeightTons.Value
            sWM = Me.txtWeighmaster.Value
            sByF = Me.txtBy.Value
            sWD = Me.txtWasteData.Value
            
            Call SaveWeighmasterCertificate( _
                sTicket, sWeighedAt, sSeller, sHauler, sAddrVal, _
                sProd, sTruck, "", "", paymentMethod, _
                sPaid, sDate1, sTime1, "", sTime2, _
                "", "", sOutWt, sInWt, sNetLb, _
                sNetTn, sWM, _
                wasteData:=sWD, _
                wasteCity:=sCityVal, _
                invoiceNumber:=Me.txtInvoiceNum.Value, _
                byField:=sByF)
            
            ' Restore before showing the success dialog and resetting form
            isInitializing = False
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            
            MsgBox "Weighmaster #" & Me.txtTicketNumber.Value & " saved!", vbInformation, "Success"
            RemoveTicket currentTicket
        End Sub

        Private Sub RemoveTicket(idx As Long)
            Dim i As Long
            Dim j As Long
            Dim newData() As String
            Dim dest As Long
            
            If ticketCount <= 1 Then
                ticketCount = 1
                currentTicket = 1
                ReDim ticketData(1 To 1, 1 To 22)
                ClearFormFields
                Me.txtTicketNumber.Value = GetNextTicketNumber()
                Me.txtDate1.Value = Format(Date, "mm/dd/yyyy")
                SaveCurrentState
                UpdateTicketCounter
                Exit Sub
            End If
            
            ReDim newData(1 To ticketCount - 1, 1 To 22)
            dest = 0
            For i = 1 To ticketCount
                If i <> idx Then
                    dest = dest + 1
                    For j = 1 To 22
                        newData(dest, j) = ticketData(i, j)
                    Next j
                End If
            Next i
            
            ticketCount = ticketCount - 1
            ticketData = newData
            
            If currentTicket > ticketCount Then currentTicket = ticketCount
            LoadTicketState currentTicket
            UpdateTicketCounter
        End Sub

        ' ========================================
        ' BUTTONS
        ' ========================================

        Private Sub cmdClear_Click()
            ClearFormFields
            Me.txtTicketNumber.Value = GetNextTicketNumber()
            Me.txtDate1.Value = Format(Date, "mm/dd/yyyy")
        End Sub

        Private Sub cmdClose_Click()
            Unload Me
        End Sub

        Private Sub cmdTime1_Click()
            Me.txtTime1.Value = Format(Now, "hh:mm AM/PM")
            If Me.txtDate1.Value = "" Then
                Me.txtDate1.Value = Format(Date, "mm/dd/yyyy")
            End If
        End Sub

        Private Sub cmdTime2_Click()
            Me.txtTime2.Value = Format(Now, "hh:mm AM/PM")
        End Sub

        Private Sub txtTime1_AfterUpdate()
            If isInitializing Then Exit Sub
            Me.txtTime1.Value = FormatAsTime(Me.txtTime1.Value)
        End Sub

        Private Sub txtTime2_AfterUpdate()
            If isInitializing Then Exit Sub
            Me.txtTime2.Value = FormatAsTime(Me.txtTime2.Value)
        End Sub

        ' ========================================
        ' PRINT TICKET
        ' ========================================

        ' -- Helper: get (or create once) the permanent hidden print sheet --
        ' KEY FIX: We NEVER Add/Delete sheets during printing — that mutates workbook
        ' structure and OneDrive detects it ? forces a sync mid-print.
        ' We also NEVER toggle visibility for direct PrintOut; Excel can print
        ' xlSheetVeryHidden sheets fine. Visibility is only toggled for PrintPreview.
        Private Function GetOrCreatePrintSheet() As Worksheet
            Const SHEET_NAME As String = "PrintTicket_Temp"
            Dim ws As Worksheet
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
            On Error GoTo 0
            
            If ws Is Nothing Then
                ' Safety fallback — should never happen if EnsurePrintSheetExists
                ' was called at form open. Create without changing AutoSave.
                Application.DisplayAlerts = False
                Dim wasAS As Boolean
                wasAS = ThisWorkbook.AutoSaveOn
                If wasAS Then ThisWorkbook.AutoSaveOn = False
                Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
                ws.Name = SHEET_NAME
                ws.Visible = xlSheetVeryHidden
                Application.DisplayAlerts = True
                ' Save the new sheet into the workbook so it persists
                On Error Resume Next
                Application.CommandBars.ExecuteMso "FileSave"
                DoEvents
                On Error GoTo 0
                If wasAS Then ThisWorkbook.AutoSaveOn = True
            End If
            
            ' Return the sheet — caller decides whether to show it
            Set GetOrCreatePrintSheet = ws
        End Function
        
        ' -- Call this once when the form opens to ensure the print sheet
        '    is already baked into the saved workbook (prevents Add() during print).
        Public Sub EnsurePrintSheetExists()
            Dim ws As Worksheet
            Set ws = GetOrCreatePrintSheet()   ' creates + saves if missing
            ' Make sure it stays very-hidden
            If ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden
        End Sub

        Private Sub cmdPrintTicket_Click()
            isInitializing = True
            Application.EnableEvents = False
            Application.Calculation = xlCalculationManual
            
            Dim wasPrintAutoSave As Boolean
            wasPrintAutoSave = ThisWorkbook.AutoSaveOn
            If wasPrintAutoSave Then ThisWorkbook.AutoSaveOn = False
            
            On Error GoTo PrintCleanup
            
            Dim wsPrint As Worksheet
            Set wsPrint = GetOrCreatePrintSheet()
            DrawCurrentTicketToSheet wsPrint
            
            ' PrintOut requires the sheet to be at least xlSheetVisible —
            ' it silently does nothing on xlSheetVeryHidden sheets.
            ' AutoSave is already off so OneDrive won't sync this momentary flip.
            Application.ScreenUpdating = False   ' keep flip invisible to user
            wsPrint.Visible = xlSheetVisible
            wsPrint.PrintOut Copies:=1, Collate:=True
            wsPrint.Visible = xlSheetVeryHidden
            
PrintCleanup:
            isInitializing = False
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            If wasPrintAutoSave Then ThisWorkbook.AutoSaveOn = True
            On Error Resume Next
            If Not Me.Visible Then Me.Show
            On Error GoTo 0
        End Sub

        ' -- Optional print preview (user explicitly requested) --
        ' NOTE: PrintPreview requires the sheet to be visible — but AutoSave is
        ' disabled before the flip so OneDrive does NOT sync the visibility change.
        Private Sub cmdPrintPreview_Click()
            isInitializing = True
            Application.EnableEvents = False
            Application.Calculation = xlCalculationManual
            
            ' Disable AutoSave BEFORE touching visibility — this prevents OneDrive
            ' from syncing the momentary xlSheetVisible state.
            Dim wasPreviewAutoSave As Boolean
            wasPreviewAutoSave = ThisWorkbook.AutoSaveOn
            If wasPreviewAutoSave Then ThisWorkbook.AutoSaveOn = False
            
            On Error GoTo PreviewCleanup
            
            Dim wsPrev As Worksheet
            Set wsPrev = GetOrCreatePrintSheet()
            DrawCurrentTicketToSheet wsPrev
            
            ' Show only for preview, then immediately re-hide
            wsPrev.Visible = xlSheetVisible
            Me.Hide
            wsPrev.PrintPreview EnableChanges:=False
            wsPrev.Visible = xlSheetVeryHidden
            
PreviewCleanup:
            isInitializing = False
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            If wasPreviewAutoSave Then ThisWorkbook.AutoSaveOn = True
            On Error Resume Next
            If Not Me.Visible Then Me.Show
            On Error GoTo 0
        End Sub

        ' ========================================
        ' SAVE & PRINT — saves to table then prints
        ' ========================================
        Private Sub cmdSaveAndPrint_Click()
            ' --- 1. Validate BEFORE touching anything ---
            If Me.txtTicketNumber.Value = "" Then
                MsgBox "Weighmaster Number is required.", vbExclamation, "Required Field"
                Exit Sub
            End If
            If Me.Controls("txtCompanyNameHauler").Value = "" Then
                MsgBox "Company Name/Hauler is required.", vbExclamation, "Required Field"
                Me.Controls("txtCompanyNameHauler").SetFocus
                Exit Sub
            End If
            Dim sprodVal As String
            sprodVal = UCase(Trim(Me.Controls("txtProductCode").Value))
            If (sprodVal = "40YD" Or sprodVal = "10YD") And Trim(Me.txtWasteData.Value) = "" Then
                MsgBox "Waste Data is required for " & sprodVal & " material type.", vbExclamation, "Required Field"
                Me.txtWasteData.SetFocus
                Exit Sub
            End If

            ' --- 2. Lock form events to prevent Change/AfterUpdate firing during save ---
            isInitializing = True
            Application.ScreenUpdating = False

            ' --- 3. Snapshot payment before print/save ---
            Dim spaymentMethod As String
            spaymentMethod = ""
            If Me.optCredit.Value Then spaymentMethod = "CREDIT"
            If Me.optDebit.Value Then spaymentMethod = "DEBIT"
            If Me.optAR.Value Then spaymentMethod = "AR"
            If Me.optCash.Value Then spaymentMethod = "CASH"
            If Me.optExpense.Value Then spaymentMethod = "EXPENSE"

            ' --- 4. Build combined City + Address origin (same logic as cmdSaveTicket_Click) ---
            Dim sSAPCity As String, sSAPAddr As String, sSAPOrigin As String
            On Error Resume Next
            sSAPCity = Trim(Me.Controls("cmbCity").Value)
            On Error GoTo 0
            sSAPAddr = Trim(Me.Controls("txtWasteOrigin").Value)
            If sSAPCity <> "" And sSAPAddr <> "" Then
                sSAPOrigin = sSAPCity & ", " & sSAPAddr
            ElseIf sSAPCity <> "" Then
                sSAPOrigin = sSAPCity
            Else
                sSAPOrigin = sSAPAddr
            End If

            Dim wasSAPAutoSave As Boolean
            Dim wsSAP As Worksheet
            Dim printErrText As String

            Application.EnableEvents = False
            Application.Calculation = xlCalculationManual
            wasSAPAutoSave = ThisWorkbook.AutoSaveOn
            If wasSAPAutoSave Then ThisWorkbook.AutoSaveOn = False

            On Error Resume Next
            Set wsSAP = GetOrCreatePrintSheet()
            If Err.Number = 0 Then DrawCurrentTicketToSheet wsSAP
            If Err.Number = 0 Then
                wsSAP.Visible = xlSheetVisible
                wsSAP.PrintOut Copies:=1, Collate:=True
                wsSAP.Visible = xlSheetVeryHidden
            End If
            If Err.Number <> 0 Then
                printErrText = Err.Description
                Err.Clear
                If Not wsSAP Is Nothing Then wsSAP.Visible = xlSheetVeryHidden
            End If
            On Error GoTo 0

            ' --- 5. Save to table — pass keepLocked:=True so ScreenUpdating stays
            '        frozen from here through the print; no flicker gap between the two.
            Call SaveWeighmasterCertificate( _
                Me.txtTicketNumber.Value, _
                Me.Controls("txtWeighedAt").Caption, _
                Me.txtWeighedForSeller.Value, _
                Me.Controls("txtCompanyNameHauler").Value, _
                sSAPAddr, _
                Me.Controls("txtProductCode").Value, _
                Me.txtTruckNumber.Value, _
                "", "", spaymentMethod, _
                Me.txtPaidAmount.Value, _
                Me.txtDate1.Value, Me.txtTime1.Value, "", Me.txtTime2.Value, _
                "", "", Me.txtOutboundWeight.Value, Me.txtInboundWeight.Value, _
                Me.txtNetWeightLbs.Value, Me.txtNetWeightTons.Value, _
                Me.txtWeighmaster.Value, _
                wasteData:=Me.txtWasteData.Value, _
                wasteCity:=sSAPCity, _
                invoiceNumber:=Me.txtInvoiceNum.Value, _
                byField:=Me.txtBy.Value, _
                keepLocked:=True)
            GoTo SAPCleanupAfterPrint

            ' --- 5. Screen is still frozen (keepLocked=True above). Disable AutoSave
            '        before printing to prevent OneDrive sync mid-print.
            Application.EnableEvents = False
            Application.Calculation = xlCalculationManual
            wasSAPAutoSave = ThisWorkbook.AutoSaveOn
            If wasSAPAutoSave Then ThisWorkbook.AutoSaveOn = False

            ' --- 6. Print directly — no preview dialog, goes straight to printer ---
            ' PrintOut requires xlSheetVisible; AutoSave is already off so the
            ' visibility flip won't trigger OneDrive sync. Screen stays frozen.
            Set wsSAP = GetOrCreatePrintSheet()
            DrawCurrentTicketToSheet wsSAP
            wsSAP.Visible = xlSheetVisible
            wsSAP.PrintOut Copies:=1, Collate:=True
            wsSAP.Visible = xlSheetVeryHidden

SAPCleanupAfterPrint:
SAPCleanup:
            ' --- 7. Restore everything once print is done ---
            isInitializing = False
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            On Error Resume Next
            If wasSAPAutoSave Then ThisWorkbook.AutoSaveOn = True
            If Not Me.Visible Then Me.Show
            On Error GoTo 0
            If printErrText <> "" Then
                MsgBox "The ticket save still ran, but printing hit an issue:" & vbCrLf & vbCrLf & _
                       printErrText, vbExclamation, "Print Warning"
            End If
        End Sub

        Public Sub DrawCurrentTicketToSheet(wsPrint As Worksheet)

            ' 1. SNAPSHOT — capture form values exactly as displayed
            Dim sTicket    As String, sWeighedAt As String, sSeller As String
            Dim sHauler    As String, sOrigin    As String, sWasteData As String
            Dim sProd      As String, sTruck     As String
            Dim sPayment   As String, sPaid      As String, sDate1    As String
            Dim sTime1     As String, sTime2     As String, sInWt     As String
            Dim sOutWt     As String, sNetLb     As String, sNetTn    As String
            Dim sWM        As String, sByF       As String
            ' Logo-related declarations (hoisted from loop to avoid re-declaration error)
            Dim assetWs2   As Worksheet
            Dim sourceShp  As Shape
            Dim tempShp    As Shape
            Dim pastedShp  As Shape
            Dim wasHidden  As Boolean
            Dim lgRng      As Range
            Dim targetW    As Double, targetH As Double

            sTicket = Me.txtTicketNumber.Value
            sWeighedAt = Me.Controls("txtWeighedAt").Caption
            If sWeighedAt = "" Then sWeighedAt = "American Industrial Services, Inc."
            sSeller = Me.txtWeighedForSeller.Value
            sHauler = Me.Controls("txtCompanyNameHauler").Value
            sOrigin = Me.Controls("txtWasteOrigin").Value
            sWasteData = Me.txtWasteData.Value
            sProd = Me.Controls("txtProductCode").Value
            
            ' Replace "MIXED" with "Mixed C&D" for printing only
            If UCase(Trim(sProd)) = "MIXED" Then
                sProd = "Mixed C&D"
            End If
            
            sTruck = Me.txtTruckNumber.Value
            sDate1 = Me.txtDate1.Value
            sTime1 = Me.txtTime1.Value
            sTime2 = Me.txtTime2.Value
            
            ' Safety fallback: format times if they are stuck as raw Excel numbers or dates
            If IsNumeric(sTime1) Or IsDate(sTime1) Then
                On Error Resume Next
                sTime1 = Format(CDate(sTime1), "hh:mm AM/PM")
                On Error GoTo 0
            End If
            If IsNumeric(sTime2) Or IsDate(sTime2) Then
                On Error Resume Next
                sTime2 = Format(CDate(sTime2), "hh:mm AM/PM")
                On Error GoTo 0
            End If
            sInWt = Me.txtInboundWeight.Value
            sOutWt = Me.txtOutboundWeight.Value
            sNetLb = Me.txtNetWeightLbs.Value
            sNetTn = Me.txtNetWeightTons.Value
            sWM = Me.txtWeighmaster.Value
            sByF = Me.txtBy.Value
            sPaid = Me.txtPaidAmount.Value
            
            Dim isWeightCert As Boolean
            Dim isRecyclingCert As Boolean
            isWeightCert = True ' default
            isRecyclingCert = False
            
            On Error Resume Next
            isWeightCert = CBool(Me.Controls("chkWeightCert").Value)
            isRecyclingCert = CBool(Me.Controls("chkRecyclingCert").Value)
            On Error GoTo 0
            
            Dim certHeader As String
            certHeader = ""
            If isWeightCert Then certHeader = "WEIGHMASTER CERTIFICATE"
            If isRecyclingCert Then
                If certHeader <> "" Then
                    certHeader = certHeader & " / RECYCLING C&D CERTIFICATE"
                Else
                    certHeader = "RECYCLING C&D CERTIFICATE"
                End If
            End If

            sPayment = ""
            If Me.optCredit.Value Then sPayment = "CREDIT"
            If Me.optDebit.Value Then sPayment = "DEBIT"
            If Me.optAR.Value Then sPayment = "AR"
            If Me.optCash.Value Then sPayment = "CASH"
            If Me.optExpense.Value Then sPayment = ""  ' Do not display EXPENSE on ticket
            
            ' If AR or EXPENSE is selected, hide prices on the printed ticket
            If Me.optAR.Value Or Me.optExpense.Value Then
                sPaid = ""
            End If

            ' Weight display — use form values directly, no recalculation
            Dim sInFmt  As String: sInFmt = sInWt
            Dim sOutFmt As String: sOutFmt = sOutWt
            Dim sNetFmt As String: sNetFmt = sNetLb

            ' 2. CLEAR THE PASSED SHEET

            ' Clear all previous content and shapes
            wsPrint.Cells.Clear
            Dim shp As Shape
            For Each shp In wsPrint.Shapes
                shp.Delete
            Next shp

            ' 3. TERMS / WM CERT strings
            Dim TERMS   As String
            Dim WMCERT  As String
            Dim ADDR1   As String
            Dim ADDR2   As String
            Dim PHONE   As String
            ADDR1 = "5626 CHERRY AVE."
            ADDR2 = "LONG BEACH, CA 90805"
            PHONE = "(562) 272-8060"
            TERMS = "This is to certify that the following described commodity was weighed, measured, or counted " & _
                "by a weighmaster, whose signature is on this certificate, who is a recognized authority of accuracy, " & _
                "as prescribed by Chapter 7 (commencing with Section 12700) of Division 5 of the California Business " & _
                "and Professions Code, administered by the Division of Measurement Standards of the California " & _
                "Department of Food and Agriculture. " & _
                "The undersigned individual signing this document on behalf of Customer acknowledges that he or she " & _
                "has read and understands the terms and conditions on the reverse side and that he or she has the " & _
                "authority to sign this document on behalf of the customer."
            WMCERT = "1. Unacceptable Waste. ""Unacceptable Waste"" means: (a) any material that is defined as " & _
                "hazardous material or hazardous waste, untreated medical or infectious waste, radioactive waste, " & _
                "or unprofiled sewage sludge under any applicable law, regulation or permit (""Applicable Law""); " & _
                "or (b) any other material that may present an endangerment to the facility, to its employees or " & _
                "to public health or safety. Customer shall not deliver any Unacceptable Waste. " & _
                "If Customer delivers Unacceptable Waste, Operator may, in its sole discretion: " & _
                "(x) reject the Unacceptable Waste; or, as Customer's agent, dispose of the Unacceptable Waste " & _
                "at a location authorized to accept it in accordance with all Applicable Laws and charge Customer " & _
                "all direct and indirect costs incurred for its handling, transportation and disposal."

            ' 4. COLUMN WIDTHS
            With wsPrint
                .Columns("A").ColumnWidth = 11
                .Columns("B").ColumnWidth = 20
                .Columns("C").ColumnWidth = 5
                .Columns("D").ColumnWidth = 5
                .Columns("E").ColumnWidth = 5
                .Columns("F").ColumnWidth = 9
                .Columns("G").ColumnWidth = 11
                .Columns("H").ColumnWidth = 5
                .Columns("I").ColumnWidth = 14
            End With

            ' 5. TWO COPIES
            Dim R   As Long
            Dim iCp As Integer
            For iCp = 1 To 2
                R = IIf(iCp = 1, 2, 24)

                wsPrint.Rows(R - 1).RowHeight = 14

                With wsPrint
                    ' -- ROW R: Company name | Logo | Ticket # --
                    .Rows(R).RowHeight = 36
                    With .Cells(R, 1): .Value = sWeighedAt: .Font.Bold = True: .Font.Size = 8: End With

                    ' --- Logo: load from hidden Assets sheet ---
                    Set assetWs2 = Nothing
                    On Error Resume Next
                    Set assetWs2 = ThisWorkbook.Worksheets("Assets")
                    On Error GoTo 0
                    
                    If Not assetWs2 Is Nothing Then
                        If assetWs2.Shapes.count > 0 Then
                            Set sourceShp = Nothing
                            For Each tempShp In assetWs2.Shapes
                                If tempShp.Type = msoPicture Or tempShp.Type = msoLinkedPicture Then
                                    If Not Intersect(tempShp.TopLeftCell, assetWs2.Range("A1")) Is Nothing Then
                                        Set sourceShp = tempShp
                                        Exit For
                                    End If
                                End If
                            Next tempShp
                            If sourceShp Is Nothing Then
                                For Each tempShp In assetWs2.Shapes
                                    If tempShp.Type = msoPicture Or tempShp.Type = msoLinkedPicture Then
                                        Set sourceShp = tempShp
                                        Exit For
                                    End If
                                Next tempShp
                            End If
                            If Not sourceShp Is Nothing Then
                                wasHidden = (assetWs2.Visible <> xlSheetVisible)
                                If wasHidden Then assetWs2.Visible = xlSheetVisible
                                sourceShp.CopyPicture xlScreen, xlPicture
                                If wasHidden Then assetWs2.Visible = xlSheetVeryHidden
                                DoEvents
                                On Error Resume Next
                                wsPrint.Paste Destination:=wsPrint.Range("C" & R)
                                On Error GoTo 0
                                If wsPrint.Shapes.count > 0 Then
                                    Set pastedShp = wsPrint.Shapes(wsPrint.Shapes.count)
                                    Set lgRng = .Range(.Cells(R, 3), .Cells(R + 2, 7))
                                    pastedShp.LockAspectRatio = msoTrue
                                    targetW = lgRng.Width - 4
                                    targetH = lgRng.Height - 4
                                    If (pastedShp.Width / pastedShp.Height) > (targetW / targetH) Then
                                        pastedShp.Width = targetW
                                    Else
                                        pastedShp.Height = targetH
                                    End If
                                    pastedShp.Left = lgRng.Left + 2
                                    pastedShp.Top = lgRng.Top + (lgRng.Height - pastedShp.Height) / 2
                                End If
                            End If
                        End If
                    End If

                    ' Ticket# — user's preferred: font 6, left-aligned, H:I
                    .Range(.Cells(R, 8), .Cells(R, 9)).Merge
                    With .Cells(R, 8)
                        .Value = "Ticket: " & sTicket
                        .Font.Bold = True
                        .Font.Size = 12
                        .HorizontalAlignment = xlLeft
                    End With

                    ' -- ROW R+1: Address 1 | Date --
                    .Rows(R + 1).RowHeight = 13
                    .Cells(R + 1, 1).Value = ADDR1
                    .Range(.Cells(R + 1, 8), .Cells(R + 1, 9)).Merge
                    With .Cells(R + 1, 8)
                        .Value = "Date: " & sDate1
                        .Font.Size = 8
                        .HorizontalAlignment = xlLeft
                    End With

                    ' -- ROW R+2: Address 2 | Time --
                    .Rows(R + 2).RowHeight = 13
                    .Cells(R + 2, 1).Value = ADDR2
                    .Range(.Cells(R + 2, 8), .Cells(R + 2, 9)).Merge
                    With .Cells(R + 2, 8)
                        .Value = "Time: " & sTime1 & IIf(sTime2 <> "", " - " & sTime2, "")
                        .Font.Size = 8
                        .HorizontalAlignment = xlLeft
                    End With

                    ' -- ROW R+3: Phone | "Scale" header --
                    .Rows(R + 3).RowHeight = 13
                    .Cells(R + 3, 1).Value = PHONE
                    .Range(.Cells(R + 3, 6), .Cells(R + 3, 9)).Merge
                    With .Range(.Cells(R + 3, 6), .Cells(R + 3, 6))
                        .Value = "Scale": .Font.Underline = xlUnderlineStyleSingle: .HorizontalAlignment = xlCenter
                    End With

                    ' -- ROW R+4: Truck | Gross --
                    .Rows(R + 4).RowHeight = 13
                    .Range(.Cells(R + 4, 1), .Cells(R + 4, 4)).Merge
                    With .Cells(R + 4, 1)
                        .Value = "Truck: " & Trim(sTruck)
                        .Font.Bold = True
                        .HorizontalAlignment = xlLeft
                    End With
                    With .Cells(R + 4, 6): .Value = "Gross:": .HorizontalAlignment = xlRight: End With
                    With .Cells(R + 4, 7): .Value = sInFmt: .HorizontalAlignment = xlRight: End With
                    .Cells(R + 4, 8).Value = "In"
                    .Cells(R + 4, 9).Value = "Scale"

                    ' -- ROW R+5: Tare --
                    .Rows(R + 5).RowHeight = 13
                    With .Cells(R + 5, 6): .Value = "Tare:": .HorizontalAlignment = xlRight: End With
                    With .Cells(R + 5, 7): .Value = sOutFmt: .HorizontalAlignment = xlRight: End With
                    .Cells(R + 5, 8).Value = "Out"
                    .Cells(R + 5, 9).Value = "Scale"

                    ' -- ROW R+6: Customer | Net --
                    .Rows(R + 6).RowHeight = 22
                    .Range(.Cells(R + 6, 1), .Cells(R + 6, 5)).Merge
                    With .Cells(R + 6, 1)
                        ' Bold label, regular/bold text handled by formatting the whole cell
                        .Value = "Customer: " & Trim(sHauler) & IIf(sWasteData <> "", "   Waste Data: " & Trim(sWasteData), "")
                        .WrapText = True
                        .Font.Size = 11
                        .Font.Bold = True
                        .VerticalAlignment = xlTop
                        .HorizontalAlignment = xlLeft
                    End With
                    With .Cells(R + 6, 6): .Value = "Net:": .HorizontalAlignment = xlRight: .Font.Bold = True: End With
                    With .Cells(R + 6, 7): .Value = sNetFmt: .Font.Bold = True: .HorizontalAlignment = xlRight: End With

                    ' -- ROW R+7: Memo --
                    .Rows(R + 7).RowHeight = 25
                    With .Cells(R + 7, 1): .Value = "Memo:": .Font.Bold = True: End With
                    .Range(.Cells(R + 7, 2), .Cells(R + 7, 9)).Merge
                    With .Cells(R + 7, 2)
                        .Value = sSeller
                        .WrapText = True
                        .VerticalAlignment = xlTop
                    End With

                    ' -- ROW R+9: Table column headers --
                    .Rows(R + 9).RowHeight = 14
                    With .Cells(R + 9, 1): .Value = "Address": .Font.Bold = True: End With
                    With .Cells(R + 9, 3): .Value = "Materials & Services": .Font.Bold = True: End With
                    With .Cells(R + 9, 6): .Value = "Quantity": .Font.Bold = True: End With
                    With .Cells(R + 9, 7): .Value = "Unit": .Font.Bold = True: End With
                    With .Cells(R + 9, 8): .Value = "$": .Font.Bold = True: End With
                    With .Cells(R + 9, 9): .Value = "Amount": .Font.Bold = True: End With
                    .Range(.Cells(R + 9, 1), .Cells(R + 9, 9)).Borders(xlEdgeTop).LineStyle = xlContinuous
                    .Range(.Cells(R + 9, 1), .Cells(R + 9, 9)).Borders(xlEdgeBottom).LineStyle = xlContinuous

                    ' -- ROW R+10: Data row --
                    .Rows(R + 10).RowHeight = 14
                    .Cells(R + 10, 1).Value = sOrigin
                    .Cells(R + 10, 3).Value = sProd
                    .Cells(R + 10, 6).Value = sNetTn
                    .Cells(R + 10, 7).Value = "Ton"
                    .Cells(R + 10, 8).Value = ""
                    .Cells(R + 10, 9).Value = sPaid

                    ' -- ROWS R+11-12: Spacers --
                    .Rows(R + 11).RowHeight = 4
                    .Rows(R + 12).RowHeight = 4

                    ' -- ROW R+13: Payment Method & Total Amount --
                    .Rows(R + 13).RowHeight = 14
                    With .Cells(R + 13, 1): .Value = "Payment Method:  " & sPayment: .Font.Bold = True: End With
                    .Range(.Cells(R + 13, 6), .Cells(R + 13, 8)).Merge
                    With .Range(.Cells(R + 13, 6), .Cells(R + 13, 6))
                        .Value = "Total Amount:": .Font.Bold = True: .HorizontalAlignment = xlRight
                    End With
                    With .Cells(R + 13, 9): .Value = sPaid: .Font.Bold = True: End With

                    ' -- ROW R+14: Terms --
                    .Rows(R + 14).RowHeight = 69
                    .Range(.Cells(R + 14, 1), .Cells(R + 14, 9)).Merge
                    With .Range(.Cells(R + 14, 1), .Cells(R + 14, 1))
                        .Value = TERMS: .Font.Size = 5: .WrapText = True
                        .VerticalAlignment = xlTop
                    End With

                    ' -- ROW R+15: Driver / Deputy Weighmaster signatures --
                    .Rows(R + 15).RowHeight = 18
                    With .Cells(R + 15, 1): .Value = "Driver:": .Font.Bold = True: End With
                    With .Cells(R + 15, 2): .Value = String(20, "_"): .Font.Size = 8: End With
                    .Range(.Cells(R + 15, 6), .Cells(R + 15, 8)).Merge
                    With .Range(.Cells(R + 15, 6), .Cells(R + 15, 6))
                        .Value = "Deputy Weighmaster:": .Font.Bold = True
                        .HorizontalAlignment = xlRight
                    End With
                    With .Cells(R + 15, 9): .Value = sByF: .Font.Italic = True: .Font.Size = 10: End With

                    ' -- ROW R+16: WEIGHMASTER CERTIFICATE header --
                    .Rows(R + 16).RowHeight = 14
                    .Range(.Cells(R + 16, 1), .Cells(R + 16, 9)).Merge
                    With .Range(.Cells(R + 16, 1), .Cells(R + 16, 1))
                        If certHeader <> "" Then
                            .Value = certHeader
                            .Font.Bold = True: .Font.Size = 10
                            .HorizontalAlignment = xlCenter
                            .Font.Underline = xlUnderlineStyleSingle
                        Else
                            .Value = ""
                        End If
                    End With

                    ' -- ROW R+17: WM cert text + WM name --
                    .Rows(R + 17).RowHeight = 68
                    .Range(.Cells(R + 17, 2), .Cells(R + 17, 8)).Merge
                    With .Range(.Cells(R + 17, 2), .Cells(R + 17, 2))
                        .Value = WMCERT: .Font.Size = 5: .WrapText = True
                        .VerticalAlignment = xlTop
                        .HorizontalAlignment = xlCenter
                    End With
                    With .Cells(R + 17, 9): .Value = sWM: .Font.Italic = True: .Font.Size = 9: End With

                    ' -- BOTTOM SPACERS --
                    .Rows(R + 18).RowHeight = 8
                    .Rows(R + 19).RowHeight = 6

                    ' -- Perf-line divider (Removed) --
                    If iCp = 1 Then
                        .Rows(22).RowHeight = 2
                        .Range(.Cells(22, 1), .Cells(22, 9)).Merge
                    End If

                End With
            Next iCp
            
            ' --- FORCE PERFECT SYMMETRY ---
            Dim rMatch As Long
            For rMatch = 1 To 21
                wsPrint.Rows(rMatch + 22).RowHeight = wsPrint.Rows(rMatch).RowHeight
            Next rMatch

            ' 6. Font
            wsPrint.Range("A1:I47").Font.Name = "Calibri"

            ' 7. PAGE SETUP
            With wsPrint.PageSetup
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = 1
                .Orientation = xlPortrait
                .PaperSize = xlPaperLetter
                .TopMargin = 0
                .BottomMargin = 0
                .LeftMargin = Application.InchesToPoints(0.2)
                .RightMargin = Application.InchesToPoints(0.2)
                .HeaderMargin = 0
                .FooterMargin = 0
                .CenterHorizontally = True
                .CenterVertically = False
                .PrintArea = "A1:I43"
                .CenterFooter = ""
            End With

        End Sub

        Public Sub ExportVisibleTicketsToPDF()
            Dim ws As Worksheet, tbl As ListObject, wmCol As ListColumn
            Dim row As ListRow
            Dim wbTemp As Workbook
            Dim wsPrint As Worksheet
            Dim ticketNum As String
            Dim currTicket As String
            Dim pdfPath As String
            Dim fDialog As FileDialog
            
            ' Ask user where to save it
            Set fDialog = Application.FileDialog(msoFileDialogSaveAs)
            With fDialog
                .Title = "Save PDF Export As"
                .InitialFileName = Environ("USERPROFILE") & "\Desktop\Exported_Tickets_" & Format(Date, "mm-dd-yyyy") & ".pdf"
                .FilterIndex = 2 ' PDF
                If .Show = -1 Then
                    pdfPath = .SelectedItems(1)
                Else
                    Exit Sub
                End If
            End With
            
            If TryGetUserFormTicketTable(ws, tbl) Then
                On Error Resume Next
                Set wmCol = tbl.ListColumns("Weighmaster")
                On Error GoTo 0
            End If
            
            If tbl Is Nothing Or wmCol Is Nothing Then
                MsgBox "Could not find Weighmaster column in table.", vbCritical
                Exit Sub
            End If
            
            ' Save current form ticket so we can restore it later
            currTicket = Me.txtTicketNumber.Value
            
            ' Suppress UI and events during batch generation
            Application.ScreenUpdating = False
            Application.EnableEvents = False
            Application.Calculation = xlCalculationManual
            
            Set wbTemp = Workbooks.Add
            
            For Each row In tbl.ListRows
                ' Check if row is visible (handling AutoFilter)
                If Not row.Range.Rows(1).Hidden Then
                    ticketNum = Trim(CStr(row.Range(1, wmCol.Index).Value))
                    If ticketNum <> "" Then
                        ' Load it into the form
                        Me.LoadTableTicket ticketNum
                        
                        ' Get the permanent print sheet (it will be cleared and overwritten)
                        Set wsPrint = GetOrCreatePrintSheet()
                        
                        ' Generate the ticket on wsPrint without showing preview
                        Call DrawCurrentTicketToSheet(wsPrint)
                        
                        ' Copy the generated sheet to the temp workbook (creates a new tab)
                        wsPrint.Copy After:=wbTemp.Sheets(wbTemp.Sheets.count)
                        DoEvents
                    End If
                End If
            Next row
            
            ' Remove the default empty "Sheet1" in wbTemp
            Application.DisplayAlerts = False
            If wbTemp.Sheets.count > 1 Then
                wbTemp.Sheets(1).Delete
            End If
            Application.DisplayAlerts = True
            
            ' If no tabs were successfully generated, abort
            If wbTemp.Sheets.count = 0 Or (wbTemp.Sheets.count = 1 And wbTemp.Sheets(1).Name Like "Sheet*") Then
                wbTemp.Close SaveChanges:=False
                Application.Calculation = xlCalculationAutomatic
                Application.EnableEvents = True
                Application.ScreenUpdating = True
                MsgBox "No visible tickets found to export.", vbInformation
                Exit Sub
            End If
            
            ' Export all sheets in temp workbook into one PDF
            On Error Resume Next
            wbTemp.Sheets.Select
            wbTemp.ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath
            Dim errNum As Long
            errNum = Err.Number
            On Error GoTo 0
            
            wbTemp.Close SaveChanges:=False
            
            ' Restore form to whatever they were looking at
            If currTicket <> "" Then
                Me.LoadTableTicket currTicket
            Else
                Call cmdNewTicket_Click
            End If
            
            Application.Calculation = xlCalculationAutomatic
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            
            If errNum = 0 Then
                MsgBox "Export complete! PDF saved to:" & vbCrLf & pdfPath, vbInformation
            Else
                MsgBox "Failed to save PDF. Make sure the file isn't open and you have permission to save there.", vbCritical
            End If
        End Sub
