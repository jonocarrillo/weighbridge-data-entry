Attribute VB_Name = "FormBuilder"
Option Explicit

Sub BuildWeighmasterForm()
    Dim vbp As Object
    Dim vbc As Object
    Dim uf As Object
    Dim ctl As Object
    
    ' ══════════════════════════════════════════════════
    ' FONT CONSTANTS — change these to restyle everything
    ' Every control on the form gets these via the post-build pass below.
    ' Edit one constant, rebuild the form, the whole UI updates.
    ' ══════════════════════════════════════════════════
    Const fFontName     As String  = "Segoe UI"  ' Universal font family for ALL controls
    Const fHeader       As Integer = 12   ' Main form title
    Const fSubtitle     As Integer = 7    ' Subtitle / fine print
    Const fFieldLabel   As Integer = 8    ' Small blue field-name labels
    Const fFieldValue   As Integer = 10   ' Standard input / display text
    Const fSectionLabel As Integer = 11   ' "Inbound Weight", "Outbound Weight"
    Const fNetWeight    As Integer = 15   ' "Net Weight:" label
    Const fNetWeightBox As Integer = 13   ' Net weight value box
    Const fNetTonsBox   As Integer = 17   ' Big green tons readout
    Const fNav          As Integer = 10   ' Ticket counter / nav labels
    Const fNavBtn       As Integer = 11   ' < / > / +New nav buttons
    Const fButton       As Integer = 12   ' Action buttons (Save, Print…)
    Const fCounter      As Integer = 9    ' lblTicketCounter row indicator
    Const fInput        As Integer = 10   ' Universal size for TextBox / ComboBox input
    ' Check if we have access to VBProject
    On Error Resume Next
    Set vbp = ActiveWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "Please enable 'Trust access to the VBA project object model' in Trust Center > Macro Settings.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Remove existing UserForm1 if it exists
    On Error Resume Next
    Set vbc = vbp.VBComponents("UserForm1")
    If Not vbc Is Nothing Then
        vbp.VBComponents.Remove vbc
    End If
    On Error GoTo 0
    
    ' ── Create wide landscape-style UserForm ──
    Set vbc = vbp.VBComponents.Add(3) ' vbext_ct_MSForm
    vbc.Name = "UserForm1"
    vbc.Properties("Caption") = "Weighmaster Certificate"
    vbc.Properties("Width") = 840
    vbc.Properties("Height") = 530
    vbc.Properties("ScrollBars") = 2  ' fmScrollBarsVertical
    vbc.Properties("ScrollHeight") = 620
    vbc.Properties("KeepScrollBarsVisible") = 2
    vbc.Properties("Zoom") = 75
    
    Set uf = vbc.Designer
    
    ' ══════════════════════════════════════════════════
    ' HEADER (full width)
    ' ══════════════════════════════════════════════════
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblHeader")
    ctl.Caption = "WEIGHMASTER CERTIFICATE": ctl.Left = 10: ctl.Top = 6: ctl.Width = 800: ctl.Height = 18
    ctl.Font.Bold = True: ctl.Font.Size = fHeader: ctl.TextAlign = 2

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSubtitle")
    ctl.Caption = "THIS IS TO CERTIFY that the following described commodity was weighed, measured or counted by a weighmaster whose signature is on this certificate."
    ctl.Left = 10: ctl.Top = 24: ctl.Width = 800: ctl.Height = 14
    ctl.Font.Size = fSubtitle: ctl.TextAlign = 2

    ' ══════════════════════════════════════════════════
    ' TOP 3-COLUMN AREA  (mirrors reference ticket top)
    '   LEFT  x=10..264    (255px)  Company / Truck / Customer / Waste Data
    '   CENTER x=274..547  (274px)  Logo placeholder
    '   RIGHT  x=557..810  (254px)  Ticket# / Date / Time / Weights
    ' ══════════════════════════════════════════════════

    ' ── LEFT COLUMN ──────────────────────────────────
    ' WEIGHED AT (company label - static display)
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblWeighedAt")
    ctl.Caption = "WEIGHED AT": ctl.Left = 10: ctl.Top = 44: ctl.Width = 120: ctl.Height = 12
    ctl.Font.Size = fFieldLabel: ctl.ForeColor = RGB(0, 64, 128)

    Set ctl = uf.Controls.Add("Forms.Label.1", "txtWeighedAt")
    ctl.Caption = "AMERICAN INDUSTRIAL SERVICES"
    ctl.Left = 10: ctl.Top = 57: ctl.Width = 255: ctl.Height = 20
    ctl.Font.Bold = True: ctl.Font.Size = fFieldValue

    ' TRUCK #
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblTruck")
    ctl.Caption = "TRUCK #": ctl.Left = 10: ctl.Top = 82: ctl.Width = 100: ctl.Height = 12
    ctl.Font.Size = fFieldLabel: ctl.ForeColor = RGB(0, 64, 128)

    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtTruckNumber")
    ctl.Left = 10: ctl.Top = 95: ctl.Width = 255: ctl.Height = 22

    ' COMPANY NAME / HAULER
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblCompany")
    ctl.Caption = "COMPANY NAME / HAULER:": ctl.Left = 10: ctl.Top = 122: ctl.Width = 200: ctl.Height = 12
    ctl.Font.Size = fFieldLabel: ctl.ForeColor = RGB(0, 64, 128)

    Set ctl = uf.Controls.Add("Forms.ComboBox.1", "txtCompanyNameHauler")
    ctl.Left = 10: ctl.Top = 135: ctl.Width = 255: ctl.Height = 20: ctl.Style = 0

    ' WASTE DATA (license plate etc.)
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblWasteData")
    ctl.Caption = "WASTE DATA:": ctl.Left = 10: ctl.Top = 162: ctl.Width = 100: ctl.Height = 12
    ctl.Font.Size = fFieldLabel: ctl.ForeColor = RGB(0, 64, 128)

    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtWasteData")
    ctl.Left = 10: ctl.Top = 175: ctl.Width = 255: ctl.Height = 22

    ' ── CENTER COLUMN (Logo from embedded Assets sheet) ──────────
    Set ctl = uf.Controls.Add("Forms.Image.1", "imgLogo")
    ctl.Left = 274: ctl.Top = 44: ctl.Width = 274: ctl.Height = 153
    ctl.BorderStyle = 0    ' fmBorderStyleNone
    ctl.BackStyle = 0      ' fmBackStyleTransparent
    ctl.PictureSizeMode = 3  ' fmPictureSizeModeZoom
    
    ' --- Load logo from hidden Assets sheet (embedded in the workbook) ---
    Dim assetWs As Worksheet
    Dim logoLoaded As Boolean
    logoLoaded = False
    
    On Error Resume Next
    Set assetWs = ThisWorkbook.Worksheets("Assets")
    On Error GoTo 0
    
    If Not assetWs Is Nothing Then
        If assetWs.Shapes.Count > 0 Then
            Dim tmpPath As String
            tmpPath = Environ("TEMP") & "\ais_logo_temp.jpg"
            
            On Error Resume Next
            If Len(Dir(tmpPath)) > 0 Then Kill tmpPath
            On Error GoTo 0
            
            ' Find the picture specifically in A1
            Dim sourceShp As Shape
            Dim tempShp As Shape
            Set sourceShp = Nothing
            
            For Each tempShp In assetWs.Shapes
                ' Check if it's a picture and if its TopLeftCell is A1
                If tempShp.Type = msoPicture Or tempShp.Type = msoLinkedPicture Then
                    If Not Intersect(tempShp.TopLeftCell, assetWs.Range("A1")) Is Nothing Then
                        Set sourceShp = tempShp
                        Exit For
                    End If
                End If
            Next tempShp
            
            ' Fallback: if we didn't find one explicitly in A1, just grab the first picture
            If sourceShp Is Nothing Then
                For Each tempShp In assetWs.Shapes
                    If tempShp.Type = msoPicture Or tempShp.Type = msoLinkedPicture Then
                        Set sourceShp = tempShp
                        Exit For
                    End If
                Next tempShp
            End If
            
            If Not sourceShp Is Nothing Then
                Dim tmpChtObj As Object
                ' Temporarily make sheet visible to export chart properly
                Dim wasHidden As Boolean
                wasHidden = (assetWs.Visible <> xlSheetVisible)
                If wasHidden Then assetWs.Visible = xlSheetVisible
                
                Set tmpChtObj = assetWs.ChartObjects.Add(0, 0, sourceShp.Width, sourceShp.Height)
                sourceShp.Copy
                tmpChtObj.Chart.Paste
                tmpChtObj.Chart.Export Filename:=tmpPath, FilterName:="JPG"
                tmpChtObj.Delete
                
                If wasHidden Then assetWs.Visible = xlSheetVeryHidden
                
                ' Wait for file system to catch up
                DoEvents
                
                If Len(Dir(tmpPath)) > 0 Then
                    On Error Resume Next
                    ctl.Picture = LoadPicture(tmpPath)
                    If Err.Number = 0 Then logoLoaded = True
                    On Error Resume Next
                    Kill tmpPath
                    On Error GoTo 0
                End If
            End If
        End If
    End If
    
    ' Fallback: show a styled placeholder if no logo was found
    If Not logoLoaded Then
        ctl.BackStyle = 1   ' fmBackStyleOpaque
        ctl.BackColor = RGB(240, 248, 240)
        ' Add a label on top as a visual placeholder
        Dim lblPlaceholder As Object
        Set lblPlaceholder = uf.Controls.Add("Forms.Label.1", "lblLogo")
        lblPlaceholder.Caption = "[ AIS LOGO ]" & vbCrLf & "(add logo to Assets sheet)"
        lblPlaceholder.Left = 274: lblPlaceholder.Top = 44
        lblPlaceholder.Width = 274: lblPlaceholder.Height = 153
    lblPlaceholder.TextAlign = 2: lblPlaceholder.Font.Size = fFieldValue: lblPlaceholder.Font.Bold = True
        lblPlaceholder.ForeColor = RGB(0, 120, 0): lblPlaceholder.BackStyle = 0
    End If

    ' ── RIGHT COLUMN ─────────────────────────────────
    ' WEIGHMASTER # / TICKET
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblTicket")
    ctl.Caption = "WEIGHMASTER #": ctl.Left = 558: ctl.Top = 44: ctl.Width = 110: ctl.Height = 12
    ctl.Font.Size = fFieldLabel: ctl.ForeColor = RGB(0, 64, 128)

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblInvoiceNum")
    ctl.Caption = "INVOICE NUM": ctl.Left = 692: ctl.Top = 44: ctl.Width = 120: ctl.Height = 12
    ctl.Font.Size = fFieldLabel: ctl.ForeColor = RGB(0, 64, 128): ctl.TextAlign = 3

    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtTicketNumber")
    ctl.Left = 558: ctl.Top = 57: ctl.Width = 126: ctl.Height = 22: ctl.TextAlign = 3

    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtInvoiceNum")
    ctl.Left = 692: ctl.Top = 57: ctl.Width = 120: ctl.Height = 22: ctl.TextAlign = 3
    ctl.Locked = True
    ctl.BackColor = RGB(240, 240, 240)

    ' DATE
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblDate1")
    ctl.Caption = "DATE": ctl.Left = 558: ctl.Top = 84: ctl.Width = 60: ctl.Height = 12
    ctl.Font.Size = fFieldLabel: ctl.ForeColor = RGB(0, 64, 128)

    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtDate1")
    ctl.Left = 558: ctl.Top = 97: ctl.Width = 254: ctl.Height = 22

    ' TIME IN
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblTimeIn")
    ctl.Caption = "TIME IN": ctl.Left = 558: ctl.Top = 124: ctl.Width = 60: ctl.Height = 12
    ctl.Font.Size = fFieldLabel: ctl.ForeColor = RGB(0, 64, 128)

    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtTime1")
    ctl.Left = 558: ctl.Top = 137: ctl.Width = 155: ctl.Height = 22

    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdTime1")
    ctl.Caption = "T": ctl.Left = 718: ctl.Top = 138: ctl.Width = 22: ctl.Height = 20
    ctl.Font.Bold = True

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblInboundTag")
    ctl.Caption = "INBOUND": ctl.Left = 744: ctl.Top = 140: ctl.Width = 68: ctl.Height = 16
    ctl.ForeColor = RGB(0, 128, 0): ctl.Font.Bold = True: ctl.Font.Size = fFieldValue

    ' TIME OUT
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblTimeOut")
    ctl.Caption = "TIME OUT": ctl.Left = 558: ctl.Top = 163: ctl.Width = 60: ctl.Height = 12
    ctl.Font.Size = fFieldLabel: ctl.ForeColor = RGB(0, 64, 128)

    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtTime2")
    ctl.Left = 558: ctl.Top = 176: ctl.Width = 155: ctl.Height = 22

    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdTime2")
    ctl.Caption = "T": ctl.Left = 718: ctl.Top = 177: ctl.Width = 22: ctl.Height = 20
    ctl.Font.Bold = True

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblOutboundTag")
    ctl.Caption = "OUTBOUND": ctl.Left = 744: ctl.Top = 179: ctl.Width = 68: ctl.Height = 16
    ctl.ForeColor = RGB(200, 0, 0): ctl.Font.Bold = True: ctl.Font.Size = fFieldValue

    ' Hidden compatibility controls
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtDate2")
    ctl.Left = 820: ctl.Top = 44: ctl.Width = 1: ctl.Height = 1: ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.OptionButton.1", "optInbound1")
    ctl.Caption = "": ctl.Left = 820: ctl.Top = 44: ctl.Width = 1: ctl.Height = 1: ctl.GroupName = "Weight1": ctl.Value = True: ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.OptionButton.1", "optOutbound1")
    ctl.Caption = "": ctl.Left = 820: ctl.Top = 44: ctl.Width = 1: ctl.Height = 1: ctl.GroupName = "Weight1": ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.OptionButton.1", "optInbound2")
    ctl.Caption = "": ctl.Left = 820: ctl.Top = 44: ctl.Width = 1: ctl.Height = 1: ctl.GroupName = "Weight2": ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.OptionButton.1", "optOutbound2")
    ctl.Caption = "": ctl.Left = 820: ctl.Top = 44: ctl.Width = 1: ctl.Height = 1: ctl.GroupName = "Weight2": ctl.Value = True: ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtTruckID1")
    ctl.Left = 820: ctl.Top = 44: ctl.Width = 1: ctl.Height = 1: ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtTruckID2")
    ctl.Left = 820: ctl.Top = 44: ctl.Width = 1: ctl.Height = 1: ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblTrk1")
    ctl.Caption = "Truck ID: ------": ctl.Left = 820: ctl.Top = 44: ctl.Width = 1: ctl.Height = 1: ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblTrk2")
    ctl.Caption = "Truck ID: ------": ctl.Left = 820: ctl.Top = 44: ctl.Width = 1: ctl.Height = 1: ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblInboundWt1")
    ctl.Caption = "Inbound Wt: 0 lb": ctl.Left = 820: ctl.Top = 44: ctl.Width = 1: ctl.Height = 1: ctl.Visible = False

    ' ══════════════════════════════════════════════════
    ' WEIGHT TOTALS SECTION (below top columns, right-aligned display)
    ' ══════════════════════════════════════════════════
    Dim yW As Long: yW = 204

    ' Separator
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSep1")
    ctl.Caption = "": ctl.Left = 10: ctl.Top = yW: ctl.Width = 800: ctl.Height = 1
    ctl.BackColor = RGB(192, 192, 192): ctl.BackStyle = 1

    yW = yW + 6

    ' Inbound Weight
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblInWt")
    ctl.Caption = "Inbound Weight (GROSS):": ctl.Left = 10: ctl.Top = yW + 4: ctl.Width = 200: ctl.Height = 18
    ctl.Font.Size = fSectionLabel
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtInboundWeight")
    ctl.Left = 560: ctl.Top = yW: ctl.Width = 160: ctl.Height = 24: ctl.TextAlign = 3
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblLb2")
    ctl.Caption = "lb": ctl.Left = 726: ctl.Top = yW + 4: ctl.Width = 25: ctl.Height = 15

    yW = yW + 28

    ' Outbound Weight
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblOutWt")
    ctl.Caption = "Outbound Weight (TARE):": ctl.Left = 10: ctl.Top = yW + 4: ctl.Width = 200: ctl.Height = 18
    ctl.Font.Size = fSectionLabel
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtOutboundWeight")
    ctl.Left = 560: ctl.Top = yW: ctl.Width = 160: ctl.Height = 24: ctl.TextAlign = 3
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblLb1")
    ctl.Caption = "lb": ctl.Left = 726: ctl.Top = yW + 4: ctl.Width = 25: ctl.Height = 15

    yW = yW + 30

    ' Net Weight
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblNetWt")
    ctl.Caption = "Net Weight:": ctl.Left = 10: ctl.Top = yW: ctl.Width = 160: ctl.Height = 28
    ctl.Font.Bold = True: ctl.Font.Size = fNetWeight
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtNetWeightLbs")
    ctl.Text = "0 lb": ctl.Left = 560: ctl.Top = yW: ctl.Width = 160: ctl.Height = 28
    ctl.Font.Bold = True: ctl.Font.Size = fNetWeightBox: ctl.TextAlign = 3

    ' Separator
    yW = yW + 32
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSep2")
    ctl.Caption = "": ctl.Left = 10: ctl.Top = yW: ctl.Width = 800: ctl.Height = 1
    ctl.BackColor = RGB(192, 192, 192): ctl.BackStyle = 1

    ' ══════════════════════════════════════════════════
    ' NAVIGATION + PAYMENT ROW
    ' ══════════════════════════════════════════════════
    Dim y As Long: y = yW + 8

    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdPrev")
    ctl.Caption = "<": ctl.Left = 10: ctl.Top = y: ctl.Width = 28: ctl.Height = 22
    ctl.Font.Bold = True: ctl.Font.Size = fNavBtn

    ' ── lblTicketCounter: widened to show full "NEW  ·  Sheet row NNN (new)" text ──
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblTicketCounter")
    ctl.Caption = "NEW": ctl.Left = 42: ctl.Top = y + 3: ctl.Width = 200: ctl.Height = 16
    ctl.TextAlign = 1: ctl.Font.Size = fCounter: ctl.Font.Bold = False

    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdNext")
    ctl.Caption = ">": ctl.Left = 96: ctl.Top = y: ctl.Width = 28: ctl.Height = 22
    ctl.Font.Bold = True: ctl.Font.Size = fNavBtn

    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdNewTicket")
    ctl.Caption = "+ New": ctl.Left = 130: ctl.Top = y: ctl.Width = 55: ctl.Height = 22
    ctl.Font.Bold = True: ctl.Font.Size = fNav

    ' Payment options (right side of nav row)
    ' Payment label must shift right to make room for wider counter label
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblPay")
    ctl.Caption = "PAID:": ctl.Left = 248: ctl.Top = y + 3: ctl.Width = 35: ctl.Height = 15
    ctl.Font.Bold = True

    Set ctl = uf.Controls.Add("Forms.OptionButton.1", "optCredit")
    ctl.Caption = "CREDIT": ctl.Left = 286: ctl.Top = y + 3: ctl.Width = 60: ctl.Height = 18: ctl.GroupName = "Payment"
    Set ctl = uf.Controls.Add("Forms.OptionButton.1", "optDebit")
    ctl.Caption = "DEBIT": ctl.Left = 350: ctl.Top = y + 3: ctl.Width = 55: ctl.Height = 18: ctl.GroupName = "Payment"
    Set ctl = uf.Controls.Add("Forms.OptionButton.1", "optAR")
    ctl.Caption = "AR": ctl.Left = 409: ctl.Top = y + 3: ctl.Width = 36: ctl.Height = 18: ctl.GroupName = "Payment"
    Set ctl = uf.Controls.Add("Forms.OptionButton.1", "optCash")
    ctl.Caption = "CASH": ctl.Left = 449: ctl.Top = y + 3: ctl.Width = 55: ctl.Height = 18: ctl.GroupName = "Payment"
    Set ctl = uf.Controls.Add("Forms.OptionButton.1", "optExpense")
    ctl.Caption = "EXPENSE": ctl.Left = 508: ctl.Top = y + 3: ctl.Width = 70: ctl.Height = 18: ctl.GroupName = "Payment"

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblAmt")
    ctl.Caption = "$": ctl.Left = 620: ctl.Top = y + 3: ctl.Width = 16: ctl.Height = 15
    ctl.Font.Bold = True
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtPaidAmount")
    ctl.Left = 638: ctl.Top = y: ctl.Width = 160: ctl.Height = 22: ctl.TextAlign = 3

    ' ══════════════════════════════════════════════════
    ' FULL-WIDTH FIELDS BELOW (PO, Comment, Origin, Product)
    ' ══════════════════════════════════════════════════
    Dim gap As Long: gap = 30
    y = y + 30

    ' Separator
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSep3")
    ctl.Caption = "": ctl.Left = 10: ctl.Top = y: ctl.Width = 800: ctl.Height = 1
    ctl.BackColor = RGB(192, 192, 192): ctl.BackStyle = 1
    y = y + 6

    ' WEIGHED FOR / SELLER (Comment)
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSeller")
    ctl.Caption = "MEMO:": ctl.Left = 10: ctl.Top = y: ctl.Width = 200: ctl.Height = 12
    ctl.Font.Size = fFieldLabel: ctl.ForeColor = RGB(0, 64, 128)
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtWeighedForSeller")
    ctl.Left = 10: ctl.Top = y + 13: ctl.Width = 800: ctl.Height = 20
    y = y + gap

    ' CITY (dropdown) — feeds into ORIGIN column
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblCity")
    ctl.Caption = "CITY:": ctl.Left = 10: ctl.Top = y: ctl.Width = 60: ctl.Height = 12
    ctl.Font.Size = fFieldLabel: ctl.ForeColor = RGB(0, 64, 128)

    ' ADDRESS — free-text, right of city
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblOrigin")
    ctl.Caption = "ADDRESS (WASTE ORIGIN):": ctl.Left = 316: ctl.Top = y: ctl.Width = 200: ctl.Height = 12
    ctl.Font.Size = fFieldLabel: ctl.ForeColor = RGB(0, 64, 128)

    Set ctl = uf.Controls.Add("Forms.ComboBox.1", "cmbCity")
    ctl.Left = 10: ctl.Top = y + 13: ctl.Width = 300: ctl.Height = 20: ctl.Style = 2 ' fmStyleDropDownList (strict dropdown)

    Set ctl = uf.Controls.Add("Forms.ComboBox.1", "txtWasteOrigin")
    ctl.Left = 316: ctl.Top = y + 13: ctl.Width = 494: ctl.Height = 20: ctl.Style = 0
    y = y + gap

    ' PRODUCT CODE
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblProduct")
    ctl.Caption = "PRODUCT CODE:": ctl.Left = 10: ctl.Top = y: ctl.Width = 200: ctl.Height = 12
    ctl.Font.Size = fFieldLabel: ctl.ForeColor = RGB(0, 64, 128)
    Set ctl = uf.Controls.Add("Forms.ComboBox.1", "txtProductCode")
    ctl.Left = 10: ctl.Top = y + 13: ctl.Width = 800: ctl.Height = 20: ctl.Style = 0
    y = y + gap

    ' Separator
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSep4")
    ctl.Caption = "": ctl.Left = 10: ctl.Top = y: ctl.Width = 800: ctl.Height = 1
    ctl.BackColor = RGB(192, 192, 192): ctl.BackStyle = 1
    y = y + 8

    ' ══════════════════════════════════════════════════
    ' NET TONS  +  WEIGHMASTER  +  BY
    ' ══════════════════════════════════════════════════
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblNetTons")
    ctl.Caption = "NET WEIGHT": ctl.Left = 10: ctl.Top = y: ctl.Width = 90: ctl.Height = 14
    ctl.Font.Size = fFieldValue
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblTons")
    ctl.Caption = "TONS": ctl.Left = 120: ctl.Top = y: ctl.Width = 40: ctl.Height = 14
    ctl.Font.Size = fFieldValue
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblWM")
    ctl.Caption = "- WEIGHMASTER": ctl.Left = 430: ctl.Top = y: ctl.Width = 130: ctl.Height = 14
    ctl.Font.Size = fFieldValue
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblByLbl")
    ctl.Caption = "BY:": ctl.Left = 666: ctl.Top = y + 16: ctl.Width = 25: ctl.Height = 15
    ctl.Font.Bold = True

    y = y + 16
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtNetWeightTons")
    ctl.Text = "0.00": ctl.Left = 10: ctl.Top = y: ctl.Width = 130: ctl.Height = 30
    ctl.Font.Bold = True: ctl.Font.Size = fNetTonsBox: ctl.ForeColor = RGB(0, 80, 0)
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtWeighmaster")
    ctl.Left = 420: ctl.Top = y + 2: ctl.Width = 240: ctl.Height = 22
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtBy")
    ctl.Left = 696: ctl.Top = y: ctl.Width = 120: ctl.Height = 22

    y = y + 36

    ' Separator
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSep5")
    ctl.Caption = "": ctl.Left = 10: ctl.Top = y: ctl.Width = 800: ctl.Height = 1
    ctl.BackColor = RGB(192, 192, 192): ctl.BackStyle = 1
    y = y + 6

    ' ══════════════════════════════════════════════════
    ' CERTIFICATE TYPE CHECKBOXES
    ' ══════════════════════════════════════════════════
    Set ctl = uf.Controls.Add("Forms.CheckBox.1", "chkWeightCert")
    ctl.Caption = "Weight Certificate": ctl.Left = 10: ctl.Top = y: ctl.Width = 150: ctl.Height = 18
    ctl.Value = False  ' Unchecked by default
    
    Set ctl = uf.Controls.Add("Forms.CheckBox.1", "chkRecyclingCert")
    ctl.Caption = "Recycling C&D Certificate": ctl.Left = 170: ctl.Top = y: ctl.Width = 200: ctl.Height = 18
    
    y = y + 24

    ' ══════════════════════════════════════════════════
    ' SAVE / SAVE & PRINT / PRINT / PREVIEW BUTTONS
    ' ══════════════════════════════════════════════════
    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdSaveTicket")
    ctl.Caption = "Save Ticket": ctl.Left = 10: ctl.Top = y: ctl.Width = 190: ctl.Height = 32
    ctl.Font.Bold = True: ctl.Font.Size = fButton: ctl.BackColor = RGB(64, 192, 255)

    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdSaveAndPrint")
    ctl.Caption = "Save && Print": ctl.Left = 208: ctl.Top = y: ctl.Width = 220: ctl.Height = 32
    ctl.Font.Bold = True: ctl.Font.Size = fButton: ctl.BackColor = RGB(255, 140, 0)

    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdPrintTicket")
    ctl.Caption = "Print": ctl.Left = 436: ctl.Top = y: ctl.Width = 175: ctl.Height = 32
    ctl.Font.Bold = True: ctl.Font.Size = fButton: ctl.BackColor = RGB(0, 180, 0)

    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdPrintPreview")
    ctl.Caption = "Preview": ctl.Left = 619: ctl.Top = y: ctl.Width = 195: ctl.Height = 32
    ctl.Font.Bold = True: ctl.Font.Size = fButton: ctl.BackColor = RGB(130, 130, 130)

    ' Hidden compatibility buttons
    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdClear")
    ctl.Caption = "Clear": ctl.Left = 820: ctl.Top = y: ctl.Width = 1: ctl.Height = 1: ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdClose")
    ctl.Caption = "Close": ctl.Left = 820: ctl.Top = y: ctl.Width = 1: ctl.Height = 1: ctl.Visible = False

    ' Adjust scroll height
    vbc.Properties("ScrollHeight") = y + 50
    vbc.Properties("Height") = 530

    ' ══════════════════════════════════════════════════
    ' UNIVERSAL FONT PASS — apply fFontName to every control,
    ' and apply fInput to TextBox / ComboBox controls (which
    ' otherwise inherit Excel's default and can't be resized
    ' from one place). Skips controls whose size was set explicitly.
    ' ══════════════════════════════════════════════════
    Dim c As Object
    For Each c In uf.Controls
        On Error Resume Next
        c.Font.Name = fFontName
        Dim tn As String
        tn = TypeName(c)
        If tn = "TextBox" Or tn = "ComboBox" Then
            ' Only override controls left at default (8pt). Anything that was
            ' explicitly bumped (e.g. txtNetWeightTons at fNetTonsBox) is preserved.
            If c.Font.Size = 8 Then c.Font.Size = fInput
        End If
        On Error GoTo 0
    Next c

    MsgBox "Form built! Now paste code from UserForm_Code.vba (F7 to open code view).", vbInformation
End Sub

' ========================================
' BUILD LOGIN FORM
' Run once alongside BuildWeighmasterForm
' Creates a small 3-button weighmaster picker
' ========================================

Sub BuildLoginForm()
    Dim vbp As Object
    Dim vbc As Object
    Dim uf As Object
    Dim ctl As Object
    
    On Error Resume Next
    Set vbp = ActiveWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "Enable 'Trust access to the VBA project object model' in Trust Center.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Remove existing LoginForm if present
    On Error Resume Next
    Set vbc = vbp.VBComponents("LoginForm")
    If Not vbc Is Nothing Then vbp.VBComponents.Remove vbc
    On Error GoTo 0
    
    ' Create LoginForm
    Set vbc = vbp.VBComponents.Add(3) ' vbext_ct_MSForm
    vbc.Name = "LoginForm"
    vbc.Properties("Caption") = "Who is entering today?"
    vbc.Properties("Width") = 230
    vbc.Properties("Height") = 140
    vbc.Properties("StartUpPosition") = 2 ' center screen
    
    Set uf = vbc.Designer
    
    ' Prompt label
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblPrompt")
    ctl.Caption = "Select your name to continue:": ctl.Left = 10: ctl.Top = 10
    ctl.Width = 200: ctl.Height = 18: ctl.TextAlign = 2
    ctl.Font.Bold = True: ctl.Font.Size = 10
    
    ' Bryan button
    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdBryan")
    ctl.Caption = "Bryan": ctl.Left = 10: ctl.Top = 36: ctl.Width = 60: ctl.Height = 28
    ctl.Font.Bold = True: ctl.Font.Size = 11: ctl.BackColor = RGB(0, 180, 0)
    
    ' Jesus button
    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdJesus")
    ctl.Caption = "Jesus": ctl.Left = 80: ctl.Top = 36: ctl.Width = 60: ctl.Height = 28
    ctl.Font.Bold = True: ctl.Font.Size = 11: ctl.BackColor = RGB(0, 192, 192)
    
    ' Noe button
    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdNoe")
    ctl.Caption = "Noe": ctl.Left = 150: ctl.Top = 36: ctl.Width = 60: ctl.Height = 28
    ctl.Font.Bold = True: ctl.Font.Size = 11: ctl.BackColor = RGB(0, 128, 255)
    
    MsgBox "LoginForm built! Paste code from LoginForm_Code.vba into LoginForm's code view.", vbInformation
End Sub

' ========================================
' SETUP LOGO SHEET
' Run this ONCE to create the hidden Assets sheet.
' Then insert your logo image onto that sheet.
' ========================================
Sub SetupLogoSheet()
    Dim ws As Worksheet
    
    ' Check if Assets sheet already exists
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Assets")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' Create the Assets sheet
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Assets"
    End If
    
    ' Hide the sheet so it doesn't clutter the workbook
    ws.Visible = xlSheetVeryHidden   ' xlVeryHidden = users can't unhide via Format > Sheet
    
    ' Navigate to it temporarily so the user can insert the picture
    ws.Visible = xlSheetVisible
    ws.Activate
    
    MsgBox "The 'Assets' sheet has been created and is now visible." & vbCrLf & vbCrLf & _
           "NEXT STEPS:" & vbCrLf & _
           "  1. While on this sheet, go to Insert > Pictures > This Device" & vbCrLf & _
           "  2. Select your company logo file (PNG or JPG recommended)" & vbCrLf & _
           "  3. Resize the image to roughly fit cell A1" & vbCrLf & _
           "  4. Then run  SetupLogoSheet_Hide  to hide the sheet again" & vbCrLf & vbCrLf & _
           "After that, re-run BuildWeighmasterForm and your logo will appear!", _
           vbInformation, "Setup Logo Sheet"
End Sub

' Run this after inserting your logo to hide the Assets sheet again
Sub SetupLogoSheet_Hide()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Assets")
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        ws.Visible = xlSheetVeryHidden
        MsgBox "Assets sheet hidden. Now run BuildWeighmasterForm to rebuild the form with your logo.", _
               vbInformation, "Done"
    Else
        MsgBox "Assets sheet not found. Run SetupLogoSheet first.", vbExclamation
    End If
End Sub
