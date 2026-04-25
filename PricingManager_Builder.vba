
' ========================================
' BUILD PRICING MANAGER FORM
' Run once to create PricingManagerForm.
' Then paste PricingManager_Code.vba into
' the form's code view (F7).
' ========================================
Sub BuildPricingManagerForm()
    Dim vbp As Object
    Dim vbc As Object
    Dim uf  As Object
    Dim ctl As Object

    On Error Resume Next
    Set vbp = ActiveWorkbook.VBProject
    If Err.Number <> 0 Then
        MsgBox "Enable 'Trust access to the VBA project object model' in Trust Center > Macro Settings.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    ' Remove existing form if present
    On Error Resume Next
    Set vbc = vbp.VBComponents("PricingManagerForm")
    If Not vbc Is Nothing Then vbp.VBComponents.Remove vbc
    On Error GoTo 0

    ' Create Form
    Set vbc = vbp.VBComponents.Add(3)
    vbc.Name = "PricingManagerForm"
    vbc.Properties("Caption") = "Pricing Manager"
    vbc.Properties("Width")   = 860
    vbc.Properties("Height")  = 560
    vbc.Properties("ScrollBars") = 0
    vbc.Properties("StartUpPosition") = 2

    Set uf = vbc.Designer

    ' TITLE STRIPE
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblTitleBar")
    ctl.Caption = "  PRICING MANAGER": ctl.Left = 0: ctl.Top = 0
    ctl.Width = 860: ctl.Height = 28
    ctl.BackColor = RGB(20, 60, 120): ctl.BackStyle = 1
    ctl.ForeColor = RGB(255, 255, 255): ctl.Font.Bold = True: ctl.Font.Size = 11

    ' TAB BUTTONS
    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdScaleTab")
    ctl.Caption = "Scale Pricing": ctl.Left = 10: ctl.Top = 34
    ctl.Width = 160: ctl.Height = 26: ctl.Font.Bold = True: ctl.Font.Size = 9
    ctl.BackColor = RGB(0, 120, 215): ctl.ForeColor = RGB(255, 255, 255)

    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdHaulTab")
    ctl.Caption = "Hauling Pricing": ctl.Left = 176: ctl.Top = 34
    ctl.Width = 160: ctl.Height = 26: ctl.Font.Bold = True: ctl.Font.Size = 9
    ctl.BackColor = RGB(220, 220, 220): ctl.ForeColor = RGB(50, 50, 50)

    ' ACTION BUTTONS
    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdNew")
    ctl.Caption = "+ New": ctl.Left = 490: ctl.Top = 34
    ctl.Width = 72: ctl.Height = 26: ctl.Font.Bold = True: ctl.Font.Size = 9
    ctl.BackColor = RGB(0, 180, 0): ctl.ForeColor = RGB(255, 255, 255)

    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdDuplicate")
    ctl.Caption = "Copy Row": ctl.Left = 568: ctl.Top = 34
    ctl.Width = 72: ctl.Height = 26: ctl.Font.Bold = True: ctl.Font.Size = 9
    ctl.BackColor = RGB(255, 165, 0)

    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdSave")
    ctl.Caption = "Save Row": ctl.Left = 646: ctl.Top = 34
    ctl.Width = 96: ctl.Height = 26: ctl.Font.Bold = True: ctl.Font.Size = 9
    ctl.BackColor = RGB(0, 120, 215): ctl.ForeColor = RGB(255, 255, 255)

    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdDelete")
    ctl.Caption = "Delete Row": ctl.Left = 748: ctl.Top = 34
    ctl.Width = 96: ctl.Height = 26: ctl.Font.Bold = True: ctl.Font.Size = 9
    ctl.BackColor = RGB(200, 40, 40): ctl.ForeColor = RGB(255, 255, 255)

    ' SEPARATOR
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSep0")
    ctl.Caption = "": ctl.Left = 0: ctl.Top = 62: ctl.Width = 860: ctl.Height = 2
    ctl.BackColor = RGB(160, 160, 160): ctl.BackStyle = 1

    ' ── SCALE PRICING LIST ──
    Set ctl = uf.Controls.Add("Forms.ListBox.1", "lstScale")
    ctl.Left = 8: ctl.Top = 66: ctl.Width = 840: ctl.Height = 185
    ctl.Font.Name = "Courier New": ctl.Font.Size = 8

    ' Scale edit field headers (y=258)
    Dim hY As Long: hY = 258
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSHdr1")
    ctl.Caption = "PRODUCT/SERVICE": ctl.Left = 8: ctl.Top = hY: ctl.Width = 95: ctl.Height = 13
    ctl.Font.Bold = True: ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128)
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSHdr2")
    ctl.Caption = "ACCOUNT (blank=public)": ctl.Left = 111: ctl.Top = hY: ctl.Width = 200: ctl.Height = 13
    ctl.Font.Bold = True: ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128)
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSHdr3")
    ctl.Caption = "START DATE": ctl.Left = 319: ctl.Top = hY: ctl.Width = 90: ctl.Height = 13
    ctl.Font.Bold = True: ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128)
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSHdr4")
    ctl.Caption = "END DATE": ctl.Left = 417: ctl.Top = hY: ctl.Width = 90: ctl.Height = 13
    ctl.Font.Bold = True: ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128)
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSHdr5")
    ctl.Caption = "RATE ($)": ctl.Left = 515: ctl.Top = hY: ctl.Width = 80: ctl.Height = 13
    ctl.Font.Bold = True: ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128)
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSHdr6")
    ctl.Caption = "TYPE (ton / run)": ctl.Left = 603: ctl.Top = hY: ctl.Width = 110: ctl.Height = 13
    ctl.Font.Bold = True: ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128)

    ' Scale edit fields (y=273)
    Dim fY As Long: fY = 273
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSProduct")
    ctl.Caption = "Product/Service": ctl.Left = 8: ctl.Top = fY: ctl.Width = 95: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128)
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtSProduct")
    ctl.Left = 8: ctl.Top = fY + 13: ctl.Width = 95: ctl.Height = 20

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSAccount")
    ctl.Caption = "Account": ctl.Left = 111: ctl.Top = fY: ctl.Width = 200: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128)
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtSAccount")
    ctl.Left = 111: ctl.Top = fY + 13: ctl.Width = 200: ctl.Height = 20

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSStart")
    ctl.Caption = "Start Date": ctl.Left = 319: ctl.Top = fY: ctl.Width = 90: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128)
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtSStart")
    ctl.Left = 319: ctl.Top = fY + 13: ctl.Width = 90: ctl.Height = 20

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSEnd")
    ctl.Caption = "End Date": ctl.Left = 417: ctl.Top = fY: ctl.Width = 90: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128)
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtSEnd")
    ctl.Left = 417: ctl.Top = fY + 13: ctl.Width = 90: ctl.Height = 20

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSRate")
    ctl.Caption = "Rate ($)": ctl.Left = 515: ctl.Top = fY: ctl.Width = 80: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128)
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtSRate")
    ctl.Left = 515: ctl.Top = fY + 13: ctl.Width = 80: ctl.Height = 20

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblSType")
    ctl.Caption = "Type": ctl.Left = 603: ctl.Top = fY: ctl.Width = 80: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128)
    Set ctl = uf.Controls.Add("Forms.ComboBox.1", "cmbSType")
    ctl.Left = 603: ctl.Top = fY + 13: ctl.Width = 80: ctl.Height = 20: ctl.Style = 2
    ctl.AddItem "ton": ctl.AddItem "run": ctl.Value = "ton"

    ' ── HAULING PRICING LIST (hidden by default) ──
    Set ctl = uf.Controls.Add("Forms.ListBox.1", "lstHaul")
    ctl.Left = 8: ctl.Top = 66: ctl.Width = 840: ctl.Height = 185
    ctl.Font.Name = "Courier New": ctl.Font.Size = 8: ctl.Visible = False

    ' Haul field headers row 1 (y=258)
    Dim hH As Long: hH = 258
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblHHdr1")
    ctl.Caption = "ACCOUNT (blank=public)": ctl.Left = 8: ctl.Top = hH: ctl.Width = 140: ctl.Height = 13
    ctl.Font.Bold = True: ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128): ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblHHdr2")
    ctl.Caption = "START DATE": ctl.Left = 156: ctl.Top = hH: ctl.Width = 100: ctl.Height = 13
    ctl.Font.Bold = True: ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128): ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblHHdr3")
    ctl.Caption = "END DATE": ctl.Left = 264: ctl.Top = hH: ctl.Width = 100: ctl.Height = 13
    ctl.Font.Bold = True: ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128): ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblHHdr4")
    ctl.Caption = "Lowboy / CONCRETE / 10YD / 40YD rates ($)": ctl.Left = 372: ctl.Top = hH: ctl.Width = 280: ctl.Height = 13
    ctl.Font.Bold = True: ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128): ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblHHdr5")
    ctl.Caption = "Allowed Tons (Low,Con,10,40)": ctl.Left = 660: ctl.Top = hH: ctl.Width = 188: ctl.Height = 13
    ctl.Font.Bold = True: ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128): ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblHHdr6")
    ctl.Caption = "": ctl.Left = 8: ctl.Top = hH: ctl.Width = 1: ctl.Height = 1: ctl.Visible = False

    ' Haul edit fields row 1 (y=273)
    Dim hfY As Long: hfY = 273
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblHAccount")
    ctl.Caption = "Account": ctl.Left = 8: ctl.Top = hfY: ctl.Width = 140: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128): ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtHAccount")
    ctl.Left = 8: ctl.Top = hfY + 13: ctl.Width = 140: ctl.Height = 20: ctl.Visible = False

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblHStart")
    ctl.Caption = "Start Date": ctl.Left = 156: ctl.Top = hfY: ctl.Width = 100: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128): ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtHStart")
    ctl.Left = 156: ctl.Top = hfY + 13: ctl.Width = 100: ctl.Height = 20: ctl.Visible = False

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblHEnd")
    ctl.Caption = "End Date": ctl.Left = 264: ctl.Top = hfY: ctl.Width = 100: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128): ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtHEnd")
    ctl.Left = 264: ctl.Top = hfY + 13: ctl.Width = 100: ctl.Height = 20: ctl.Visible = False

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblHLowboy")
    ctl.Caption = "Lowboy ($)": ctl.Left = 372: ctl.Top = hfY: ctl.Width = 64: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128): ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtHLowboy")
    ctl.Left = 372: ctl.Top = hfY + 13: ctl.Width = 64: ctl.Height = 20: ctl.Visible = False

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblHConcrete")
    ctl.Caption = "CONCRETE ($)": ctl.Left = 444: ctl.Top = hfY: ctl.Width = 70: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128): ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtHConcrete")
    ctl.Left = 444: ctl.Top = hfY + 13: ctl.Width = 70: ctl.Height = 20: ctl.Visible = False

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblH10YD")
    ctl.Caption = "10YD ($)": ctl.Left = 522: ctl.Top = hfY: ctl.Width = 64: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128): ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtH10YD")
    ctl.Left = 522: ctl.Top = hfY + 13: ctl.Width = 64: ctl.Height = 20: ctl.Visible = False

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblH40YD")
    ctl.Caption = "40YD ($)": ctl.Left = 594: ctl.Top = hfY: ctl.Width = 64: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128): ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtH40YD")
    ctl.Left = 594: ctl.Top = hfY + 13: ctl.Width = 64: ctl.Height = 20: ctl.Visible = False

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblHAllowed")
    ctl.Caption = "Allowed Tons (Low,Con,10,40)": ctl.Left = 666: ctl.Top = hfY: ctl.Width = 182: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128): ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtHAllowed")
    ctl.Left = 666: ctl.Top = hfY + 13: ctl.Width = 182: ctl.Height = 20: ctl.Visible = False

    ' Haul edit fields row 2 (y=313)
    Dim hfY2 As Long: hfY2 = 313
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblHOverage")
    ctl.Caption = "Overage ($/ton over)": ctl.Left = 8: ctl.Top = hfY2: ctl.Width = 130: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128): ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtHOverage")
    ctl.Left = 8: ctl.Top = hfY2 + 13: ctl.Width = 130: ctl.Height = 20: ctl.Visible = False

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblHMiles")
    ctl.Caption = "Miles": ctl.Left = 146: ctl.Top = hfY2: ctl.Width = 60: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128): ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtHMiles")
    ctl.Left = 146: ctl.Top = hfY2 + 13: ctl.Width = 60: ctl.Height = 20: ctl.Visible = False

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblHDistRate")
    ctl.Caption = "Dist. Rate ($/mile)": ctl.Left = 214: ctl.Top = hfY2: ctl.Width = 110: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128): ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtHDistRate")
    ctl.Left = 214: ctl.Top = hfY2 + 13: ctl.Width = 110: ctl.Height = 20: ctl.Visible = False

    Set ctl = uf.Controls.Add("Forms.Label.1", "lblHSurch")
    ctl.Caption = "Surcharge (%)": ctl.Left = 332: ctl.Top = hfY2: ctl.Width = 90: ctl.Height = 12
    ctl.Font.Size = 7: ctl.ForeColor = RGB(0, 64, 128): ctl.Visible = False
    Set ctl = uf.Controls.Add("Forms.TextBox.1", "txtHSurch")
    ctl.Left = 332: ctl.Top = hfY2 + 13: ctl.Width = 90: ctl.Height = 20: ctl.Visible = False

    ' STATUS BAR
    Set ctl = uf.Controls.Add("Forms.Label.1", "lblStatus")
    ctl.Caption = "Ready.": ctl.Left = 0: ctl.Top = 528: ctl.Width = 860: ctl.Height = 18
    ctl.BackColor = RGB(240, 240, 240): ctl.BackStyle = 1
    ctl.Font.Size = 7: ctl.ForeColor = RGB(80, 80, 80)

    ' CLOSE BUTTON
    Set ctl = uf.Controls.Add("Forms.CommandButton.1", "cmdClose")
    ctl.Caption = "Close": ctl.Left = 752: ctl.Top = 496: ctl.Width = 96: ctl.Height = 26
    ctl.Font.Bold = True: ctl.Font.Size = 9
    ctl.BackColor = RGB(100, 100, 100): ctl.ForeColor = RGB(255, 255, 255)

    MsgBox "PricingManagerForm built!" & vbCrLf & vbCrLf & _
           "NEXT STEP:" & vbCrLf & _
           "  1. Press Alt+F11 to open the VBA editor" & vbCrLf & _
           "  2. Click 'PricingManagerForm' in the left panel" & vbCrLf & _
           "  3. Press F7 to open the code view" & vbCrLf & _
           "  4. Paste the entire contents of PricingManager_Code.vba" & vbCrLf & vbCrLf & _
           "Then run ShowPricingManager to open it!", vbInformation, "Done"
End Sub

' Open the Pricing Manager from a sheet button or macro
Sub ShowPricingManager()
    PricingManagerForm.Show
End Sub
