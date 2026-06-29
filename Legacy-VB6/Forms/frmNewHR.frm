VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmNewHR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hico Resources"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   389
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   546
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2640
      Top             =   5055
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Employee's Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2010
      Top             =   5430
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewHR.frx":0000
            Key             =   "0'"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewHR.frx":0892
            Key             =   "7'"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewHR.frx":0CE4
            Key             =   "5'"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewHR.frx":1136
            Key             =   "6'"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewHR.frx":3470
            Key             =   "4'"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewHR.frx":414A
            Key             =   "2'"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewHR.frx":49DC
            Key             =   "3'"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmNewHR.frx":718E
            Key             =   "1'"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Resources"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5295
      Left            =   90
      TabIndex        =   3
      Top             =   15
      Width           =   7980
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   105
         TabIndex        =   10
         Top             =   270
         Width           =   7725
         Begin MSForms.ComboBox cmbCatID 
            Height          =   315
            Left            =   1485
            TabIndex        =   14
            Tag             =   "GroupID"
            Top             =   210
            Width           =   2640
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            DisplayStyle    =   3
            Size            =   "4657;556"
            MatchEntry      =   1
            ListStyle       =   1
            ShowDropButtonWhen=   1
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Group :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   16
            Left            =   810
            TabIndex        =   13
            Top             =   270
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Resource Code :"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Index           =   4
            Left            =   4200
            TabIndex        =   12
            Top             =   270
            Width           =   1350
         End
         Begin MSForms.TextBox TBox 
            Height          =   315
            Index           =   3
            Left            =   5580
            TabIndex        =   11
            Tag             =   "RID1"
            Top             =   225
            Width           =   1650
            VariousPropertyBits=   679495707
            BorderStyle     =   1
            Size            =   "2910;556"
            SpecialEffect   =   0
            FontName        =   "Tahoma"
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   4170
         Left            =   120
         TabIndex        =   4
         Top             =   1005
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   7355
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   582
         ShowFocusRect   =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "   General Info."
         TabPicture(0)   =   "frmNewHR.frx":7A20
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         Begin VB.Frame Frame4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2370
            Left            =   360
            TabIndex        =   5
            Top             =   525
            Width           =   6765
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Resource Description"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   9
               Top             =   180
               Width           =   1515
            End
            Begin MSForms.TextBox TBox 
               Height          =   315
               Index           =   0
               Left            =   105
               TabIndex        =   8
               Tag             =   "RName"
               Top             =   390
               Width           =   6555
               VariousPropertyBits=   679495707
               BorderStyle     =   1
               Size            =   "11562;556"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin MSForms.TextBox TBox 
               Height          =   1275
               Index           =   5
               Left            =   105
               TabIndex        =   7
               Tag             =   "RUsage"
               Top             =   975
               Width           =   6555
               VariousPropertyBits=   -1467987941
               BorderStyle     =   1
               Size            =   "11562;2249"
               SpecialEffect   =   0
               FontName        =   "Tahoma"
               FontHeight      =   165
               FontCharSet     =   0
               FontPitchAndFamily=   2
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Repair Vender Information"
               ForeColor       =   &H00400000&
               Height          =   195
               Index           =   0
               Left            =   135
               TabIndex        =   6
               Top             =   765
               Width           =   1905
            End
         End
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   4
         Left            =   90
         TabIndex        =   15
         Top             =   645
         Visible         =   0   'False
         Width           =   1725
         VariousPropertyBits=   679495707
         MaxLength       =   9
         Size            =   "3043;556"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   6495
      TabIndex        =   16
      Top             =   5370
      Width           =   1575
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2778;661"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CheckBox chkInactive 
      Height          =   345
      Left            =   105
      TabIndex        =   2
      Tag             =   "InActive"
      Top             =   5355
      Width           =   1170
      BackColor       =   -2147483633
      ForeColor       =   8388608
      DisplayStyle    =   4
      Size            =   "2064;609"
      Value           =   "0"
      Caption         =   "Inactive"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   375
      Left            =   4890
      TabIndex        =   1
      Top             =   5370
      Width           =   1575
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2778;661"
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdNext 
      Height          =   375
      Left            =   3285
      TabIndex        =   0
      Top             =   5370
      Width           =   1575
      Caption         =   "  Save & New"
      PicturePosition =   327683
      Size            =   "2778;661"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmNewHR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ItemCode As Long

Dim SelItm As ListItem

Dim DoPrint As Boolean

Public Sub EditItem(strItemCode As Long)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    ItemCode = strItemCode
    'cmbCatID.Locked = True
    'TBox(3).Locked = True
    With rs
        .Open "Select * From VHR Where RID=" & ItemCode, con, adOpenForwardOnly, adLockReadOnly
        
        For Each C In Controls
            If C.Tag <> "" Then
                If TypeOf C Is MSForms.CheckBox Then
                    C.value = .Fields(C.Tag)
                ElseIf TypeOf C Is ComboList.Usercontrol1 Then
                    'C.Text = .Fields(C.Tag & "Text") & ""
                    C.ID = .Fields(C.Tag) & ""
                ElseIf TypeOf C Is VB.Label Then
                    C.Caption = .Fields(C.Tag) & ""
                Else
                    If C.Name = "cmbCatID" Then
                        C.Text = .Fields("GroupName")
                    Else
                        C.Text = .Fields(C.Tag) & ""
                    End If
                End If
            End If
        Next
        
        'TBox(8) = ![CatID] & "-"
        TBox(3) = ![RID1] & ""
        .Close
    End With
    
    Set rs = Nothing
    Me.Show 1
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub FillCmbs()
        
    'Call AddToCombo(cmbPchAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '" & PchAcc & "%' AND Parent=0 Order By ACCTitle")
    'Call cmbSaleAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where AccNo Like '" & SALEACC & "%' AND Parent=0 Order By ACCTitle")
    'cmbSaleAcc.ID = "31-001-001"
    'Call AddToCombo(cmbPRTAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '" & PRTAcc & "%' AND Parent=0 Order By ACCTitle")
    'Call cmbSaleRAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where AccNo Like '" & SRTACC & "%' AND Parent=0 Order By ACCTitle")
    'cmbSaleRAcc.ID = "31-003-003"
    
    'Call cmbAssetAcc.AddVals(con, "AccTitle", "VActiveAccounts", "AccNo", " Where AccNo Like '" & STOCKASSETACC & "%' AND Parent=0 Order By ACCTitle")
    'cmbAssetAcc.ID = "15-005-001"
    'Call AddToCombo(cmbAssetAcc, "AccTitle", "VActiveAccounts", " Where AccNo Like '" & StockAssetAcc & "%' AND Parent=0 Order By ACCTitle")
    
    'Call AddToCombo(cmbUnit, "Distinct Unit", "Items", "Order By Unit")
    'Call AddToCombo(cmbWeightUnit, "Distinct UnitWeight", "Items", "Order By UnitWeight")
    
    Call AddToCombo(cmbCatID, "Description", "HRGroups", " Order By ID ", , "ID")
    
    
    
    '----------------------------------------------------------------
End Sub

Private Sub cmbCatID_Change()
    Call cmbCatID_Click
End Sub

Private Sub cmbCatID_Click()
    If Not cmbCatID.MatchFound Then Exit Sub

End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    If SaveItem Then
        Unload Me
        
        Me.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If SaveItem Then Unload Me
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Public Sub AddNew()
    Call GetID
    Me.Show 1
End Sub
Private Sub GetID()
    
    On Error GoTo err
    If ShowingData Then Exit Sub
    Dim strVendID As String
    
    
    'strVendID = Val(con.Execute("Select Max(Cast(SubString(VendID1,5,4) as int)) From Venders Where UPPER(Left(VendID1,3))=" & UCase(Left(cmbProcess.Text, 3))).Fields(0).value & "")
    
    'If VendID <> 0 Then
        
    '    TBox(3) = VendID1
    '    Exit Sub
        
    'End If
    
    
    strVendID = Val(con.Execute("Select Max(Cast(SubString(RID1,4,4) as int)) From HICOResources").Fields(0).value & "")
    
    If strVendID = 0 Then
        strVendID = 0
    End If
    
    strVendID = strVendID + 1
    
    TBox(3) = "HS" & Format(strVendID, "-0000")
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    
    On Error Resume Next
    For i = 0 To SSTab1.Tabs - 1
        SSTab1.Tab = i
        SSTab1.Picture = ImageList1.ListImages(i & "'").ExtractIcon
    Next
   
    SSTab1.Tab = 0
    
    Me.KeyPreview = True
    '    DTPicker1.value = ServerDate
    'cmbSaleAcc.ListHeight = 1500
    'cmbSaleRAcc.ListHeight = 1500
    'cmbAssetAcc.ListHeight = 1500
    
    cmbCatID.Locked = False
    TBox(3).Locked = False
    
    
    
    Call FillCmbs
    
    'FRev.Visible = False
    
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err

    Dim OpBalance As Double
    
    If Trim(TBox(0)) = "" Then
        MsgBox "Invalid Item Name", vbInformation
        SaveItem = False
        Exit Function
    End If
    
    Dim PchAcc As String, ItmPRTAcc As String, AssetAcc As String

    Dim Sql As String
    
    Call StartTrans(con)

    If ItemCode = 0 Then
            
        Sql = CreateInsertSQL("HicoResources")
        'con.Execute "Update Accounts Set OpenBal=OpenBal+" & TotalCost & " Where AccNo='" & AssetAcc & "'"
        'Call adjustLedger(AssetAcc, "1/1/1900", 0, TotalCost)
    Else
        Sql = CreateUpdateSQL("HicoResources")
        Sql = Sql & " Where RID=" & ItemCode
    End If
    
    con.Execute Sql
    
    'If ItemCode <> "" Then Call SaveRevision
    
    con.CommitTrans
    '---------------------------------------------------------
    'If DoPrint Then
    '    If MsgBox("Item Has Changed." & vbNewLine & "Do You Want To Print This Item?", vbQuestion + vbYesNo) = vbYes Then
    '        Load frmPrevRpt
    '        frmPrevRpt.ShowReport "{VrptDMR.ItemID}='" & TBox(4) & "'", rptDMR
    '    End If
    'End If
    '---------------------------------------------------------
    ItemCode = 0
    SaveItem = True
    
    Exit Function

err:
    MsgBox err.Description
    
End Function



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Set RevEntry = Nothing
End Sub


Private Sub Form_Unload(Cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub Tbox_Change(Index As Integer)
    Exit Sub
    Select Case Index
    
    Case 8, 3
        TBox(4) = TBox(8) & TBox(3)
    End Select
    
End Sub

Private Sub Tbox_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Exit Sub
    Select Case Index
    Case 2, 3
        KeyAscii = OnlyNumber(Val(KeyAscii))
    Case 5
        KeyAscii = OnlyNumber(Val(KeyAscii), True)
    End Select
    
End Sub

Private Sub Tbox_Validate(Index As Integer, Cancel As Boolean)
    Exit Sub
    If Index = 5 Then
        TBox(Index) = Format(Val(TBox(Index)), "#0.00")
    ElseIf Index = 2 Or Index = 3 Then
        TBox(Index) = Format(Val(TBox(Index)), "#0")
    End If
End Sub


Private Function CreateInsertSQL(strTable As String) As String
    
    Dim C As Control
    Dim RetString As String
    
    
    'Create The First Part Of Insert Query...
    RetString = "Insert Into " & strTable & "("
    '---------------------------------------------------------------------------
    For Each C In Controls
        
        If (TypeOf C Is MSForms.TextBox) Or (TypeOf C Is MSForms.ComboBox) Or (TypeOf C Is MSForms.CheckBox) Or (TypeOf C Is ComboList.Usercontrol1) Then
            If C.Tag <> "" Then
                RetString = RetString & C.Tag & ","
            End If
        End If
    Next
    
    'Get For ListView
    'For i = 1 To LV2.ListItems.Count
    '    RetString = RetString & LV2.ListItems(i).Key & ","
    'Next i
    '---------------------------------------------------------------------------
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ") Values("
    
    
    'Now  Get The Second Part...
    '---------------------------------------------------------------------------
    For Each C In Controls
        
        
        If C.Tag <> "" Then
            If (TypeOf C Is MSForms.TextBox) Or (TypeOf C Is MSForms.ComboBox) Then
            
                If C.Name = "cmbCatID" Then
                    RetString = RetString & Val(C.List(C.ListIndex, 1)) & ","
                ElseIf C.Name = "NBox" Then
                    RetString = RetString & Val(C.Text & "") & ","
                ElseIf C.Name = "cmbSteel" Or C.Name = "cmbGuage" Or C.Name = "cmbItemGroup" Then
                    RetString = RetString & C.List(C.ListIndex, 1) & ","
                Else
                    RetString = RetString & "'" & C.Text & "',"
                End If
                
            ElseIf (TypeOf C Is MSForms.CheckBox) Then
                RetString = RetString & Abs(C.value) & ","
            ElseIf (TypeOf C Is ComboList.Usercontrol1) Then
                'If C.Name = cmbSteel Or C.Name = cmbGuage Then
                '    RetString = RetString & C.ID & ","
                RetString = RetString & "'" & C.ID & "',"
            End If
        End If
    Next
    
    
    'Get For ListView
    'For i = 1 To LV2.ListItems.Count
    '    RetString = RetString & Abs(LV2.ListItems(i).Checked) & ","
    'Next i
    
    '---------------------------------------------------------------------------
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ")"
    CreateInsertSQL = RetString
    
End Function
Private Function CreateUpdateSQL(strTable As String) As String
    
    Dim C As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "Update " & strTable & " Set "
    '---------------------------------------------------------------------------
    For Each C In Controls
    
        If C.Tag <> "" Then
            If (TypeOf C Is MSForms.TextBox) Or (TypeOf C Is MSForms.ComboBox) Or (TypeOf C Is MSForms.ComboBox) Then '
                If C.Name = "cmbCatID" Then
                    RetString = RetString & C.Tag & "=" & Val(C.List(C.ListIndex, 1)) & ","
                ElseIf C.Name = "NBox" Then
                    RetString = RetString & C.Tag & "=" & C.Text & ","
                ElseIf C.Name = "cmbSteel" Or C.Name = "cmbGuage" Or C.Name = "cmbItemGroup" Then
                    RetString = RetString & C.Tag & "=" & C.List(C.ListIndex, 1) & ","
                Else
                    RetString = RetString & C.Tag & "='" & C.Text & "',"
                End If
            ElseIf (TypeOf C Is MSForms.CheckBox) Then
                RetString = RetString & C.Tag & "=" & Abs(C.value) & ","
            ElseIf (TypeOf C Is ComboList.Usercontrol1) Then
                RetString = RetString & C.Tag & "='" & C.ID & "',"
            End If
        End If
    Next
    
    If NotEdited = False Then
        
        DoPrint = True
    Else
        DoPrint = False
    End If
    '---------------------------------------------------------------------------
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function

Private Sub PrintDMR()
    
End Sub
