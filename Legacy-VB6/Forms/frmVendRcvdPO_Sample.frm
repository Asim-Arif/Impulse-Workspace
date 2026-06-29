VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmVendRcvdPO_Sample 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accept / Reject"
   ClientHeight    =   8355
   ClientLeft      =   1095
   ClientTop       =   285
   ClientWidth     =   11940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   11940
   WindowState     =   2  'Maximized
   Begin VB.Frame FAll 
      Height          =   8205
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11790
      Begin VB.Frame FFB 
         Caption         =   "Feedback :"
         Height          =   3900
         Left            =   0
         TabIndex        =   16
         Top             =   2100
         Visible         =   0   'False
         Width           =   11580
         Begin VB.TextBox txtQtyRejected 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5205
            TabIndex        =   24
            Top             =   1020
            Width           =   1725
         End
         Begin VB.TextBox txtQtyAccepted 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1860
            TabIndex        =   23
            Top             =   1020
            Width           =   1725
         End
         Begin VB.CommandButton cmdCancelFB 
            Caption         =   "Cance&l"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   10095
            TabIndex        =   22
            Top             =   3240
            Width           =   1305
         End
         Begin VB.CommandButton cmdUpdateFB 
            Caption         =   "&Update"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   8730
            TabIndex        =   21
            Top             =   3240
            Width           =   1305
         End
         Begin VB.TextBox txtItemCode 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   255
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   585
            Width           =   11160
         End
         Begin VB.TextBox txtFilePath 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   255
            TabIndex        =   19
            Top             =   3330
            Width           =   7845
         End
         Begin VB.CommandButton cmdBrowseBarcode 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   8130
            TabIndex        =   18
            Top             =   3315
            Width           =   405
         End
         Begin VB.TextBox txtRemarks 
            Alignment       =   2  'Center
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1605
            Left            =   240
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   1470
            Width           =   11160
         End
         Begin MSComDlg.CommonDialog CD1 
            Left            =   10875
            Top             =   2850
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
            Filter          =   "PDF Files Only|*.pdf"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Qty Rejected :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   3690
            TabIndex        =   28
            Top             =   1095
            Width           =   1440
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Qty Accepted :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   270
            TabIndex        =   27
            Top             =   1080
            Width           =   1515
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Item"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   270
            TabIndex        =   26
            Top             =   300
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "PDF Path to Attach"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   270
            TabIndex        =   25
            Top             =   3105
            Width           =   2040
         End
      End
      Begin VB.Frame fTitle 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   15
         TabIndex        =   7
         Top             =   0
         Width           =   11775
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Accept / Reject"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   390
            Index           =   0
            Left            =   4575
            TabIndex        =   8
            Top             =   120
            Width           =   2340
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Accept / Reject"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   390
            Index           =   1
            Left            =   4590
            TabIndex        =   9
            Top             =   135
            Width           =   2340
         End
      End
      Begin VB.Frame Frame6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   9765
         TabIndex        =   4
         Top             =   7425
         Width           =   1710
         Begin MSForms.CommandButton cmdCancel 
            CausesValidation=   0   'False
            Height          =   360
            Left            =   120
            TabIndex        =   5
            Top             =   210
            Width           =   1425
            Caption         =   "Close"
            PicturePosition =   327683
            Size            =   "2514;635"
            Accelerator     =   67
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.PictureBox PicEdit 
         Appearance      =   0  'Flat
         BackColor       =   &H00E7EBEF&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7860
         ScaleHeight     =   285
         ScaleWidth      =   2355
         TabIndex        =   1
         Top             =   2055
         Visible         =   0   'False
         Width           =   2385
         Begin VB.TextBox txtSheets 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1170
            TabIndex        =   3
            Top             =   -15
            Width           =   1200
         End
         Begin VB.TextBox txtQty 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   -15
            TabIndex        =   2
            Top             =   -15
            Width           =   1200
         End
      End
      Begin MSComctlLib.ListView LV 
         Height          =   5955
         Left            =   420
         TabIndex        =   6
         Top             =   1440
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   10504
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   15199215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Material"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Rcvd Qty"
            Object.Width           =   2117
         EndProperty
      End
      Begin MSComCtl2.DTPicker DT 
         Height          =   300
         Left            =   9825
         TabIndex        =   10
         Top             =   1140
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd-MMM-yyyy"
         Format          =   243204099
         CurrentDate     =   -8461
      End
      Begin MSForms.Label Label3 
         Height          =   240
         Left            =   420
         TabIndex        =   15
         Top             =   930
         Width           =   11055
         BackColor       =   11517387
         Caption         =   $"frmVendRcvdPO_Sample.frx":0000
         Size            =   "19500;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.TextBox txtRcvNo 
         Height          =   285
         Left            =   420
         TabIndex        =   14
         Top             =   1155
         Width           =   1755
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "3096;503"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtDesc 
         Height          =   285
         Left            =   6825
         TabIndex        =   13
         Top             =   1155
         Width           =   3015
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "5318;503"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtVender 
         Height          =   285
         Left            =   2145
         TabIndex        =   12
         Top             =   1155
         Width           =   3015
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "5318;503"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtPONo 
         Height          =   285
         Left            =   5145
         TabIndex        =   11
         Top             =   1155
         Width           =   1695
         VariousPropertyBits=   746604575
         BorderStyle     =   1
         Size            =   "2990;503"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmVendRcvdPO_Sample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strRcvNo As String
Dim strFileName  As String
Private Sub cmbOrderNo_matched()

    On Error GoTo err
    Dim LastGroup As String
    Dim rs As New ADODB.Recordset
    
    With rs
        
        .Open "SELECT * FROM VVendOrdersToRcv WHERE PONo='" & cmbOrderNo.ID & "' ORDER BY GroupID,MaterialName", con, adOpenForwardOnly, adLockReadOnly
        Dim ITM As ListItem
        LV.ListItems.Clear
        
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, , "")
            ITM.Tag = 0 'Val(![WORef] & "")
            
            'Display the Customer Order No if chkOrderNo is checked Else Show the Company Order No
            ITM.ListSubItems.add(, , ![MaterialID]).Tag = Val(!Rate & "")
            ITM.ListSubItems.add , , ![MaterialName] & ""
            ITM.ListSubItems.add(, , Val(![QtyOrdered]) & " " & ![Unit]).Tag = Abs(!CCItem)
            ITM.ListSubItems.add , , Val(![QtyToRcv]) & " " & ![Unit]
            ITM.ListSubItems.add().Tag = ![Unit] & ""
            ITM.ListSubItems.add().Tag = ""
            ITM.ListSubItems.add(, , Val(!Rate & "")).Tag = Val(!Rate & "")
            'Set Tool Tips
            
            ITM.ListSubItems(4).Tag = Val(!OtherUnitEntryID & "")
            
            For i = 1 To ITM.ListSubItems.count
                ITM.ListSubItems(i).ToolTipText = ITM.ListSubItems(i).Text
            Next
            
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    cmbMaterial.ClearVals
    cmbMaterial.AddVals con, "'{' + MaterialID + '} ' + MaterialName", "VVendOrdersToRcv", "MaterialID", " WHERE PONo='" & cmbOrderNo.ID & "' ORDER BY GroupID,MaterialName"
    
    
Exit Sub

err:
   MsgBox err.Description
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdSave_Click()
    
    If Saved Then
        Unload Me
        Me.Show
    End If
    
End Sub

Private Sub cmdSaveClose_Click()
    
    If Saved Then
        Unload Me
    End If
    
End Sub


Private Sub cmdCancelFB_Click()
    FFB.Visible = False
    LV.Enabled = True
    LV.SetFocus
End Sub

Private Sub LV_DblClick()
   
    If LV.ListItems.count = 0 Then Exit Sub
    
    LV.Enabled = False

    txtItemCode = LV.SelectedItem.Text
    Call ShowFeedback(Val(LV.SelectedItem.Tag))
     
    FFB.Visible = True
    
    
End Sub

Private Sub ShowFeedback(lVRD_RefID As Long)

    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM VendRcvdDetailPO_Feedback WHERE VRDPO_RefID=" & lVRD_RefID, con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
        
            txtRemarks = !Remarks & ""
            txtQtyAccepted = Val(!QtyAccepted & "")
            txtQtyRejected = Val(!QtyRejected & "")
            
            txtFilePath = !PDFFileName & ""
            txtFilePath.Tag = !PDFFileName & ""
        Else
        
            txtFilePath = ""
            txtFilePath.Tag = ""
            txtQtyAccepted = ""
            txtQtyRejected = ""
            txtRemarks = ""
            
        End If
        .Close
    End With
    Set rs = Nothing
    
    txtRemarks.Tag = lVRD_RefID 'Val(!EntryID & "")
    
End Sub


Private Sub Form_Load()

    DT = GetServerDate
    
    
    For i = 1 To LV.ColumnHeaders.count
        LV.ColumnHeaders(i).Tag = LV.ColumnHeaders(i).Width
    Next
    
    
    strRcvNo = ""
        
    'lShelfNoForAutoPlacement = GetSingleLongValue("MAX(EntryID)", "StoreShelfs")
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    FAll.Left = (Me.ScaleWidth - FAll.Width) / 2
    FAll.Top = (Me.ScaleHeight - FAll.Height) / 2

 
End Sub

Private Function Saved() As Boolean

    On Error GoTo err
    Dim i As Integer, lEntryID As Long
    Call StartTrans(con)
    
    con.Execute "INSERT INTO VendRcvdDetailPO_Rejection(DT,UserName,MachineName,OrderNo) VALUES('" & DT.Value & "','" & _
     CurrentUserName & "','" & strComputerName & "','" & txtPONo.Text & "')"
    
    lEntryID = GetSingleLongValue("MAX(EntryID)", "VendRcvdDetailPO_Rejection", " WHERE MachineName='" & strComputerName & "'")
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If Val(.SubItems(3)) > 0 Then
                con.Execute "INSERT INTO VendRcvdDetailPO_Rejection_Detail(VRDPO_Rej_RefID,VRDPO_RefID,QtyRejected,SheetsRejected)" & _
                 " VALUES(" & lEntryID & "," & Val(.Tag) & "," & Val(.SubItems(3)) & "," & Val(.SubItems(4)) & ")"
            End If
        End With
    Next
    con.CommitTrans
    
    Unload Me
    Exit Function
err:
    MsgBox err.Description
End Function

Public Sub ShowMe(p_strRcvID As String)

    strRcvNo = p_strRcvID
    txtRcvNo = strRcvNo
    
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
    
        strSQL = "SELECT * FROM VVendRcvdDetailPO WHERE RcvID='" & strRcvNo & "' ORDER BY EntryID"
        .Open strSQL, con, adOpenForwardOnly, adLockReadOnly
            
        txtVender = !AccTitle & ""
        txtPONo = !PORefNo & ""
        
        LV.ListItems.Clear
        
        Do Until .EOF
            
            Set ITM = LV.ListItems.add(, , !MaterialID & " " & !RMName & "")
            ITM.Tag = !EntryID & ""
            
            ITM.ListSubItems.add , , Val(!QtyRcvd & "") 'Round(Val(![QtyPassed] & "") - Val(![QtyPlaced] & ""), 2)
             
            .MoveNEXT
        Loop
        .Close
        
    End With
    
    Me.Show
    
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If LV.ListItems.count = 0 Then Exit Sub
    If Button <> 2 Then Exit Sub
    
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtQty) > Val(LV.SelectedItem.SubItems(1)) Then
            MsgBox "Invalid Qty.", vbInformation
            Exit Sub
        End If
        
        If txtSheets.Enabled = False Then
            LV.SelectedItem.SubItems(3) = Val(txtQty)
            
            LV.SetFocus
            
            PicEdit.Visible = False
        Else
        
            With txtSheets
                .SelStart = 0
                .SelLength = Len(.Text)
                .SetFocus
            End With
            
        End If
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub

Private Sub txtSheets_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Val(txtSheets) > Val(LV.SelectedItem.SubItems(2)) Then
            MsgBox "Invalid Sheets.", vbInformation
            Exit Sub
        End If
        If txtSheets.Enabled = False Then
            ITM.SubItems(3) = Val(txtQty)
            ITM.SubItems(4) = Val(txtSheets)
            LV.SetFocus
            PicEdit.Visible = False
        Else
            txtSheets.SetFocus
        End If
        
    ElseIf KeyAscii = 27 Then
        PicEdit.Visible = False
    End If
    
End Sub


Private Sub cmdUpdateFB_Click()
   
    Dim lRecAffect As Long
    
    con.Execute "UPDATE VendRcvdDetailPO_Feedback SET QtyAccepted=" & Val(txtQtyAccepted) & ",QtyRejected=" & Val(txtQtyRejected) & ",Remarks='" & txtRemarks & "' WHERE VRDPO_RefID=" & Val(txtRemarks.Tag), lRecAffect
    
    If lRecAffect = 0 Then
        con.Execute "INSERT INTO VendRcvdDetailPO_Feedback(VRDPO_RefID,QtyAccepted,QtyRejected,Remarks,UserName,MachineName) VALUES(" & Val(txtRemarks.Tag) & "," & Val(txtQtyAccepted) & "," & Val(txtQtyRejected) & ",'" & txtRemarks & "','" & _
         CurrentUserName & "','" & strComputerName & "')"
         
   
    End If
           
    Dim lmyEntryID As Long
    lmyEntryID = GetSingleLongValue("EntryID", "VendRcvdDetailPO_Feedback", " WHERE VRDPO_RefID=" & Val(txtRemarks.Tag))
    Call AttachPDF(lmyEntryID)
    
    FFB.Visible = False
    LV.Enabled = True
    LV.SetFocus
    
End Sub

Private Sub AttachPDF(p_lEntryID As Long)
    
    On Error Resume Next
    Dim strPDFPath As String
    strPDFPath = txtFilePath
    If strPDFPath <> "" And txtFilePath.Tag = "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EntryID,PDF,PDFFileName FROM VendRcvdDetailPO_Feedback WHERE EntryID=" & p_lEntryID, con, adOpenStatic, adLockOptimistic
        If Rec.RecordCount > 0 Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile strPDFPath
            Rec.Fields(1) = rsPic.Read
            Rec.Fields(2) = strFileName
            rsPic.Close
            Rec.Update
        End If
        
        Rec.Close
        
        Set Rec = Nothing
        Set rsPic = Nothing
        
    End If

End Sub


Private Sub cmdBrowseBarcode_Click()

    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        txtFilePath = Picfile
        txtFilePath.Tag = ""
        strFileName = CD1.FileTitle
    End If

End Sub

