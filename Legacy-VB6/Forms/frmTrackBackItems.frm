VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmTrackBackItems 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "fv"
   ClientHeight    =   10560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTrackBackItems.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10560
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10965
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   10110
      Width           =   1215
   End
   Begin VB.PictureBox PicWC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillColor       =   &H00005601&
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00005601&
      Height          =   9495
      Index           =   0
      Left            =   60
      ScaleHeight     =   9495
      ScaleWidth      =   12120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   570
      Width           =   12120
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Height          =   9495
         Left            =   30
         TabIndex        =   3
         Top             =   -75
         Width           =   12060
         Begin ComboList.Usercontrol1 cmbInvoice 
            Height          =   330
            Left            =   60
            TabIndex        =   4
            Top             =   360
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   582
            BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ListIndex       =   -1
            Appearance      =   0
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Master POs :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5910
            Left            =   4245
            TabIndex        =   10
            Top             =   3495
            Width           =   4140
            Begin MSComctlLib.ListView LVMasterPOs 
               Height          =   5550
               Left            =   75
               TabIndex        =   11
               Top             =   285
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   9790
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   15983848
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Master PO"
                  Object.Width           =   2469
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  SubItemIndex    =   1
                  Text            =   "Maker"
                  Object.Width           =   3528
               EndProperty
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Lots :"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5910
            Left            =   60
            TabIndex        =   8
            Top             =   3495
            Width           =   4140
            Begin MSComctlLib.ListView LVLots 
               Height          =   5550
               Left            =   75
               TabIndex        =   9
               Top             =   270
               Width           =   3975
               _ExtentX        =   7011
               _ExtentY        =   9790
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   -1  'True
               FullRowSelect   =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   15983848
               BorderStyle     =   1
               Appearance      =   0
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Calibri"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   2
               BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Text            =   "Lot No."
                  Object.Width           =   2822
               EndProperty
               BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
                  Alignment       =   2
                  SubItemIndex    =   1
                  Text            =   "Qty"
                  Object.Width           =   2117
               EndProperty
            End
         End
         Begin MSComctlLib.ListView LVItems 
            Height          =   2730
            Left            =   60
            TabIndex        =   7
            Top             =   750
            Width           =   11910
            _ExtentX        =   21008
            _ExtentY        =   4815
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   15784654
            BorderStyle     =   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Item Code"
               Object.Width           =   3528
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Item Name"
               Object.Width           =   14843
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Qty"
               Object.Width           =   2117
            EndProperty
         End
         Begin MSForms.CommandButton cmdRefresh 
            Height          =   420
            Left            =   3570
            TabIndex        =   6
            Top             =   300
            Width           =   1125
            ForeColor       =   65280
            BackColor       =   4210752
            Caption         =   "Refresh        "
            PicturePosition =   327683
            Size            =   "1984;741"
            Accelerator     =   82
            FontName        =   "Calibri"
            FontEffects     =   1073741825
            FontHeight      =   195
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice No."
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   60
            TabIndex        =   5
            Top             =   150
            Width           =   900
         End
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TRACKBACK ITEMS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   12120
   End
End
Attribute VB_Name = "frmTrackBackItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bMouseDown As Boolean
Const lCaptionHeight = 600
Dim lDownX As Long, lDownY As Long
Dim iColNo As Integer
Dim iColNoTemp As Integer
Dim bSaved As Boolean
Dim lEntryID As Long

Private Sub DrawBorder(Pic As PictureBox)

    Dim lBorderColor As Long
    lBorderColor = RGB(220, 219, 244) 'RGB(0, 200, 0)    'RGB(69, 140, 201) 'RGB(69, 140, 201)
    Pic.Line (0, 0)-(Pic.Width, 0), lBorderColor 'Top
    Pic.Line (0, Pic.Height - 85)-(Pic.Width, Pic.Height - 85), lBorderColor 'Bottom
    Pic.Line (0, 0)-(0, Pic.Height), lBorderColor    'Left
    Pic.Line (Pic.Width - 85, 0)-(Pic.Width - 85, Pic.Height), lBorderColor  'Right
    
End Sub

Private Sub DrawPicCaptions()

    Dim i As Integer
    Dim lBorderColor As Long, lForeColor As Long
    For i = 0 To PicWC.count - 1
        If i = 1 Then
            lBorderColor = RGB(153, 204, 255)
            lForeColor = 0
        ElseIf i = 0 Then
            lBorderColor = RGB(220, 219, 244) 'RGB(255, 204, 153)
            lForeColor = 0
        ElseIf i = 3 Then
            lBorderColor = RGB(255, 153, 204)
            lForeColor = 0
        ElseIf i = 2 Then
            lBorderColor = RGB(0, 255, 0)
            lForeColor = 0
        ElseIf i = 4 Then
            lBorderColor = RGB(200, 200, 192)
            lForeColor = 0
        End If
        Call DrawCaptions(PicWC(i), lBorderColor, lForeColor)
    Next
    
End Sub

Private Sub DrawCaptions(Pic As PictureBox, lBorderColor As Long, lForeColor As Long)

    Dim lFillColor
    'lBorderColor = 0 'RGB(226, 103, 161) 'RGB(70, 70, 181)
    lFillColor = lBorderColor 'RGB(226, 103, 161) 'RGB(70, 70, 181)
    Pic.Line (0, 20)-(Pic.Width, 20), lBorderColor
    
    If Pic.Tag <> "" Then
        Pic.Line (0, 150)-(Pic.Width, 150), lBorderColor
        
        Pic.Line (0, 50)-(Pic.Width, 300), lFillColor, BF
        Pic.ForeColor = lForeColor 'vbGreen
        Pic.CurrentX = (Pic.Width - Pic.TextWidth(Pic.Tag)) / 2
        Pic.CurrentY = 0
        Pic.Print Pic.Tag
    End If
        
    'Pic.Line (0, Pic.Height - 30)-(Pic.Width, Pic.Height - 30), lBorderColor 'Bottom
    Pic.Line (0, Pic.Height - 50)-(Pic.Width, Pic.Height), lFillColor, BF
    
    Pic.Line (20, 0)-(20, Pic.Height), lBorderColor                          'Left
    Pic.Line (Pic.Width - 20, 0)-(Pic.Width - 30, Pic.Height), lBorderColor  'Right
    
    'Pic.Line (0, Pic.Height - 500)-(50, Pic.Height), lBorderColor
'    Me.BackColor = lBorderColor
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()

    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT CompItemCode,ItemName,SUM(Qty) AS Qty FROM VCustomInvoice WHERE CustomInvoice='" & cmbInvoice.Text & "' GROUP BY CompItemCode,ItemName", con, adOpenForwardOnly, adLockReadOnly
        LVItems.ListItems.Clear
        Do Until .EOF
            Set ITM = LVItems.ListItems.add(, , !CompItemCode & "")
            ITM.ListSubItems.add , , !ItemName & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing

End Sub

Private Sub Form_Load()

    Call DrawPicCaptions
    bMouseDown = False
    bSaved = False
    lEntryID = 0
        
    cmbInvoice.ListHeight = 4000
    cmbInvoice.AddVals con, "CustomInvoice", "CustomInvoice", "CustomInvoice"
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = 1 Then
        If bMouseDown = False Then
            bMouseDown = True
            lDownX = x
            lDownY = y
        End If
    End If

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If bMouseDown Then
        Me.Move Me.Left + ((x - lDownX) / 15), Me.Top + ((y - lDownY) / 15)
    End If
End Sub

Private Sub Label2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseDown = False
End Sub

Private Sub LVItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        'Show Lots
        .Open "SELECT LotNo,SUM(DLD_VRD_Qty) AS Qty FROM VCustomPList_DispatchListDetail WHERE CustomInvoice='" & cmbInvoice.Text & "' AND CompItemCode='" & Item.Text & "' GROUP BY LotNo", con, adOpenForwardOnly, adLockReadOnly
        LVLots.ListItems.Clear
        Do Until .EOF
            Set ITM = LVLots.ListItems.add(, , !LotNo & "")
            ITM.ListSubItems.add , , Val(!Qty & "")
            .MoveNext
        Loop
        .Close
        
        .Open "SELECT DISTINCT VenderName,MasterPONo FROM VLotsWithMasterPONo INNER JOIN VCustomPList_DispatchListDetail ON " & _
         "VCustomPList_DispatchListDetail.LotNo=VLotsWithMasterPONo.LotNo WHERE CustomInvoice='" & cmbInvoice.Text & _
         "' AND CompItemCode='" & Item.Text & "' AND MasterPONo IS NOT NULL", con, adOpenForwardOnly, adLockReadOnly
         
        LVMasterPOs.ListItems.Clear
        Do Until .EOF
            Set ITM = LVMasterPOs.ListItems.add(, , !MasterPONo & "")
            ITM.ListSubItems.add , , !VenderName & ""
            .MoveNext
        Loop
        .Close
        
        
    End With
    Set rs = Nothing
    
End Sub

Private Sub LVLots_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button <> 2 Then Exit Sub
    If LVLots.ListItems.count = 0 Then Exit Sub
    
    Dim strRet As String
    strRet = ShowPopUpMenu1("Print PTC")
    If strRet = "Print PTC" Then
        Call PrintPTC(LVLots.SelectedItem.Text)
    End If
    
End Sub

Private Sub LVMasterPOs_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    If LVMasterPOs.ListItems.count = 0 Then Exit Sub
    Dim strRet As String
    strRet = ShowPopUpMenu1("Master PO Status")
    If strRet = "Master PO Status" Then
        Call PrintMasterPOStatus(LVMasterPOs.SelectedItem.Text)
    End If
    
End Sub

Private Sub PicWC_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If y < 600 Then
        If Button = 1 Then
            If bMouseDown = False Then
                bMouseDown = True
                lDownX = x
                lDownY = y
            End If
        End If
    End If
End Sub

Private Sub PicWC_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If bMouseDown Then
        Me.Move Me.Left + ((x - lDownX) / 15), Me.Top + ((y - lDownY) / 15)
    End If
End Sub

Private Sub PicWC_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    bMouseDown = False
End Sub

Public Function ShowMe() As Boolean
    
    'Will hold the handle to the new region.
    Dim hRegion As Long
    '
    'This API can create a Round or Oval Shaped Region.
'    hRegion = CreateRoundRectRgn(0, 0, Me.Width / 15, Me.Height / 15, 20, 20)
    '
    'Change the shape of the Form to the newly created Region.
'    SetWindowRgn Me.hwnd, hRegion, True
    '
    'Show the new form with the specified shape.
    'frmRegion.Show
   
    Me.Show
    ShowMe = bSaved
    
End Function
