VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmVendorAssItemHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vendor Assigned Items"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8100
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
   ScaleHeight     =   6795
   ScaleWidth      =   8100
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
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
      Height          =   5565
      Left            =   105
      TabIndex        =   4
      Top             =   675
      Width           =   7890
      Begin VB.TextBox txtRemarksEdit 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3825
         TabIndex        =   12
         Top             =   3045
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4170
         Left            =   105
         TabIndex        =   5
         Top             =   1290
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   7355
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "S #"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "User Name"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Rate"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Remarks"
            Object.Width           =   2469
         EndProperty
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   4
         Left            =   6420
         TabIndex        =   11
         Tag             =   "CustCode"
         Top             =   930
         Width           =   1335
         VariousPropertyBits=   679495711
         MaxLength       =   9
         BorderStyle     =   1
         Size            =   "2355;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Current Rate"
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
         Left            =   6435
         TabIndex        =   10
         Top             =   705
         Width           =   1095
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Remarks"
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
         Index           =   3
         Left            =   135
         TabIndex        =   9
         Top             =   705
         Width           =   765
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   3
         Left            =   105
         TabIndex        =   8
         Tag             =   "CustCode"
         Top             =   930
         Width           =   6285
         VariousPropertyBits=   679495711
         MaxLength       =   9
         BorderStyle     =   1
         Size            =   "11086;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox TBox 
         Height          =   315
         Index           =   1
         Left            =   105
         TabIndex        =   7
         Tag             =   "CustCode"
         Top             =   345
         Width           =   7665
         VariousPropertyBits=   679495711
         MaxLength       =   9
         BorderStyle     =   1
         Size            =   "13520;556"
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Item"
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
         Index           =   2
         Left            =   135
         TabIndex        =   6
         Top             =   120
         Width           =   420
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   4710
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Employee's Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Vendor Name"
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
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   135
      Width           =   1125
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   2
      Left            =   105
      TabIndex        =   2
      Tag             =   "CustCode"
      Top             =   360
      Width           =   7860
      VariousPropertyBits=   679495711
      MaxLength       =   9
      BorderStyle     =   1
      Size            =   "13864;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   375
      Left            =   3930
      TabIndex        =   0
      Top             =   6315
      Visible         =   0   'False
      Width           =   2295
      Caption         =   " Update Remarks"
      PicturePosition =   327683
      Size            =   "4048;661"
      Picture         =   "frmVendorAssItemHistory.frx":0000
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   375
      Left            =   6285
      TabIndex        =   1
      Top             =   6315
      Width           =   1695
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2990;661"
      Picture         =   "frmVendorAssItemHistory.frx":044E
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
Attribute VB_Name = "frmVendorAssItemHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lEntryID As Long

Public Sub ViewHistory(p_lEntryID As Long)

    lEntryID = p_lEntryID
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT EntryID,RMID,RMID1,RMName,Rate,GroupName,Remarks,AccTitle FROM VVenderAssItems WHERE EntryID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
        TBox(2) = !AccTitle & ""
        TBox(1) = !RMID1 & ""
        TBox(3) = !Remarks & ""
        TBox(4) = Val(!Rate & "")
        .Close
        
        .Open "SELECT * FROM VenderAssItems_Revisions WHERE VAI_RefID=" & lEntryID & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !EntryID & "'", LV.ListItems.count + 1)
            ITM.ListSubItems.add , , Format(!DTEntry, "dd-MMM-yyyy HH:ss")
            ITM.ListSubItems.add , , !UserName & ""
            ITM.ListSubItems.add , , Val(!Rate & "")
            ITM.ListSubItems.add , , !Remarks & ""
            .MoveNEXT
        Loop
        .Close
        
    End With
    Set rs = Nothing
    
    Me.Show 1
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    con.Execute "UPDATE VenderAssItems_Revisions SET Remarks='" & TBox(3) & "' WHERE EntryID=" & lEntryID
End Sub

Private Sub LV_DblClick()
    Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    With txtRemarksEdit
        .Move LV.Left + LV.ColumnHeaders(5).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(5).Width
        .Text = LV.SelectedItem.SubItems(4)
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Private Sub txtRemarksEdit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(4) = txtRemarksEdit.Text
        con.Execute "UPDATE VenderAssItems_Revisions SET Remarks='" & txtRemarksEdit & "' WHERE EntryID=" & Val(LV.SelectedItem.key)
        txtRemarksEdit.Visible = False
    ElseIf KeyAscii = 27 Then
        txtRemarksEdit.Visible = False
    End If
    
End Sub

Private Sub txtRemarksEdit_LostFocus()
    txtRemarksEdit.Visible = False
End Sub
