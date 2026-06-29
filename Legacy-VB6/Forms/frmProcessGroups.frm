VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmProcessGroups 
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10620
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form19"
   MDIChild        =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   10620
   WindowState     =   2  'Maximized
   Begin VB.Frame FAcc 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   8400
      Left            =   60
      TabIndex        =   0
      Top             =   -45
      Width           =   10515
      Begin MSComctlLib.ListView LV 
         Height          =   7665
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   10350
         _ExtentX        =   18256
         _ExtentY        =   13520
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   15199215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Group No."
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   60
         TabIndex        =   3
         Top             =   7845
         Width           =   7620
         Begin VB.TextBox txtFind 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   1965
            TabIndex        =   6
            Top             =   135
            Width           =   4530
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6540
            TabIndex        =   5
            Top             =   135
            Width           =   1020
         End
         Begin VB.ComboBox cmbfield 
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "frmProcessGroups.frx":0000
            Left            =   60
            List            =   "frmProcessGroups.frx":0007
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   135
            Width           =   1875
         End
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   3180
         Top             =   4980
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   11
         ImageHeight     =   10
         MaskColor       =   255
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProcessGroups.frx":0018
               Key             =   "Down"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmProcessGroups.frx":01D4
               Key             =   "Up"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3180
         Top             =   4980
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin MSForms.CommandButton cmdClose 
         Height          =   375
         Left            =   9075
         TabIndex        =   2
         Top             =   7950
         Width           =   1335
         ForeColor       =   0
         Caption         =   "Close     "
         PicturePosition =   327683
         Size            =   "2355;661"
         Accelerator     =   67
         FontName        =   "Segoe UI"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Pop Up Menu"
      Begin VB.Menu mnuNew 
         Caption         =   "New Group"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit Group"
      End
      Begin VB.Menu mnuCopyGroupDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopyGroup 
         Caption         =   "Copy Group"
      End
      Begin VB.Menu mnuDash 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintAssignedItems 
         Caption         =   "Print Assigned Items"
      End
      Begin VB.Menu mnuDashPrint 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpdateAll 
         Caption         =   "Update All Linked Items"
      End
   End
End
Attribute VB_Name = "frmProcessGroups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strSQLRpt As String
Dim PrevIndex As Long

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdFind_Click()

    Dim StartIdx As Integer
    
    If LV.SelectedItem Is Nothing Then
       StartIdx = 1
    Else
       If LV.SelectedItem.Index = LV.ListItems.count Then
          StartIdx = 1
       Else
          StartIdx = LV.SelectedItem.Index
       End If
    End If

    Dim fldValue As String

    For i = StartIdx To LV.ListItems.count
       
'        If cmbfield.ListIndex = 3 Then
'            fldValue = LV.ListItems(i).SubItems(8)
'        Else
'            fldValue = LV.ListItems(i).SubItems(cmbfield.ListIndex)
'        End If
        fldValue = LV.ListItems(i).SubItems(1)
        If LCase(fldValue) Like "*" & LCase(Trim(txtFind)) & "*" Then
            LV.ListItems(i).Selected = True
            LV.ListItems(i).EnsureVisible
            LV.SetFocus
            If i <> StartIdx Then Exit For
        End If
       
    Next

    If i = LV.ListItems.count + 1 Then
        MsgBox "Record Not Found.", vbInformation
        Set LV.SelectedItem = Nothing
    End If
    
End Sub

Public Sub cmdRefresh_Click()
    Call RefreshLV
End Sub


Private Sub Form_Load()
    
    mnuPop.Visible = False
    RefreshLV
    
End Sub

Private Sub RefreshLV()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim Cond As String
    Dim ITM As ListItem
    
    Screen.MousePointer = vbHourglass
  
    With rs
        .Open "SELECT * FROM ProcessGroups" & Cond & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, ![EntryID] & "'", LV.ListItems.count + 1)
            ITM.ListSubItems.add , , !GroupName & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Screen.MousePointer = vbDefault
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    FAcc.Move (Me.Width - FAcc.Width) / 2, (Me.Height - FAcc.Height) / 2
End Sub

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    
    Load frmNewProcessGroup
    
    Call frmNewProcessGroup.EditItem(Val(LV.SelectedItem.key))
    
End Sub


Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then
        mnuEdit.Enabled = False
    Else
        mnuEdit.Enabled = True
    End If
    Me.PopUpMenu mnuPop
End Sub

Private Sub mnuAdd_Click()
    
End Sub

Private Sub mnuCopyGroup_Click()

    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim f As New frmNewProcessGroup
    Load f
    Call f.EditItem(Val(LV.SelectedItem.key), True)

End Sub

Private Sub mnuedit_Click()
    Call LV_DblClick
End Sub

Private Sub mnuNew_Click()
    Load frmNewProcessGroup
    frmNewProcessGroup.Show 1
    Call RefreshLV
End Sub

Private Sub mnuPrintAssignedItems_Click()
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ItemProcessGroups.rpt")
    Dim f As New frmPrevRpt
    f.ShowReport "{ProcessGroups.EntryID}=" & Val(LV.SelectedItem.key), rpt
End Sub

Private Sub mnuUpdateAll_Click()

    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Dim cmd As New ADODB.Command
    With cmd
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "UpdateItemProcessesFromProcessGroups_SP"
        .Parameters("@PG_RefID").Value = Val(LV.SelectedItem.key)
        .Execute
    End With
    MsgBox "Done.", vbInformation
    
End Sub

Private Sub txtFind_Change()
    cmdFind.Enabled = Len(Trim(txtFind)) > 0
End Sub

Private Sub txtFind_GotFocus()
   txtFind.SelStart = 0
   txtFind.SelLength = Len(txtFind)
   cmdFind.Default = True
End Sub

Private Sub txtFind_LostFocus()
   cmdFind.Default = False
End Sub
