VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmProcesses_Inspection_Points 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Process Inspection Points"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
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
   ScaleHeight     =   7110
   ScaleWidth      =   7155
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      ForeColor       =   &H00800000&
      Height          =   6255
      Left            =   105
      TabIndex        =   4
      Top             =   195
      Width           =   6915
      Begin VB.TextBox txtInspectionPoint 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   150
         TabIndex        =   8
         Top             =   1005
         Width           =   6615
      End
      Begin ComboList.Usercontrol1 cmbProcess 
         Height          =   285
         Left            =   150
         TabIndex        =   0
         Top             =   360
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   503
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBackColor   =   16777215
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin VB.TextBox txtSortNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2775
         TabIndex        =   6
         Top             =   3540
         Visible         =   0   'False
         Width           =   1000
      End
      Begin MSComctlLib.ListView LV 
         Height          =   4695
         Left            =   150
         TabIndex        =   1
         Top             =   1395
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   8281
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDropMode     =   1
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   8388608
         BackColor       =   16777215
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
         OLEDropMode     =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Inspection Point"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Inspection Point"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   7
         Top             =   750
         Width           =   1155
      End
      Begin VB.Label label1 
         AutoSize        =   -1  'True
         Caption         =   "Process"
         ForeColor       =   &H00400000&
         Height          =   195
         Index           =   9
         Left            =   165
         TabIndex        =   5
         Top             =   150
         Width           =   555
      End
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5910
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Employee's Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   6585
      Visible         =   0   'False
      Width           =   1695
      Caption         =   " Save            "
      PicturePosition =   327683
      Size            =   "2990;661"
      Picture         =   "frmProcesses_Inspection_Points.frx":0000
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
      Left            =   5340
      TabIndex        =   3
      Top             =   6585
      Width           =   1695
      Caption         =   "Close"
      PicturePosition =   327683
      Size            =   "2990;661"
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
Attribute VB_Name = "frmProcesses_Inspection_Points"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iProcessID As Integer

Public Sub RefreshList()

    On Error GoTo err
    Dim rs As New ADODB.Recordset, ITM As ListItem
    
    With rs
        .Open "SELECT * FROM Process_Inspection_Points WHERE ProcessID=" & cmbProcess.ID & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , !Point_Description & "")
            ITM.Tag = Val(!EntryID & "")
            .MoveNEXT
        Loop
        .Close
         
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub cmbProcess_matched()
    Call RefreshList
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    On Error GoTo err
    
    With cmbProcess
        .ListHeight = 3000
        .AddVals con, "Code + '-' + Description", "Processes", "ProcessID", " ORDER BY ProcessID"
    End With
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Do you want to Delete?", vbQuestion + vbYesNo) = vbYes Then
            Dim iRowAffect As Integer
            con.Execute "DELETE FROM Process_Inspection_Points WHERE EntryID=" & Val(LV.SelectedItem.Tag), iRowAffect
            If iRowAffect = 1 Then
                Call RefreshList
            End If
        End If
    End If
    
End Sub

Private Sub txtInspectionPoint_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmbProcess.MatchFound = False Then Exit Sub
        Dim iRowAffect As Integer
        con.Execute "INSERT INTO Process_Inspection_Points(ProcessID,Point_Description) VALUES(" & cmbProcess.ID & ",'" & txtInspectionPoint & "')", iRowAffect
        If iRowAffect = 1 Then
            Call RefreshList
            txtInspectionPoint = ""
        End If
    End If
End Sub

Public Sub ShowMe(p_iProcessID As Integer)
    iProcessID = p_iProcessID
    cmbProcess.ID = iProcessID
    Me.Show 1
End Sub
