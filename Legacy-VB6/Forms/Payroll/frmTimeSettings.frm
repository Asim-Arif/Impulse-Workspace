VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmTimeSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Time Settings..."
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
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
   ScaleHeight     =   4035
   ScaleWidth      =   4890
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtMinSalary 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "4000"
      Top             =   3690
      Width           =   1200
   End
   Begin VB.TextBox txtComputer 
      Height          =   345
      Left            =   1995
      TabIndex        =   6
      Top             =   660
      Width           =   2250
   End
   Begin VB.TextBox txtDiff 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   3045
      TabIndex        =   3
      Top             =   255
      Width           =   1200
   End
   Begin MSComctlLib.ListView LV 
      Height          =   2100
      Left            =   60
      TabIndex        =   7
      Top             =   1050
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   3704
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   1
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
         Text            =   "SNo"
         Object.Width           =   1773
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Computer Name"
         Object.Width           =   6481
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Minimum Salary :"
      Height          =   195
      Index           =   4
      Left            =   90
      TabIndex        =   11
      Top             =   3735
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Rs."
      Height          =   195
      Index           =   3
      Left            =   2595
      TabIndex        =   10
      Top             =   3840
      Width           =   240
   End
   Begin MSForms.CommandButton cmdAdd 
      Height          =   390
      Left            =   4275
      TabIndex        =   8
      ToolTipText     =   "Click Here To Add Computer"
      Top             =   630
      Width           =   525
      PicturePosition =   65543
      Size            =   "926;688"
      Picture         =   "frmTimeSettings.frx":0000
      Accelerator     =   65
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Computers Running RFID :"
      Height          =   195
      Index           =   2
      Left            =   60
      TabIndex        =   5
      Top             =   735
      Width           =   1920
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mins"
      Height          =   195
      Index           =   1
      Left            =   4320
      TabIndex        =   4
      Top             =   405
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Difference Between In Time && Out Time :"
      Height          =   195
      Index           =   0
      Left            =   75
      TabIndex        =   2
      Top             =   300
      Width           =   2940
   End
   Begin MSForms.CommandButton cmdSave 
      Height          =   360
      Left            =   2010
      TabIndex        =   1
      Top             =   3240
      Width           =   1365
      Caption         =   "Save"
      PicturePosition =   327683
      Size            =   "2408;635"
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
      Height          =   360
      Left            =   3405
      TabIndex        =   0
      Top             =   3240
      Width           =   1365
      Caption         =   "Cancel   "
      PicturePosition =   327683
      Size            =   "2408;635"
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
Attribute VB_Name = "frmTimeSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IsNew As Boolean
Public add As Boolean
Dim FirstEntry As Boolean

Private Sub cmdAdd_Click()
    
    If txtComputer = "" Then
        MsgBox "Computer Name Could Not Be Null !", vbInformation
        txtComputer.SetFocus
        Exit Sub
    End If
        
    Dim ITM As ListItem
    Set ITM = LV.ListItems.add(, , LV.ListItems.Count + 1)
    ITM.ListSubItems.add , , txtComputer
    
End Sub

Private Sub cmdCancel_Click()
    
    Unload Me
    
End Sub

Private Function Saved() As Boolean
    
    Dim BGHoliday As Integer, AGHoliday As Integer, _
    BPHoliday As Integer, APHoliday As Integer

    'On Error GoTo ERR
    
    
    
    Dim SQL As String, LHours As Double, LunchTime As Double
    
    StartTrans con
    
    con.Execute "Update TimeSettings set TimeDiff=" & Val(txtDiff) & ""
    con.Execute "DELETE FROM Computers"
    
    For i = 1 To LV.ListItems.Count
        con.Execute "INSERT INTO Computers(CompName) VALUES('" & LV.ListItems(i).SubItems(1) & "')"
    Next
    con.CommitTrans
    
    Saved = True
        
    Exit Function
err:
    MsgBox err.Description
End Function

Private Sub cmdSave_Click()
    If Saved Then Unload Me
End Sub

Public Sub Editrec()
    
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "select * from TimeSettings", con, adOpenForwardOnly, adLockReadOnly
        If .RecordCount > 0 Then
            txtDiff = Val(!TimeDiff & "")
        End If
        .Close
    
        .Open "SELECT * FROM Computers", con, adOpenStatic, adLockReadOnly
        For i = 1 To .RecordCount
            Set ITM = LV.ListItems.add(, , LV.ListItems.Count + 1)
            ITM.ListSubItems.add , , !CompName & ""
            .MoveNext
        Next
    End With
    
    Set rs = Nothing
End Sub


Private Sub Form_Load()
    Call Editrec
End Sub



Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 46 Then
        LV.ListItems.Remove LV.SelectedItem.Index
    End If

End Sub

