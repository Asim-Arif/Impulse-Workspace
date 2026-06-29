VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmGenFormE 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form E"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGenFormE.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   4365
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ListView LV 
      Height          =   6570
      Left            =   45
      TabIndex        =   1
      Top             =   645
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   11589
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Form E No."
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "No."
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   2040
      TabIndex        =   6
      Top             =   75
      Width           =   255
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   1
      Left            =   1995
      TabIndex        =   5
      Tag             =   "ItemUsage"
      Top             =   300
      Width           =   960
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "1693;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdGen 
      Height          =   360
      Left            =   3045
      TabIndex        =   4
      Top             =   255
      Width           =   1305
      Caption         =   "Generate"
      PicturePosition =   327683
      Size            =   "2302;635"
      Accelerator     =   71
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Starting Form E No."
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   3
      Left            =   75
      TabIndex        =   3
      Top             =   75
      Width           =   1410
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   0
      Left            =   45
      TabIndex        =   2
      Tag             =   "ItemUsage"
      Top             =   300
      Width           =   1920
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "3387;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   2685
      TabIndex        =   0
      Top             =   7245
      Width           =   1665
      Caption         =   "Close            "
      PicturePosition =   327683
      Size            =   "2937;635"
      Picture         =   "frmGenFormE.frx":0442
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
Attribute VB_Name = "frmGenFormE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub RefreshList()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select FormENo From GeneratedFormENos Where Used=0", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            LV.ListItems.Add , , ![FormENo] & ""
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdGen_Click()

    On Error GoTo err
    
    If TBox(0) = "" Or TBox(1) = "" Then
        MsgBox "Invalid Entries.", vbInformation
        Exit Sub
    End If
    
    Call StartTrans(con)
    
    For i = 0 To Val(TBox(1)) - 1
        con.Execute "Insert Into GeneratedFormENos(FormENo) Values(" & Val(TBox(0)) + i & ")"
    Next i
    
    Call con.CommitTrans
    
    Call RefreshList
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub Form_Load()
    Call RefreshList
End Sub

Private Sub Form_Unload(Cancel As Integer)
    OurOwnForm = True
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error GoTo err
    
    If LV.ListItems.Count = 0 Then Exit Sub
    
    If KeyCode = 46 Then
        If MsgBox("Are You Sure To Delete This No.?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        Dim RecAff As Integer
        con.Execute "Delete From GeneratedFormENos Where FormENo='" & LV.SelectedItem.Text & "'", RecAff
        
        If RecAff = 1 Then
            LV.ListItems.Remove LV.SelectedItem.Index
        End If
        
    End If
    
    Exit Sub
err:
    MsgBox err.Description
End Sub
