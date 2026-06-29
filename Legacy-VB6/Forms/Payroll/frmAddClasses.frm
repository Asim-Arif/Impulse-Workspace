VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAddClasses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Classes"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5580
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4605
      Left            =   135
      TabIndex        =   2
      Top             =   855
      Width           =   5325
      Begin MSComctlLib.ListView Lv 
         Height          =   4290
         Left            =   135
         TabIndex        =   3
         Top             =   225
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   7567
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "EducationLevel"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Class Name"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.TextBox txtClassName 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2295
      TabIndex        =   1
      Top             =   450
      Width           =   1950
   End
   Begin VB.Label Label2 
      Caption         =   "Class Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2250
      TabIndex        =   6
      Top             =   225
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Education Level "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   180
      TabIndex        =   5
      Top             =   180
      Width           =   1365
   End
   Begin MSForms.CommandButton cmdAdd 
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   450
      Width           =   1140
      Caption         =   "Add"
      Size            =   "2011;661"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
   Begin MSForms.ComboBox cmbEducationLevel 
      Height          =   330
      Left            =   180
      TabIndex        =   0
      Top             =   450
      Width           =   2085
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3678;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
End
Attribute VB_Name = "frmAddClasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

End Sub

Private Sub cmdAdd_Click()
    
    Dim itm As ListItem
    If txtClassName.Text = "" Then
        MsgBox "Enter the class name !", vbInformation
        Exit Sub
    End If
    con.Execute "insert into ClassesDetail(EducationLevel,ClassName) values('" & cmbEducationLevel.Text & "','" & txtClassName.Text & "')"
    Call RefreshList
    txtClassName.Text = ""
    txtClassName.SetFocus
    
End Sub

Private Sub Form_Load()
    
    With cmbEducationLevel
        .AddItem ("Primary")
        .AddItem ("Middle")
        .AddItem ("Matric")
        .AddItem ("Intermediate")
        .AddItem ("Graducation")
        .AddItem ("Post Graducate")
    End With
    Call RefreshList
End Sub
Private Sub RefreshList()
    
    Dim itm As ListItem
    Dim rs As New ADODB.Recordset
    If rs.State Then rs.Close
    rs.Open "select * from ClassesDetail", con, adOpenForwardOnly, adLockReadOnly
    Lv.ListItems.Clear
    While Not rs.EOF
        Set itm = Lv.ListItems.add(, , rs.Fields("EducationLevel"))
        itm.ListSubItems.add , , rs.Fields("ClassName")
        rs.MoveNEXT
    Wend
    
End Sub
