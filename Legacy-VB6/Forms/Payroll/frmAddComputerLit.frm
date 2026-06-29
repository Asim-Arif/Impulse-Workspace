VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAddComputerLit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Computer Literacy....."
   ClientHeight    =   2775
   ClientLeft      =   7170
   ClientTop       =   5715
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2745
      Left            =   90
      TabIndex        =   9
      Top             =   0
      Width           =   4755
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   6
         Left            =   90
         TabIndex        =   3
         Top             =   1770
         Width           =   1470
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   1620
         TabIndex        =   4
         Top             =   1770
         Width           =   1470
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1770
         Width           =   1485
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   3360
         TabIndex        =   2
         Top             =   1095
         Width           =   1275
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "C&lose"
         Height          =   405
         Left            =   3345
         TabIndex        =   8
         Top             =   2265
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Save && &Close"
         Height          =   405
         Left            =   1950
         TabIndex        =   7
         Top             =   2265
         Width           =   1335
      End
      Begin VB.CommandButton cmdSavenNew 
         Caption         =   "&Save && New"
         Height          =   405
         Left            =   555
         TabIndex        =   6
         Top             =   2265
         Width           =   1335
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   90
         TabIndex        =   1
         Top             =   1095
         Width           =   3240
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   2865
         TabIndex        =   11
         Top             =   195
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   495
         Width           =   600
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Marks :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1650
         TabIndex        =   18
         Top             =   1545
         Width           =   930
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marks Obtained :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   1545
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3195
         TabIndex        =   16
         Top             =   1545
         Width           =   930
      End
      Begin MSForms.ComboBox cmbCourse 
         Height          =   300
         Left            =   750
         TabIndex        =   0
         Top             =   510
         Width           =   3885
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "6853;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   1
         BorderColor     =   12164479
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Period :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3420
         TabIndex        =   15
         Top             =   900
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name Of Institute :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   105
         TabIndex        =   14
         Top             =   885
         Width           =   1395
      End
      Begin VB.Label lblCourses 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Course Title"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   780
         MouseIcon       =   "frmAddComputerLit.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sr No :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   105
         TabIndex        =   12
         Top             =   300
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmAddComputerLit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RetVal As Integer
Dim sSNo As String, sDiploma As String, sInstitute As String
Dim sstrYear As String, iTotalMarks As Integer, iMarksObtained As Integer, iPercentage As Integer

Public Function ShowMe(ByRef SNo As String, ByRef Diploma As String, ByRef Institute As String, ByRef strYear As String, ByRef TotalMarks As Integer, ByRef MarksObtained As Integer, ByRef Percentage As Integer) As Integer
    
    txts(0) = SNo
    Me.Show 1
    
    SNo = sSNo
    Diploma = sDiploma
    Institute = sInstitute
    strYear = sstrYear
    TotalMarks = iTotalMarks
    MarksObtained = iMarksObtained
    Percentage = iPercentage
    
    ShowMe = RetVal
        
End Function

Private Sub cmbCourse_Change()
    Call cmbCourse_Click
End Sub

Private Sub cmbCourse_Click()
    txts(1) = cmbCourse.Text
End Sub

Private Sub cmdCancel_Click()
    RetVal = 0
    Unload Me
End Sub

Private Sub cmdOK_Click()

    RetVal = 1
    
    sSNo = txts(0)
    sDiploma = txts(1)
    sInstitute = txts(2)
    sstrYear = txts(5)
    iTotalMarks = Val(txts(4))
    iMarksObtained = Val(txts(6))
    iPercentage = Val(txts(3))
    
    Unload Me
    
End Sub

Private Sub cmdSavenNew_Click()

    RetVal = 2
    
    sSNo = txts(0)
    sDiploma = txts(1)
    sInstitute = txts(2)
    sstrYear = txts(5)
    iTotalMarks = Val(txts(4))
    iMarksObtained = Val(txts(6))
    iPercentage = Val(txts(3))
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    AddToCombo cmbCourse, "Description", "ComputerCourses", " Order By Description"
End Sub

Private Sub lblCourses_Click()

    Dim f As New frmMiscS_New
    f.ShowForm "ComputerCourses", "Description", "Courses List", "EntryID"
    
    AddToCombo cmbCourse, "Description", "ComputerCourses", " Order By Description"
    
End Sub

Private Sub txts_Validate(Index As Integer, cancel As Boolean)

    Select Case Index
        Case 4, 6
            If Val(txts(4)) > 0 And Val(txts(6)) > 0 Then
                txts(3) = Round(((Val(txts(6)) / Val(txts(4))) * 100), 2)
            Else
                txts(3) = 0
            End If
    End Select
    
End Sub

