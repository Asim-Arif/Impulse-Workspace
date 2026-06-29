VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAddQualAca 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Academic Qualification..."
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
      TabIndex        =   8
      Top             =   0
      Width           =   4755
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   5
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1470
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         Height          =   405
         Left            =   3345
         TabIndex        =   7
         Top             =   2265
         Width           =   1335
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Save && Close"
         Height          =   405
         Left            =   1950
         TabIndex        =   6
         Top             =   2265
         Width           =   1335
      End
      Begin VB.CommandButton cmdSavenNew 
         Caption         =   "Save && &New"
         Height          =   405
         Left            =   555
         TabIndex        =   5
         Top             =   2265
         Width           =   1335
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   1620
         TabIndex        =   4
         Top             =   1800
         Width           =   1470
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   4
         Left            =   90
         TabIndex        =   3
         Top             =   1800
         Width           =   1470
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
         TabIndex        =   10
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
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   495
         Width           =   600
      End
      Begin MSComCtl2.DTPicker DT 
         Height          =   315
         Left            =   3390
         TabIndex        =   2
         Top             =   1095
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "yyyy"
         Format          =   89718787
         UpDown          =   -1  'True
         CurrentDate     =   37683
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Percentage :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3195
         TabIndex        =   18
         Top             =   1575
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marks Obtained :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   1575
         Width           =   1230
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total Marks :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1650
         TabIndex        =   15
         Top             =   1575
         Width           =   930
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Passing Year :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3420
         TabIndex        =   14
         Top             =   900
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name Of Institute :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   105
         TabIndex        =   13
         Top             =   885
         Width           =   1395
      End
      Begin VB.Label lblCourses 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Diploma/Degree/Certificate :"
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
         MouseIcon       =   "frmAddQualAca.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   12
         Top             =   300
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sr No :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   105
         TabIndex        =   11
         Top             =   300
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmAddQualAca"
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
    
    If ValidateData = False Then
        Exit Sub
    End If
    
    RetVal = 1
    
    sSNo = txts(0)
    sDiploma = txts(1)
    sInstitute = txts(2)
    sstrYear = DT.year
    iTotalMarks = Val(txts(3))
    iMarksObtained = Val(txts(4))
    iPercentage = Val(txts(5).Tag)
    
    Unload Me
End Sub

Private Sub cmdSavenNew_Click()
    
    If ValidateData = False Then
        Exit Sub
    End If
    
    RetVal = 2
    
    sSNo = txts(0)
    sDiploma = txts(1)
    sInstitute = txts(2)
    sstrYear = DT.year
    iTotalMarks = Val(txts(3))
    iMarksObtained = Val(txts(4))
    iPercentage = Val(txts(5).Tag)
    
    Unload Me
    
End Sub

Private Function ValidateData() As Boolean

    If txts(1) = "" Then
        MsgBox "Enter Diploma/Degree/Certificate.", vbInformation
        Exit Function
    ElseIf txts(2) = "" Then
        MsgBox "Enter Institute Name.", vbInformation
        Exit Function
    ElseIf Val(txts(3)) = 0 Then
        MsgBox "Invalid Total Marks.", vbInformation
        Exit Function
    End If
    
    ValidateData = True
    
End Function
Private Sub Form_Load()
    AddToCombo cmbCourse, "Description", "AcaCourses", " Order By Description"
End Sub

Private Sub lblCourses_Click()

    Dim f As New frmMiscS_New
    f.ShowForm "AcaCourses", "Description", "Courses List", "Description"
    
    AddToCombo cmbCourse, "Description", "AcaCourses", " Order By Description"
    
End Sub

Private Sub txts_LostFocus(Index As Integer)

    txts(Index) = StrConv(txts(Index), vbProperCase)
    
End Sub

Private Sub txts_Validate(Index As Integer, cancel As Boolean)

    Select Case Index
        Case 3, 4
'            If Val(txts(4)) > Val(txts(3)) Then
'                MsgBox "Marks Obtained cannot be Greater than Total Marks.", vbInformation
'                Cancel = True
'            End If
            If Val(txts(3)) > 0 And Val(txts(4)) > 0 Then
                txts(5) = Round(((Val(txts(4)) / Val(txts(3))) * 100), 2) & " %"
                txts(5).Tag = Round(((Val(txts(4)) / Val(txts(3))) * 100), 2)
            Else
                txts(5) = 0
                txts(5).Tag = 0
            End If
    End Select
    
End Sub
