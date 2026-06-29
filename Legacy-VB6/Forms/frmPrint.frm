VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "   Print Reports...."
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   60
      Width           =   3390
      _ExtentX        =   5980
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
   End
   Begin Crystal.CrystalReport Cr1 
      Left            =   75
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ComboList.Usercontrol1 cmbJob 
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   390
      Width           =   3390
      _ExtentX        =   5980
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
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   2910
      Picture         =   "frmPrint.frx":0442
      Top             =   750
      Width           =   480
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   360
      Left            =   2955
      TabIndex        =   3
      Top             =   1305
      Width           =   1515
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2672;635"
      Picture         =   "frmPrint.frx":0A31
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   360
      Left            =   1395
      TabIndex        =   2
      Top             =   1305
      Width           =   1515
      Caption         =   "    Print          "
      PicturePosition =   327683
      Size            =   "2672;635"
      Picture         =   "frmPrint.frx":0B43
      Accelerator     =   80
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
      Caption         =   "Job ID :"
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
      Left            =   390
      TabIndex        =   5
      Top             =   420
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Customer :"
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
      Left            =   105
      TabIndex        =   4
      Top             =   75
      Width           =   915
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rpt As CRAXDDRT.Report
Public SelFor As String
Public JobField As String
Public JobIDField As String
Public JobTable As String
Public JobCondField As String
Public TheEnd As String

Private Sub cmbCust_matched()

    cmbJob.ClearVals
    If cmbCust.ID = "0" Then
        cmbJob.AddVals con, JobField, JobTable, JobIDField
    Else
        cmbJob.AddVals con, JobField, JobTable, JobIDField, " Where " & JobCondField & "='" & cmbCust.ID & "'"
    End If
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub



Private Sub cmdPrint_Click()
    
    On Error GoTo err
    
    If cmbJob.MatchFound = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    Dim TempSelFor As String
    TempSelFor = SelFor & cmbJob.ID & TheEnd & " "
    Load frmPrevRpt
    frmPrevRpt.ShowReport TempSelFor, rpt
    
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
    
End Sub

Private Sub Form_Load()
    cmbCust.ListHeight = 1500
    cmbJob.ListHeight = 1200
    TheEnd = "'"
End Sub

Private Sub Form_Resize()
    Me.Move (MainForm.Width - Me.Width) / 2, (MainForm.Height - Me.Height) / 2 - 1200
    Me.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    OurOwnForm = True
End Sub
