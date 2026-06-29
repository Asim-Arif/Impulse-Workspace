VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMakeEmployeeInActive 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Make Employee In-Active"
   ClientHeight    =   2760
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMakeEmployeeInActive.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRemarks 
      Height          =   1560
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   4395
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   135
      TabIndex        =   0
      Top             =   405
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dddd, dd-MMM-yyyy"
      Format          =   78577667
      CurrentDate     =   39372
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4590
      TabIndex        =   3
      Top             =   2265
      Width           =   1335
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4590
      TabIndex        =   2
      Top             =   1815
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   135
      TabIndex        =   5
      Top             =   870
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   150
      TabIndex        =   4
      Top             =   165
      Width           =   405
   End
End
Attribute VB_Name = "frmMakeEmployeeInActive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim iReportOption As Integer
Dim bCancelled As Boolean
'Dim bEarlier As Boolean
Dim DT As Date
Dim strRemarks As String
Dim strEmpID As String

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
End Sub

Private Sub OKButton_Click()

    bCancelled = False
    If Trim(txtRemarks) = "" Then
        MsgBox "Please Enter Remarks.", vbInformation
        Exit Sub
    End If
    Dim dtMax As Date
    dtMax = GetSingleDateValue("MAX(DT)", "AttendanceSheet", " WHERE EmpID='" & strEmpID & "'")
    If dtMax > DTPicker1.Value Then
        MsgBox "Can't Make Employee In-Active, Last Attendance Activity Found on " & Format(dtMax, "dd-MMM-yyyy"), vbInformation
        Exit Sub
    End If
    
    DT = DTPicker1.Value
    strRemarks = txtRemarks.Text
    
    Unload Me
    
End Sub

Public Function getDate(p_strEmpID As String, ByRef myDT As Date, ByRef p_strRemarks As String) As Boolean
    
    strEmpID = p_strEmpID
    Me.Show 1
'    bPrintEarlier = bEarlier
    myDT = DT
    p_strRemarks = strRemarks
    getDate = Not bCancelled
    
End Function
