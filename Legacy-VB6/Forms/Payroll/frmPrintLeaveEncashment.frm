VERSION 5.00
Begin VB.Form frmPrintLeaveEncashment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Annual Leave Encashment"
   ClientHeight    =   1560
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3045
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   360
      Left            =   1740
      TabIndex        =   3
      Top             =   1050
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      Height          =   360
      Left            =   465
      TabIndex        =   2
      Top             =   1050
      Width           =   1215
   End
   Begin VB.ComboBox cmbPeriod 
      Height          =   315
      Left            =   165
      TabIndex        =   0
      Top             =   600
      Width           =   2790
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Period"
      Height          =   195
      Left            =   165
      TabIndex        =   1
      Top             =   375
      Width           =   450
   End
End
Attribute VB_Name = "frmPrintLeaveEncashment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    
    On Error GoTo err
    If cmbPeriod.ListIndex = -1 Then Exit Sub
    
    Me.MousePointer = vbHourglass
    Dim cmd As New ADODB.command
    With cmd
        .CommandTimeout = 120
        Set .ActiveConnection = con
        .CommandType = adCmdStoredProc
        .CommandText = "SP_EmpLeaeveEncashment"
        .Parameters("@StartDate").value = DateSerial(Left(cmbPeriod.Text, 4), 7, 1)
        .Parameters("@EndDate").value = DateSerial(Right(cmbPeriod.Text, 4), 6, 30)
        .Execute
    End With
    
    Dim rpt As CRAXDDRT.Report, f As New frmPrevRpt
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptLeaveEncashment.rpt")
    rpt.FormulaFields(3).Text = "'" & Left(cmbPeriod.Text, 4) & " - " & Right(cmbPeriod.Text, 4) & "'"
    f.ShowReport "", rpt
    
    Me.MousePointer = vbNormal
    
    Exit Sub
err:
    Me.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Form_Load()

    Dim i As Integer
    For i = 2005 To year(Date)
        cmbPeriod.AddItem i & " * " & i + 1
    Next
    
End Sub
