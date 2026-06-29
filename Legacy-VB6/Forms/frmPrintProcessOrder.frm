VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPrintProcessOrder 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Print Process Order"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtProcess 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   945
      Width           =   5850
   End
   Begin VB.TextBox txtMaker 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   345
      Width           =   5850
   End
   Begin MSComCtl2.DTPicker DTIssuance 
      Height          =   315
      Left            =   285
      TabIndex        =   3
      Top             =   1710
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   52494339
      CurrentDate     =   39701
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Process Order"
      Height          =   390
      Left            =   2460
      TabIndex        =   1
      Top             =   2160
      Width           =   1845
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   4335
      TabIndex        =   0
      Top             =   2160
      Width           =   1845
   End
   Begin MSComCtl2.DTPicker DTReturn 
      Height          =   315
      Left            =   2070
      TabIndex        =   4
      Top             =   1710
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   52494339
      CurrentDate     =   39701
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Process :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   300
      TabIndex        =   8
      Top             =   735
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Maker :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   300
      TabIndex        =   6
      Top             =   135
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Return Date :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2115
      TabIndex        =   5
      Top             =   1485
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Issuance Date :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   330
      TabIndex        =   2
      Top             =   1485
      Width           =   1305
   End
End
Attribute VB_Name = "frmPrintProcessOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lEntryID As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error Resume Next
    
     
    Dim rpt As CRAXDDRT.Report
'    If Format(DT.Value, "dddd") = "Sunday" Then
'        MsgBox "Invalid Date, Can't Accept Sunday.", vbInformation
'        Me.Show
'        Exit Sub
'    End If
    
'    If iBankID = 7 Then
'        Set rpt = rptApp.OpenReport(rptAccountsPath & "\rptABNAmroChq.rpt")
'    ElseIf iBankID = 8 Or iBankID = 9 Then
'        Set rpt = rptApp.OpenReport(rptAccountsPath & "\rptUBLChq.rpt")
'    Else
'        MsgBox "Bank Not Supported.", vbInformation
'        Exit Sub
'    End If
    
    con.Execute "UPDATE VendIssdDetail SET ProcessOrderIssuanceDT='" & DTIssuance & "',ProcessOrderRcvingDT='" & DTReturn & "' WHERE RefID=" & lEntryID
    On Error GoTo err
    
    Set rpt = rptApp.OpenReport(RptPath & "\ProcessOrder.rpt")
    Dim f As New frmPrevRpt
    
    Me.Hide
    
    f.ShowReport "{VendIssued.EntryID}=" & lEntryID, rpt
    
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Sub ShowMe(p_lEntryID As Long)

    lEntryID = p_lEntryID
    
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "SELECT * FROM VVendIssdDetail WHERE RefID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
        txtMaker.Text = !VenderName & ""
        txtProcess.Text = GetSingleStringValue("Description", "Processes", " WHERE ProcessID=" & Val(!ProcessID & "")) '!Description & ""
        
        If IsNull(!ProcessOrderIssuanceDT) Then
            DTIssuance = Date
            DTReturn = Date
        Else
            DTIssuance = !ProcessOrderIssuanceDT
            DTReturn = !ProcessOrderRcvingDT
        End If
        .Close
    End With
    Set rs = Nothing
    
    Me.Show 1
    
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

