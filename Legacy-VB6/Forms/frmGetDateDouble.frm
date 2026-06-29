VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGetDateDouble 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Date..."
   ClientHeight    =   1140
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGetDateDouble.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPrintEarlier 
      Caption         =   "Print &Earlier"
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   135
      TabIndex        =   4
      Top             =   780
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   135
      TabIndex        =   3
      Top             =   405
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dddd, dd-MMM-yyyy"
      Format          =   270860291
      CurrentDate     =   39372
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5205
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   5205
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   2610
      TabIndex        =   5
      Top             =   405
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dddd, dd-MMM-yyyy"
      Format          =   270860291
      CurrentDate     =   39372
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&To Date"
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
      Index           =   1
      Left            =   2595
      TabIndex        =   6
      Top             =   180
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&From Date"
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
      Index           =   0
      Left            =   150
      TabIndex        =   2
      Top             =   165
      Width           =   885
   End
End
Attribute VB_Name = "frmGetDateDouble"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim iReportOption As Integer
Dim bCancelled As Boolean
Dim bChecked As Boolean
Dim DT As Date
Dim DtTo As Date
Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub Form_Load()
    DTPicker1.Value = DateAdd("d", -7, Date)
    DTPicker2.Value = Date
    bChecked = False
End Sub

Private Sub OKButton_Click()
    bCancelled = False
    If chkPrintEarlier.Value = vbChecked Then
        bChecked = True
    Else
        bChecked = False
    End If
    DT = DTPicker1.Value
    DtTo = DTPicker2.Value
    Unload Me
End Sub

Public Function getDate(ByRef myDT As Date, ByRef myDTTo As Date, Optional ByRef m_bChecked As Boolean = False) As Boolean
     
    Me.Show 1
'    bPrintEarlier = bEarlier
    myDT = DT
    myDTTo = DtTo
    m_bChecked = bChecked
    getDate = Not bCancelled
    
End Function
