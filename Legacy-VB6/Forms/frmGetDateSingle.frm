VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGetDateSingle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Date..."
   ClientHeight    =   1140
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
   Icon            =   "frmGetDateSingle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   6030
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
      Format          =   88211459
      CurrentDate     =   39372
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
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
      TabIndex        =   2
      Top             =   165
      Width           =   405
   End
End
Attribute VB_Name = "frmGetDateSingle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim iReportOption As Integer
Dim bCancelled As Boolean
'Dim bEarlier As Boolean
Dim DT As Date
Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
End Sub

Private Sub OKButton_Click()
    bCancelled = False
'    If chkPrintEarlier.value = vbChecked Then
'        bEarlier = True
'    Else
'        bEarlier = False
'    End If
    DT = DTPicker1.Value
    Unload Me
End Sub

Public Function getDate(ByRef myDT As Date) As Boolean
     
    Me.Show 1
'    bPrintEarlier = bEarlier
    myDT = DT
    getDate = Not bCancelled
    
End Function
