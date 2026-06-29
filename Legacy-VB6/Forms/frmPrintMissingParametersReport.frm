VERSION 5.00
Begin VB.Form frmPrintMissingParametersReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Missing Report"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5850
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
   ScaleHeight     =   2115
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2910
      TabIndex        =   3
      Top             =   1470
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4230
      TabIndex        =   2
      Top             =   1470
      Width           =   1215
   End
   Begin VB.OptionButton Opts 
      Caption         =   "Running Products"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   2820
      TabIndex        =   1
      Top             =   675
      Width           =   2370
   End
   Begin VB.OptionButton Opts 
      Caption         =   "Print All"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   675
      Value           =   -1  'True
      Width           =   2370
   End
End
Attribute VB_Name = "frmPrintMissingParametersReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ItemListMissingParaments.rpt")
    
    If Opts(0).Value Then
        rpt.ParameterFields.GetItemByName("@Running").SetCurrentValue False
    Else
        rpt.ParameterFields.GetItemByName("@Running").SetCurrentValue True
    End If
    
    Me.Hide
    
    Dim f As New frmPrevRpt
    f.ShowReport "", rpt, , False
    
    Unload Me
    
End Sub
