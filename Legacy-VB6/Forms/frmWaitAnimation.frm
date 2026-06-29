VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmWaitAnimation 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7665
   ClientLeft      =   2310
   ClientTop       =   1515
   ClientWidth     =   7920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5290.519
   ScaleMode       =   0  'User
   ScaleWidth      =   7437.291
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.Animation Animation1 
      Height          =   5265
      Left            =   210
      TabIndex        =   0
      Top             =   2100
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   9287
      _Version        =   393216
      AutoPlay        =   -1  'True
      Center          =   -1  'True
      FullWidth       =   498
      FullHeight      =   351
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Please wait While Loading Data"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1680
      Left            =   1380
      TabIndex        =   1
      Top             =   180
      Width           =   5310
   End
End
Attribute VB_Name = "frmWaitAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    Animation1.Open App.Path & "\bonhms.avi"
    
    'Animation1.Open App.Path & "\loading_round.avi"
End Sub
