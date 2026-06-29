VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Date"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   2880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1530
      TabIndex        =   1
      Top             =   2610
      Width           =   1245
   End
   Begin MSComCtl2.MonthView DT 
      Height          =   2370
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   64946177
      CurrentDate     =   38918
   End
End
Attribute VB_Name = "frmDT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myDT As Date

Public Function GetDate(p_myDT As Date) As Date
    
    DT.value = p_myDT
    Me.Show 1
'    myDT = DT.value
    GetDate = myDT
    
End Function

Private Sub cmdOK_Click()
    myDT = DT.value
    Unload Me
End Sub
