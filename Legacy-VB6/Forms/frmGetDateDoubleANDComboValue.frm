VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmGetDateDoubleANDComboValue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Date..."
   ClientHeight    =   3030
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGetDateDoubleANDComboValue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmb 
      Height          =   285
      Left            =   105
      TabIndex        =   0
      Top             =   420
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextBackColor   =   12640511
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.CheckBox chkPrintEarlier 
      Caption         =   "Print &Earlier"
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   135
      TabIndex        =   6
      Top             =   1515
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   135
      TabIndex        =   1
      Top             =   1140
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dddd, dd-MMM-yyyy"
      Format          =   153944067
      CurrentDate     =   39372
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4770
      TabIndex        =   4
      Top             =   2580
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3495
      TabIndex        =   3
      Top             =   2580
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   315
      Left            =   2595
      TabIndex        =   2
      Top             =   1140
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dddd, dd-MMM-yyyy"
      Format          =   153878531
      CurrentDate     =   39372
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "To"
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
      Left            =   2610
      TabIndex        =   8
      Top             =   900
      Width           =   210
   End
   Begin VB.Label lblCombo 
      AutoSize        =   -1  'True
      Caption         =   "Order No."
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
      Left            =   105
      TabIndex        =   7
      Top             =   165
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "From"
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
      TabIndex        =   5
      Top             =   900
      Width           =   435
   End
End
Attribute VB_Name = "frmGetDateDoubleANDComboValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim iReportOption As Integer
Dim bCancelled As Boolean
'Dim bEarlier As Boolean
Dim DT As Date, strComboValue As String, DTTo As Date

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub Form_Load()
    DTPicker1.Value = Date
    DTPicker2.Value = Date
End Sub

Private Sub OKButton_Click()
    bCancelled = False
'    If chkPrintEarlier.value = vbChecked Then
'        bEarlier = True
'    Else
'        bEarlier = False
'    End If
    If Cmb.MatchFound = False Then
        MsgBox "Please Select Value.", vbInformation
        Cmb.SetFocus
        Exit Sub
    End If
    
    strComboValue = Cmb.ID
    DT = DTPicker1.Value
    DTTo = DTPicker2.Value
    
    Unload Me
End Sub

Public Function getDate(ByRef myDT As Date, ByRef myDTTo As Date, ByRef p_strComboValue As String) As Boolean
    
    
    Me.Show 1
'    bPrintEarlier = bEarlier
    myDT = DT
    myDTTo = DTTo
    p_strComboValue = strComboValue
    
    getDate = Not bCancelled
    
End Function
