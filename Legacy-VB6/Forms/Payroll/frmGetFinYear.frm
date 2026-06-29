VERSION 5.00
Begin VB.Form frmGetFinYear 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Employee"
   ClientHeight    =   2520
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4740
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGetFinYear.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbFinYear 
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Top             =   270
      Width           =   4515
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3390
      TabIndex        =   2
      Top             =   2055
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2085
      TabIndex        =   1
      Top             =   2055
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Financial Year"
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
      Left            =   120
      TabIndex        =   3
      Top             =   90
      Width           =   1170
   End
End
Attribute VB_Name = "frmGetFinYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCancelled As Boolean

Dim strYear As String

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub Form_Load()

    cmbFinYear.Clear
    Dim iStartYear As Integer
    Dim rs As New ADODB.Recordset
    
    iStartYear = 2005
    Dim i As Integer
    For i = iStartYear To year(Date)
        cmbFinYear.AddItem i & " - " & i + 1 & " [ " & Format(DateSerial(i, 7, 1), "dd-MMM-yyyy") & " to " & Format(DateSerial(i + 1, 6, 30), "dd-MMM-yyyy") & " ]"
        cmbFinYear.ItemData(cmbFinYear.ListCount - 1) = i
    Next i
    
End Sub

Private Sub OKButton_Click()

    If cmbFinYear.ListIndex = -1 Then
        MsgBox "Please Select Year.", vbInformation
        Exit Sub
    End If
    
    bCancelled = False

    strYear = cmbFinYear.Text
    
    Unload Me
End Sub

Public Function getValues(ByRef p_strYear As String) As Boolean
   
  
    Me.Show 1

    p_strYear = strYear
    
    getValues = Not bCancelled
    
End Function
