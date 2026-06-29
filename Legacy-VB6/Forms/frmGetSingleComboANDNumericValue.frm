VERSION 5.00
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmGetSingleComboANDNumericValue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Comments...."
   ClientHeight    =   3495
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
   Icon            =   "frmGetSingleComboANDNumericValue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbCombo 
      Height          =   285
      Left            =   105
      TabIndex        =   2
      Top             =   300
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextBackColor   =   12640511
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.TextBox txtNumeric 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   4965
      TabIndex        =   4
      Top             =   285
      Width           =   1020
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4755
      TabIndex        =   1
      Top             =   3075
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3465
      TabIndex        =   0
      Top             =   3075
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Lable 2"
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
      Left            =   4950
      TabIndex        =   5
      Top             =   75
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Comments"
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
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   915
   End
End
Attribute VB_Name = "frmGetSingleComboANDNumericValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim strReturn As String, dReturn As Double

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub Form_Load()
    strReturn = ""
    cmbCombo.ListHeight = 2400
End Sub

Private Sub OKButton_Click()
    
    If cmbCombo.MatchFound = False Then
        MsgBox "Invalid Selection.", vbInformation
        Exit Sub
    End If
    bCancelled = False
    
    strReturn = cmbCombo.id
    dReturn = Val(txtNumeric)
    
    Unload Me
    
End Sub

Public Function getComboValue(ByRef strValue As String, ByRef dValue As Double, Optional strLabelCaption As String = "Comments", Optional strLabel2Caption As String = "", Optional strFormCaption As String = "Comments") As Boolean

    If strValue <> "" Then
        cmbCombo.Text = strValue
    End If
    
    Label1.Caption = strLabelCaption
    Label2.Caption = strLabel2Caption
    
    Me.Caption = strFormCaption
    Me.Show 1

    strValue = strReturn
    dValue = dReturn
    
    getComboValue = Not bCancelled
    
End Function

Private Sub txtComments_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call OKButton_Click
    End If
End Sub
