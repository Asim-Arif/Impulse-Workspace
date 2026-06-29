VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmGetTextAndNumericValue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Comments...."
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
   Icon            =   "frmGetTextAndNumericValue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumeric 
      Height          =   315
      Left            =   3570
      TabIndex        =   6
      Top             =   405
      Width           =   1020
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   630
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   150
      Width           =   1215
   End
   Begin VB.TextBox txtComments 
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   405
      Width           =   3450
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
      Left            =   3585
      TabIndex        =   5
      Top             =   195
      Width           =   600
   End
   Begin MSForms.ComboBox cmbComments 
      Height          =   300
      Left            =   105
      TabIndex        =   4
      Top             =   405
      Visible         =   0   'False
      Width           =   3450
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "6085;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
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
      Left            =   150
      TabIndex        =   0
      Top             =   195
      Width           =   915
   End
End
Attribute VB_Name = "frmGetTextAndNumericValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCancelled As Boolean

Dim strReturn As String, dReturn As Double
Dim bShowCombo As String

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub Form_Load()
    strReturn = ""
End Sub

Private Sub OKButton_Click()
    
    If bShowCombo Then
        strReturn = cmbComments
    Else
        strReturn = txtComments
    End If
    
    If strReturn = "" Or Val(txtNumeric) = 0 Then
        MsgBox "Invalid Data.", vbInformation
        Exit Sub
    End If
    dReturn = Val(txtNumeric)
    bCancelled = False
    Unload Me
    
End Sub

Public Function getTextANDNumericValue(ByRef strValue As String, ByRef dValue As Double, Optional p_bShowCombo As Boolean = False, Optional strLabelCaption As String = "Comments", Optional strLabel2Caption As String = "Comments", Optional strFormCaption As String = "Comments") As Boolean
    
    bShowCombo = p_bShowCombo
    If bShowCombo Then
        txtComments.Visible = False
        cmbComments.Visible = True
        If strValue <> "" Then
            Dim i As Integer
            For i = 0 To cmbComments.ListCount - 1
                If cmbComments.List(i) = strValue Then
                    cmbComments.ListIndex = i
                    Exit For
                End If
            Next
        End If
    Else
        txtComments.Visible = True
        cmbComments.Visible = False
    End If
    Label1.Caption = strLabelCaption
    Label2.Caption = strLabel2Caption
    
    Me.Caption = strFormCaption
    Me.Show 1

    strValue = strReturn
    dValue = dReturn
    
    getTextANDNumericValue = Not bCancelled
    
End Function

Private Sub txtNumeric_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call OKButton_Click
    End If
End Sub
