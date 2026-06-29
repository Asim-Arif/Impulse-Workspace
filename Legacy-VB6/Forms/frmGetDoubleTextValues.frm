VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmGetDoubleTextValues 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Comments...."
   ClientHeight    =   2880
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4785
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGetDoubleTextValues.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRemarks 
      Height          =   1245
      Left            =   105
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1050
      Width           =   4515
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3405
      TabIndex        =   3
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2085
      TabIndex        =   2
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtComments 
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   405
      Width           =   4515
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
      Left            =   105
      TabIndex        =   7
      Top             =   840
      Width           =   915
   End
   Begin MSForms.ComboBox ComboBox1 
      Height          =   300
      Left            =   105
      TabIndex        =   6
      Top             =   1050
      Visible         =   0   'False
      Width           =   4515
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "7964;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbComments 
      Height          =   300
      Left            =   105
      TabIndex        =   4
      Top             =   405
      Visible         =   0   'False
      Width           =   4515
      VariousPropertyBits=   746604571
      DisplayStyle    =   7
      Size            =   "7964;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
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
Attribute VB_Name = "frmGetDoubleTextValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCancelled As Boolean

Dim strReturn As String, strRemarks As String
Dim bShowCombo As String

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub Form_Load()
    strReturn = ""
End Sub

Private Sub OKButton_Click()
    bCancelled = False
    If bShowCombo Then
        strReturn = cmbComments
    Else
        strReturn = txtComments
    End If
    strRemarks = txtRemarks
    Unload Me
End Sub

Public Function getTextValue(ByRef strValue As String, ByRef p_strRemarks As String, Optional p_bShowCombo As Boolean = False, Optional strLabelCaption As String = "Comments", Optional strFormCaption As String = "Comments", Optional strLabel2Caption As String) As Boolean
    
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
    Me.Caption = strFormCaption
    Me.Show 1

    strValue = strReturn
    p_strRemarks = strRemarks
    getTextValue = Not bCancelled
    
End Function

Private Sub txtComments_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call OKButton_Click
    End If
End Sub
