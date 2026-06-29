VERSION 5.00
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmAddReference 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Reference"
   ClientHeight    =   2835
   ClientLeft      =   7170
   ClientTop       =   5820
   ClientWidth     =   4965
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
   ScaleHeight     =   2835
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2730
      Left            =   45
      TabIndex        =   7
      Top             =   0
      Width           =   4845
      Begin ComboList.Usercontrol1 cmbGuarantors 
         Height          =   285
         Left            =   90
         TabIndex        =   12
         Top             =   360
         Width           =   4665
         _ExtentX        =   8229
         _ExtentY        =   503
         BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TextBackColor   =   12640511
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ListIndex       =   -1
         Appearance      =   0
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   3
         Left            =   3150
         TabIndex        =   3
         Top             =   1605
         Width           =   1590
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   75
         TabIndex        =   0
         Top             =   1020
         Width           =   2325
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   2430
         TabIndex        =   1
         Top             =   1020
         Width           =   2325
      End
      Begin VB.CommandButton cmdSavenNew 
         Caption         =   "Save && &New"
         Height          =   405
         Left            =   1350
         TabIndex        =   6
         Top             =   2190
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.TextBox txts 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   75
         TabIndex        =   2
         Top             =   1605
         Width           =   3045
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Save && Close"
         Height          =   405
         Left            =   2490
         TabIndex        =   4
         Top             =   2190
         Width           =   1110
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   3645
         TabIndex        =   5
         Top             =   2190
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C/O"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   13
         Top             =   165
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   2
         Left            =   3165
         TabIndex        =   11
         Top             =   1410
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Relation"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   2460
         TabIndex        =   10
         Top             =   825
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   9
         Top             =   825
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Organization"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   8
         Top             =   1395
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmAddReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCancelled As Boolean
Dim strInstitute As String, strRelation As String, strName As String, strContactNo As String, iGuarantorID As Integer

Public Function ShowMe(ByRef p_strInstitute As String, ByRef p_strRelation As String, ByRef p_strName As String, ByRef p_strContactNo As String, ByRef p_iGuarantorID As Integer) As Integer

    Me.Show 1
    
    p_strInstitute = strInstitute
    
    p_strRelation = strRelation
    p_strName = strName
    p_strContactNo = strContactNo
    p_iGuarantorID = iGuarantorID
    ShowMe = Not bCancelled
    
End Function


Private Sub cmdCancel_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If Not ValidateData Then
        Exit Sub
    End If
    bCancelled = False
    strInstitute = txts(0)
    strRelation = txts(1)
    strName = txts(2)
    strContactNo = txts(3)
    
    iGuarantorID = IIf(cmbGuarantors.MatchFound, cmbGuarantors.ID, 0)
    
    Unload Me
    
End Sub

Private Function ValidateData() As Boolean

    If txts(0) = "" Then
        MsgBox "Invalid Organization.", vbInformation
        Exit Function
    ElseIf txts(1) = "" Then
        MsgBox "Invalid Relation.", vbInformation
        Exit Function
    ElseIf txts(2) = "" Then
        MsgBox "Invalid Name.", vbInformation
        Exit Function
    ElseIf txts(3) = "" Then
        MsgBox "Invalid Contact No.", vbInformation
        Exit Function
'    ElseIf cmbGuarantors.MatchFound = False Then
'        MsgBox "Please Select C/O.", vbInformation
'        Exit Function
    End If
    ValidateData = True
    
End Function


Private Sub Form_Load()
    cmbGuarantors.ListHeight = 1600
    cmbGuarantors.AddVals con, "REPLACE(strFullName,'#','No.') + ' ' + REPLACE(strFatherName,'#','No.') + ' {' + CAST(intGuarantorID AS VARCHAR(5))+ '}'", "GuarantorList_Imported", "intGuarantorID"
End Sub
