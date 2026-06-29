VERSION 5.00
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmGetEmpIDnFinYear 
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
   Icon            =   "frmGetEmpIDnFinYear.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   4740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   105
      TabIndex        =   0
      Tag             =   "MainUnit"
      Top             =   285
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.CommandButton cmdSpecial 
      Caption         =   "&Special"
      Height          =   375
      Left            =   780
      TabIndex        =   6
      Top             =   2055
      Width           =   1215
   End
   Begin VB.ComboBox cmbFinYear 
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   930
      Width           =   4515
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3390
      TabIndex        =   3
      Top             =   2055
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2085
      TabIndex        =   2
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
      TabIndex        =   5
      Top             =   750
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Employee ID"
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
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   60
      Width           =   1065
   End
End
Attribute VB_Name = "frmGetEmpIDnFinYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bCancelled As Boolean

Dim strEmpID As String
Dim strYear As String
Dim bSpecialButtonClicked As Boolean

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub cmdSpecial_Click()

    If cmbEmp.MatchFound = False Then
        MsgBox "Please Select Employee.", vbInformation
        Exit Sub
    End If
    If cmbFinYear.ListIndex = -1 Then
        MsgBox "Please Select Year.", vbInformation
        Exit Sub
    End If

    bSpecialButtonClicked = True
    
    bCancelled = False
    
    strEmpID = cmbEmp.ID
    strYear = cmbFinYear.Text
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    cmbEmp.ListHeight = 2200
    Call cmbEmp.AddVals(con, "'{' + EmpID + '} ' + Name", "Employees", "EmpID")
    
    cmbFinYear.Clear
    Dim iStartYear As Integer
    Dim rs As New ADODB.Recordset
    
    iStartYear = 2005
    Dim i As Integer
    For i = iStartYear To year(Date)
        cmbFinYear.AddItem i & " - " & i + 1 & " [ " & Format(DateSerial(i, 7, 1), "dd-MMM-yyyy") & " to " & Format(DateSerial(i + 1, 6, 30), "dd-MMM-yyyy") & " ]"
        cmbFinYear.ItemData(cmbFinYear.ListCount - 1) = i
    Next i
    bSpecialButtonClicked = False
    
End Sub

Private Sub OKButton_Click()

    If cmbEmp.MatchFound = False Then
        MsgBox "Please Select Employee.", vbInformation
        Exit Sub
    End If
    If cmbFinYear.ListIndex = -1 Then
        MsgBox "Please Select Year.", vbInformation
        Exit Sub
    End If
    
    bCancelled = False
    
    strEmpID = cmbEmp.ID
    strYear = cmbFinYear.Text
    
    Unload Me
End Sub

Public Function getValues(ByRef p_strEmpID As String, ByRef p_strYear As String, Optional ByRef p_bSpecialButtonClicked As Boolean = False, Optional strCaption As String = "") As Boolean
   
    If strCaption <> "" Then
        cmdSpecial.Caption = strCaption
        cmdSpecial.Visible = True
    End If
    
    Me.Show 1

    p_strEmpID = strEmpID
    p_strYear = strYear
    p_bSpecialButtonClicked = bSpecialButtonClicked
    
    getValues = Not bCancelled
    
End Function

Private Sub txtComments_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call OKButton_Click
    End If
End Sub
