VERSION 5.00
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmGetSingleComboValue_Ex 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List"
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
   Icon            =   "frmGetSingleComboValue_Ex.frx":0000
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
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   375
      Left            =   105
      TabIndex        =   6
      Top             =   3075
      Width           =   1215
   End
   Begin VB.TextBox txtRemarks 
      Height          =   2055
      Left            =   105
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   945
      Width           =   5880
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Remarks"
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
      TabIndex        =   4
      Top             =   735
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "List Name"
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
      TabIndex        =   3
      Top             =   60
      Width           =   825
   End
End
Attribute VB_Name = "frmGetSingleComboValue_Ex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim strReturn As String
Dim bFOI As Boolean, strTableName As String

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub cmbCombo_matched()
    txtRemarks = GetSingleStringValue("Remarks", strTableName, "WHERE EntryID=" & cmbCombo.ID)
End Sub

Private Sub cmdDelete_Click()

    If cmbCombo.MatchFound = False Then Exit Sub
    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    Dim strListUser As String
    strListUser = GetSingleStringValue("UserName", strTableName, " WHERE EntryID=" & cmbCombo.ID)
    
    If strListUser <> CurrentUserName Then
        MsgBox "Only the Author can delete.", vbInformation
        Exit Sub
    End If
    
    con.Execute "DELETE FROM " & strTableName & " WHERE EntryID=" & cmbCombo.ID
        
    cmbCombo.ClearVals
    cmbCombo.AddVals con, "ListName + ' / ' + UserName", strTableName, "EntryID", " WHERE Active=1"
    
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
    
    strReturn = cmbCombo.ID
    
    Unload Me
    
End Sub

Public Function getComboValue(ByRef strValue As String, Optional strLabelCaption As String = "Comments", Optional strFormCaption As String = "Comments", Optional p_bFOI As Boolean = True, Optional p_strTableName As String = "") As Boolean
    
    bFOI = p_bFOI
    
    If p_strTableName <> "" Then
        strTableName = p_strTableName
    Else
        If bFOI Then
            strTableName = "FOI_List"
        Else
            strTableName = "POI_List"
        End If
    End If
    
    cmbCombo.ClearVals
    cmbCombo.AddVals con, "ListName + ' / ' + UserName", strTableName, "EntryID", " WHERE Active=1"
    
    If strValue <> "" Then
        cmbCombo.Text = strValue
    End If
    
    Me.Caption = strFormCaption
    Me.Show 1

    strValue = strReturn
    getComboValue = Not bCancelled
    
End Function

Private Sub txtComments_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call OKButton_Click
    End If
End Sub

Public Sub ShowMe(Optional p_bFOI As Boolean = True)
        
End Sub
