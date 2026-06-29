VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGetNumericDataValues 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Groups"
   ClientHeight    =   7215
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGetNumericDataValues.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUnSelectAll 
      Caption         =   "Uncheck All"
      Height          =   375
      Left            =   1455
      TabIndex        =   4
      Top             =   6705
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Check All"
      Height          =   375
      Left            =   135
      TabIndex        =   3
      Top             =   6705
      Width           =   1215
   End
   Begin MSComctlLib.ListView LV 
      Height          =   5925
      Left            =   150
      TabIndex        =   2
      Top             =   690
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   10451
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Description"
         Object.Width           =   8819
      EndProperty
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4575
      TabIndex        =   1
      Top             =   6705
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3330
      TabIndex        =   0
      Top             =   6705
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   150
      TabIndex        =   5
      Top             =   315
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dddd, dd-MMM-yyyy"
      Format          =   115408899
      CurrentDate     =   39372
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Date"
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
      TabIndex        =   6
      Top             =   75
      Width           =   405
   End
End
Attribute VB_Name = "frmGetNumericDataValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim bChecked As Boolean
Dim lBalance As Long
Dim strReturn As String
Dim DTReturn As Date
Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub cmdSelectAll_Click()
    Dim i As Integer
    LV.ListItems(1).Checked = False
    For i = 2 To LV.ListItems.count
        LV.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdUnSelectAll_Click()
    Dim i As Integer
    
    For i = 1 To LV.ListItems.count
        LV.ListItems(1).Checked = False
    Next
End Sub

Private Sub Form_Load()

    bChecked = False
    DTPicker1.Value = Date
    
End Sub


Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)

    Dim i As Integer
    If Item.Tag = "0" Then
        For i = 2 To LV.ListItems.count
            LV.ListItems(i).Checked = False
        Next
    Else
        LV.ListItems(1).Checked = False
    End If
    
End Sub

Private Sub OKButton_Click()
    
    Dim i As Integer, bNothing As Boolean
    bNothing = True
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            bNothing = False
            Exit For
        End If
    Next
    
    If bNothing Then
        MsgBox "Select Something.", vbInformation
        LV.SetFocus
        Exit Sub
    End If
    
    bCancelled = False
    
    strReturn = ""
    If LV.ListItems(1).Checked Then
        strReturn = "All"
    Else
        For i = 2 To LV.ListItems.count
            With LV.ListItems(i)
                If .Checked Then
                    strReturn = strReturn & .Tag & ","
                End If
            End With
        Next
        If strReturn <> "" Then
            strReturn = Left(strReturn, Len(strReturn) - 1)
        End If
    End If
    DTReturn = DTPicker1.Value
    Unload Me
    
End Sub

Public Function getData(strTableName As String, strDescriptionField As String, strIDField As String, ByRef p_strReturn As String, ByRef p_DTReturn As Date) As Boolean

    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    
    With rs
        .Open "SELECT " & strIDField & "," & strDescriptionField & " FROM " & strTableName & " ORDER BY " & strDescriptionField, con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Set ITM = LV.ListItems.add(, , "< All >")
        ITM.Tag = "0"
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , .Fields(1).Value & "")
            ITM.Tag = Val(.Fields(0).Value & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    LV.ListItems(1).Checked = True
    
    Me.Show 1
    
    p_strReturn = strReturn
    p_DTReturn = DTReturn
    
    getData = Not bCancelled
    
End Function

