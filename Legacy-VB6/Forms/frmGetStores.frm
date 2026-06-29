VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGetStores 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Stores"
   ClientHeight    =   7785
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGetStores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUnSelectAll 
      Caption         =   "Uncheck All"
      Height          =   375
      Left            =   1455
      TabIndex        =   4
      Top             =   7230
      Width           =   1215
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "Check All"
      Height          =   375
      Left            =   135
      TabIndex        =   3
      Top             =   7230
      Width           =   1215
   End
   Begin MSComctlLib.ListView LV 
      Height          =   7065
      Left            =   135
      TabIndex        =   2
      Top             =   75
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   12462
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
         Text            =   "Store"
         Object.Width           =   7937
      EndProperty
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4575
      TabIndex        =   1
      Top             =   7230
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3330
      TabIndex        =   0
      Top             =   7230
      Width           =   1215
   End
End
Attribute VB_Name = "frmGetStores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim bChecked As Boolean
Dim lBalance As Long
Dim strStoreIDs As String


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
      
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT EntryID,StoreName FROM Stores ORDER BY StoreName", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Set ITM = LV.ListItems.add(, , "<All Stores>")
        ITM.Tag = "0"
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , !StoreName & "")
            ITM.Tag = Val(!EntryID & "")
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    LV.ListItems(1).Checked = True
    
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
        MsgBox "Select Customer.", vbInformation
        LV.SetFocus
        Exit Sub
    End If
    
    bCancelled = False
    
    strStoreIDs = ""
    If LV.ListItems(1).Checked Then
        strStoreIDs = "All"
    Else
        For i = 2 To LV.ListItems.count
            With LV.ListItems(i)
                If .Checked Then
                    strStoreIDs = strStoreIDs & Val(.Tag) & ","
                End If
            End With
        Next
        If strStoreIDs <> "" Then
            strStoreIDs = Left(strStoreIDs, Len(strStoreIDs) - 1)
        End If
    End If
    
    Unload Me
    
End Sub

Public Function getData(ByRef p_strStoreIDs As String) As Boolean
     
    Me.Show 1
    
    p_strStoreIDs = strStoreIDs
    
    getData = Not bCancelled
    
End Function
