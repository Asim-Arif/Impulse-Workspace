VERSION 5.00
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmNewShelf 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bin Defination"
   ClientHeight    =   2865
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
   Icon            =   "frmNewShelf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCont 
      Caption         =   "Make Bins by Continous Numbers"
      Height          =   210
      Left            =   105
      TabIndex        =   10
      Top             =   2265
      Width           =   2880
   End
   Begin VB.TextBox txtTotalBins 
      Height          =   315
      Left            =   3375
      TabIndex        =   9
      Top             =   1905
      Width           =   1245
   End
   Begin ComboList.Usercontrol1 cmbStore 
      Height          =   285
      Left            =   105
      TabIndex        =   1
      Top             =   375
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
   Begin VB.TextBox txtShelf 
      Height          =   315
      Left            =   105
      TabIndex        =   5
      Top             =   1905
      Width           =   3255
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   2115
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   1635
      Width           =   1215
   End
   Begin ComboList.Usercontrol1 cmbRack 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   945
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total Bins"
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
      Index           =   3
      Left            =   3375
      TabIndex        =   8
      Top             =   1695
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Rack"
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
      Index           =   2
      Left            =   135
      TabIndex        =   2
      Top             =   705
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Store Name"
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
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Bin Name"
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
      Left            =   135
      TabIndex        =   4
      Top             =   1695
      Width           =   780
   End
End
Attribute VB_Name = "frmNewShelf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim lEntryID As Long

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub cmbStore_matched()
    cmbRack.ClearVals
    cmbRack.AddVals con, "RackNo", "StoreRacks", "EntryID", " WHERE Store_RefID=" & cmbStore.ID
End Sub

Private Sub Form_Load()
    cmbStore.ListHeight = 2400
    cmbRack.ListHeight = 2400
    cmbStore.AddVals con, "StoreName", "Stores", "EntryID", " ORDER BY StoreName"
End Sub

Private Sub OKButton_Click()
     
    On Error GoTo err
    If cmbStore.MatchFound = False Then
        MsgBox "Please Select Store.", vbInformation
        Exit Sub
    ElseIf cmbRack.MatchFound = False Then
        MsgBox "Please Select Rack.", vbInformation
        Exit Sub
    ElseIf txtShelf = "" Then
        MsgBox "Invalid Shelf.", vbInformation
        Exit Sub
    End If
    
    Dim iBinNo As Integer
    'iBinNo = GetSingleLongValue("COUNT(ShelfNo)", "StoreShelfs", " WHERE Rack_RefID=" & cmbRack.ID)
    iBinNo = GetSingleLongValue("COUNT(ShelfNo)", "VStoreShelfs", " WHERE Store_RefID=" & cmbStore.ID)
    If lEntryID = 0 Then
        Dim i As Integer
        For i = 1 To Val(txtTotalBins)
            iBinNo = iBinNo + 1
            con.Execute "INSERT INTO StoreShelfs(Rack_RefID,ShelfNo) VALUES(" & cmbRack.ID & ",'" & txtShelf & "-" & IIf(chkCont.Value = vbChecked, iBinNo, i) & "')"
        Next
    Else
        con.Execute "UPDATE StoreShelfs SET ShelfNo='" & txtShelf & "' WHERE EntryID=" & lEntryID
    End If
    bCancelled = False
    Unload Me
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Public Function ShowMe(Optional p_lEntryID As Long = 0) As Boolean
    
    lEntryID = p_lEntryID
    
    If lEntryID <> 0 Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT * FROM VStoreShelfs WHERE EntryID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
            cmbStore.ID = !Store_RefID
            cmbRack.ID = !Rack_RefID
            txtShelf = !ShelfNo & ""
            .Close
        End With
        Set rs = Nothing
    End If
    
    Me.Show 1

    ShowMe = Not bCancelled
    
End Function
