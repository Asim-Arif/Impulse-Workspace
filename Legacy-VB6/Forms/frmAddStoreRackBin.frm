VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmAddStoreRackBin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stores / Racks / Bins"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8325
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
   ScaleHeight     =   8490
   ScaleWidth      =   8325
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   195
      TabIndex        =   11
      Top             =   7875
      Width           =   2160
   End
   Begin ComboList.Usercontrol1 cmbBin 
      Height          =   285
      Left            =   5475
      TabIndex        =   5
      Top             =   555
      Width           =   2610
      _ExtentX        =   4604
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
      TextBackColor   =   12640511
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
      Appearance      =   0
   End
   Begin ComboList.Usercontrol1 cmbRack 
      Height          =   285
      Left            =   2835
      TabIndex        =   3
      Top             =   555
      Width           =   2610
      _ExtentX        =   4604
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
      TextBackColor   =   12640511
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
      Appearance      =   0
   End
   Begin ComboList.Usercontrol1 cmbStore 
      Height          =   285
      Left            =   195
      TabIndex        =   1
      Top             =   555
      Width           =   2610
      _ExtentX        =   4604
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
      TextBackColor   =   12640511
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
      Appearance      =   0
   End
   Begin VB.CommandButton cmdAddRack 
      Caption         =   "Add"
      Height          =   330
      Left            =   4410
      TabIndex        =   10
      Top             =   195
      Width           =   1035
   End
   Begin VB.CommandButton cmdAddStore 
      Caption         =   "Add"
      Height          =   330
      Left            =   1755
      TabIndex        =   9
      Top             =   195
      Width           =   1035
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6840
      TabIndex        =   8
      Top             =   7875
      Width           =   1245
   End
   Begin MSComctlLib.ListView LV 
      Height          =   6960
      Left            =   195
      TabIndex        =   7
      Top             =   855
      Width           =   7890
      _ExtentX        =   13917
      _ExtentY        =   12277
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S #"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bin #"
         Object.Width           =   5292
      EndProperty
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   330
      Left            =   7050
      TabIndex        =   6
      Top             =   195
      Width           =   1035
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bin"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   5505
      TabIndex        =   4
      Top             =   315
      Width           =   255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Rack"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   2880
      TabIndex        =   2
      Top             =   315
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Store"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   0
      Left            =   195
      TabIndex        =   0
      Top             =   315
      Width           =   465
   End
End
Attribute VB_Name = "frmAddStoreRackBin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbRack_matched()
    Call RefreshBins
    Call RefreshLVBins
End Sub

Private Sub cmbStore_matched()
    Call RefreshRacks
    Call RefreshLVRacks
End Sub

Private Sub cmdAdd_Click()

    'Check for Existing Rack..
    If cmbStore.MatchFound = False Then
        MsgBox "Please Select Store.", vbInformation
        Exit Sub
    ElseIf cmbRack.MatchFound = False Then
        MsgBox "Please Select Rack.", vbInformation
        Exit Sub
    End If
    
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(*)", "StoreShelfs", " WHERE ShelfNo='" & cmbBin.Text & "' AND Rack_RefID=" & cmbRack.ID)
    If lCount > 0 Then
        MsgBox "Already Added.", vbInformation
        Exit Sub
    End If
    
    Dim lEntryID As Long
    con.Execute "INSERT INTO StoreShelfs(ShelfNo,Rack_RefID) VALUES('" & cmbBin.Text & "'," & cmbRack.ID & ")"
    
    lEntryID = GetSingleLongValue("MAX(EntryID)", "StoreShelfs")
    Call RefreshBins(lEntryID)
    Call RefreshLVBins
    
End Sub

Private Sub cmdAddRack_Click()

    'Check for Existing Rack..
    If cmbStore.MatchFound = False Then
        MsgBox "Please Select Store.", vbInformation
        Exit Sub
    End If
    
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(*)", "StoreRacks", " WHERE RackNo='" & cmbRack.Text & "' AND Store_RefID=" & cmbStore.ID)
    If lCount > 0 Then
        MsgBox "Already Added.", vbInformation
        Exit Sub
    End If
    
    Dim lEntryID As Long
    con.Execute "INSERT INTO StoreRacks(RackNo,Store_RefID) VALUES('" & cmbRack.Text & "'," & cmbStore.ID & ")"
    
    lEntryID = GetSingleLongValue("MAX(EntryID)", "StoreRacks")
    Call RefreshRacks(lEntryID)
    
End Sub

Private Sub cmdAddStore_Click()

    'Check for Existing..
    Dim lCount As Long
    lCount = GetSingleLongValue("COUNT(*)", "Stores", " WHERE StoreName='" & cmbStore.Text & "'")
    If lCount > 0 Then
        MsgBox "Already Added.", vbInformation
        Exit Sub
    End If
    
    Dim lEntryID As Long
    con.Execute "INSERT INTO Stores(StoreName) VALUES('" & cmbStore.Text & "')"
    
    lEntryID = GetSingleLongValue("MAX(EntryID)", "Stores")
    Call RefreshStores(lEntryID)
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\Stores_Report.rpt")
    
    Dim f As New frmPrevRpt
    Me.Hide
    f.ShowReport "", rpt, , , , Me
    Unload Me
    
End Sub

Private Sub Form_Load()

    cmbStore.ListHeight = 4000
    
    Call RefreshStores
    Call RefreshLVStores
    
End Sub

Private Sub RefreshStores(Optional lStoreID As Long)

    'Dim iStoreID As Integer
    'iStoreID = cmbStore.ID
    cmbStore.ClearVals
    cmbStore.AddVals con_ServerSide, "StoreName", "Stores", "EntryID"
    
    If lStoreID > 0 Then
        cmbStore.ID = lStoreID
    End If
    
End Sub

Private Sub RefreshRacks(Optional lRackID As Long = 0)

    cmbRack.ClearVals
    cmbRack.AddVals con_ServerSide, "RackNo", "StoreRacks", "EntryID", " WHERE Store_RefID=" & cmbStore.ID
    
    If lRackID > 0 Then
        cmbRack.ID = lRackID
    End If
    
End Sub

Private Sub RefreshBins(Optional lEntryID As Long = 0)

    cmbBin.ClearVals
    cmbBin.AddVals con_ServerSide, "ShelfNo", "StoreShelfs", "EntryID", " WHERE Rack_RefID=" & cmbRack.ID
    
    If lEntryID > 0 Then
        cmbBin.ID = lEntryID
    End If
    
End Sub


Private Sub RefreshLVStores()
    
    LV.ColumnHeaders(2).Text = "Store"
    LV.Tag = "Stores"
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT * FROM Stores ORDER BY StoreName", con_ServerSide, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.ListSubItems.add , , !StoreName & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub RefreshLVRacks()
    
    LV.ColumnHeaders(2).Text = "Rack No."
    LV.Tag = "Racks"
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT * FROM StoreRacks WHERE Store_RefID=" & cmbStore.ID & " ORDER BY RackNo", con_ServerSide, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.ListSubItems.add , , !RackNo & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub


Private Sub RefreshLVBins()
    
    LV.ColumnHeaders(2).Text = "Bin No."
    LV.Tag = "Bins"
    Dim rs As New ADODB.Recordset, ITM As ListItem
    With rs
        .Open "SELECT * FROM StoreShelfs WHERE Rack_RefID=" & cmbRack.ID & " ORDER BY ShelfNo", con_ServerSide, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.ListSubItems.add , , !ShelfNo & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

