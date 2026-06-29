VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStoreRacks 
   Caption         =   "Stores"
   ClientHeight    =   8595
   ClientLeft      =   225
   ClientTop       =   465
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkShowShelfs 
      Caption         =   "Show Bins"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   5685
      TabIndex        =   7
      Top             =   7965
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Height          =   480
      Left            =   60
      TabIndex        =   3
      Top             =   7980
      Width           =   5475
      Begin VB.TextBox txtFind 
         Height          =   315
         Left            =   1950
         TabIndex        =   6
         Top             =   135
         Width           =   2430
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Enabled         =   0   'False
         Height          =   300
         Left            =   4380
         TabIndex        =   5
         Top             =   135
         Width           =   1035
      End
      Begin VB.ComboBox cmbfield 
         Height          =   315
         ItemData        =   "frmStoreRacks.frx":0000
         Left            =   60
         List            =   "frmStoreRacks.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   135
         Width           =   1875
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1695
      Top             =   3330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreRacks.frx":0004
            Key             =   "ITM"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreRacks.frx":0458
            Key             =   "ORD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreRacks.frx":08B0
            Key             =   "PARIN"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreRacks.frx":0D04
            Key             =   "ITMIN"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreRacks.frx":1158
            Key             =   "CAT"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreRacks.frx":15AC
            Key             =   "OK"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStoreRacks.frx":1A04
            Key             =   "TYPE"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LV 
      Height          =   7845
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   11730
      _ExtentX        =   20690
      _ExtentY        =   13838
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   15199215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Store"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Racks"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Shelfs"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSForms.CommandButton cmdOpt 
      Height          =   390
      Left            =   8700
      TabIndex        =   1
      Top             =   8025
      Width           =   1545
      Caption         =   "Options        "
      PicturePosition =   327683
      Size            =   "2725;688"
      Accelerator     =   79
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   390
      Left            =   10275
      TabIndex        =   2
      Top             =   8025
      Width           =   1545
      ForeColor       =   64
      Caption         =   "Close          "
      PicturePosition =   327683
      Size            =   "2725;688"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Menu mnupop 
      Caption         =   "mnupop"
      Begin VB.Menu mnurefresh 
         Caption         =   "Re&fresh"
      End
      Begin VB.Menu ln1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNewItem 
         Caption         =   "&Add"
      End
      Begin VB.Menu mnuAddSub 
         Caption         =   "&Add Sub"
      End
      Begin VB.Menu mnuEditItem 
         Caption         =   "&Edit Item"
      End
      Begin VB.Menu mnuDash2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddMore 
         Caption         =   "Add More"
      End
   End
End
Attribute VB_Name = "frmStoreRacks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub ShowMe()
    Call FillStores
    Me.Show
End Sub

Private Sub FillStores()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT * FROM Stores", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !EntryID & "'", !StoreName & "")
            ITM.Bold = True
            ITM.ForeColor = vbRed
            ITM.Tag = "0" 'Means It's Store
            ITM.ListSubItems.add().ForeColor = vbRed
            ITM.ListSubItems.add().ForeColor = vbRed
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Call FillRacks
    Call FillShelfs
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub FillRacks()
    
    If LV.ListItems.count = 0 Then Exit Sub
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim iIndex As Integer
    Dim strKeys() As String
    ReDim strKeys(LV.ListItems.count)
    For i = 1 To LV.ListItems.count
        strKeys(i - 1) = LV.ListItems(i).key
    Next
    
    For i = 0 To UBound(strKeys) - 1
        iIndex = LV.ListItems(strKeys(i)).Index
        With rs
            .Open "SELECT * FROM StoreRacks WHERE Store_RefID=" & Val(strKeys(i)), con, adOpenForwardOnly, adLockReadOnly
            LV.ListItems(strKeys(i)).SubItems(1) = .RecordCount
            Do Until .EOF
                Set ITM = LV.ListItems.add(iIndex + 1, !EntryID & "''")
                ITM.Tag = 1 'Means It's Rack
                ITM.ListSubItems.add(, , !RackNo & "").ForeColor = vbBlue
                iIndex = iIndex + 1
                .MoveNEXT
            Loop
            .Close
        End With
    Next
    Set rs = Nothing
    
End Sub

Private Sub FillShelfs()
    
    If LV.ListItems.count = 0 Then Exit Sub
    Dim i As Integer
    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    Dim iIndex As Integer
    Dim strKeys() As String
    Dim iRackCount As Integer
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Tag = "1" Then
            iRackCount = iRackCount + 1
        End If
    Next
    If iRackCount = 0 Then Exit Sub
    
    ReDim strKeys(iRackCount)
    
    iIndex = 0
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Tag = "1" Then   'Get Racks only
            strKeys(iIndex) = LV.ListItems(i).key
            iIndex = iIndex + 1
        End If
    Next
    
    For i = 0 To UBound(strKeys) - 1
        iIndex = LV.ListItems(strKeys(i)).Index
        With rs
            .Open "SELECT * FROM StoreShelfs WHERE Rack_RefID=" & Val(strKeys(i)), con, adOpenForwardOnly, adLockReadOnly
            LV.ListItems(strKeys(i)).SubItems(2) = .RecordCount
            If chkShowShelfs.Value = vbChecked Then
                Do Until .EOF
                    Set ITM = LV.ListItems.add(iIndex + 1, !EntryID & "'''")
                    ITM.Tag = 2 'Shelf
                    ITM.SubItems(1) = ""
                    ITM.SubItems(2) = !ShelfNo & ""
                    iIndex = iIndex + 1
                    .MoveNEXT
                Loop
            End If
            .Close
        End With
    Next
    Set rs = Nothing
    
End Sub

Private Sub chkShowShelfs_Click()
    Call FillStores
End Sub

Private Sub ShowHideShelfs(bShow As Boolean)
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Tag = "2" Then
            LV.ListItems(i).Ghosted = Not bShow
        End If
    Next
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    mnuPop.Visible = False
End Sub

Private Sub mnunewcat_Click()

End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 2 Then Exit Sub
    
    On Error Resume Next
    Dim c As Control
    For Each c In Controls
        If TypeOf c Is Menu Then
            If c.Name <> "mnupop" Then
                c.Visible = True
            End If
        End If
    Next
        
    If LV.ListItems.count = 0 Then
        mnuNewItem.Caption = "Add New Store"
        mnuAddSub.Visible = False
        mnuEditItem.Visible = False
    Else
        If LV.SelectedItem.Tag = "0" Then
            mnuNewItem.Caption = "Add New Store"
            mnuAddSub.Caption = "Add New Rack"
            mnuEditItem.Caption = "Edit Store"
        ElseIf LV.SelectedItem.Tag = "1" Then
            mnuNewItem.Caption = "Add New Rack"
            mnuAddSub.Caption = "Add New Bin"
            mnuEditItem.Caption = "Edit Rack"
        ElseIf LV.SelectedItem.Tag = "2" Then
            mnuNewItem.Caption = "Add New Rack"
            mnuAddSub.Caption = "Add New Bin"
            mnuEditItem.Caption = "Edit Bin"
        End If
    End If
    Me.PopUpMenu mnuPop
    
End Sub

Private Sub mnuAddMore_Click()
    frmAddStoreRackBin.Show 1
End Sub

Private Sub mnuAddSub_Click()
    Call ExecuteCommand(mnuAddSub.Caption)
End Sub

Private Sub MnuEditItem_Click()
    Call ExecuteCommand(mnuEditItem.Caption)
End Sub

Private Sub mnuNewItem_Click()
    Call ExecuteCommand(mnuNewItem.Caption)
End Sub

Private Sub mnuRefresh_Click()
    Call FillStores
End Sub

Private Sub ExecuteCommand(strCommand As String)

    Dim f As Form
    If strCommand = "Add New Store" Then
        Set f = New frmNewStore
        Load f
        If f.ShowMe() Then
            Call FillStores
        End If
    ElseIf strCommand = "Edit Store" Then
        Set f = New frmNewStore
        Load f
        If f.ShowMe(Val(LV.SelectedItem.key)) Then
            Call FillStores
        End If
    ElseIf strCommand = "Add New Rack" Then
        Set f = New frmNewRack
        Load f
        If f.ShowMe Then
            Call FillStores
        End If
    ElseIf strCommand = "Edit Rack" Then
        Set f = New frmNewRack
        Load f
        If f.ShowMe(Val(LV.SelectedItem.key)) Then
            Call FillStores
        End If
    ElseIf strCommand = "Add New Bin" Then
        Set f = New frmNewShelf
        Load f
        If f.ShowMe Then
            Call FillStores
        End If
    ElseIf strCommand = "Edit Bin" Then
        Set f = New frmNewShelf
        Load f
        If f.ShowMe(Val(LV.SelectedItem.key)) Then
            Call FillStores
        End If
    End If
    
End Sub
