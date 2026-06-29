VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmGetLocation 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6615
   ClientLeft      =   1095
   ClientTop       =   435
   ClientWidth     =   11565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   11565
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9675
      TabIndex        =   14
      Top             =   5895
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8085
      TabIndex        =   13
      Top             =   5895
      Width           =   1500
   End
   Begin VB.TextBox txtItem 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   405
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   870
      Width           =   10740
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10380
      TabIndex        =   10
      Top             =   1215
      Width           =   780
   End
   Begin VB.Frame Frame2 
      Caption         =   "Move To :"
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
      Height          =   1260
      Left            =   405
      TabIndex        =   4
      Top             =   1515
      Width           =   10770
      Begin VB.TextBox txtInHandTo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   9855
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   585
         Width           =   810
      End
      Begin MSForms.ComboBox cmbStore 
         Height          =   315
         Left            =   90
         TabIndex        =   9
         Top             =   585
         Width           =   3705
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "6535;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbRack 
         Height          =   315
         Left            =   3810
         TabIndex        =   8
         Top             =   585
         Width           =   2565
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "4524;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.ComboBox cmbShelf 
         Height          =   315
         Left            =   6390
         TabIndex        =   7
         Top             =   585
         Width           =   3450
         VariousPropertyBits=   746604571
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "6085;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label Label4 
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   6
         Top             =   345
         Width           =   10575
         BackColor       =   11517387
         Caption         =   $"frmGetLocation.frx":0000
         Size            =   "18653;423"
         BorderColor     =   16512
         BorderStyle     =   1
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.Frame fTitle 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Get Location"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   390
         Index           =   0
         Left            =   15
         TabIndex        =   1
         Top             =   135
         Width           =   10335
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Get Location"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   390
         Index           =   1
         Left            =   30
         TabIndex        =   2
         Top             =   150
         Width           =   10335
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Qty :"
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
      Left            =   9945
      TabIndex        =   11
      Top             =   1230
      Width           =   390
   End
   Begin MSForms.Label Label2 
      Height          =   240
      Left            =   405
      TabIndex        =   3
      Top             =   645
      Width           =   10740
      BackColor       =   11517387
      Caption         =   $"frmGetLocation.frx":00D6
      Size            =   "18944;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmGetLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim lShelf_RefID As Long
Dim strItemID As String

Private Sub cmbRack_Change()
    Call cmbRack_Click
End Sub

Private Sub cmbRack_Click()

    On Error GoTo err
    
    If cmbRack.MatchFound = False Then Exit Sub
    Call AddToCombo(cmbShelf, "ShelfNo", "StoreShelfs", " WHERE Rack_RefID=" & cmbRack.List(cmbRack.ListIndex, 1), , "EntryID")
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbShelf_Change()
    Call cmbShelf_Click
End Sub

Private Sub cmbShelf_Click()

    Dim dQty As Double
    If cmbShelf.MatchFound = False Then Exit Sub
    dQty = GetSingleDoubleValue("NetQty", "VItemsWithShelfWiseStock", " WHERE Shelf_RefID=" & cmbShelf.List(cmbShelf.ListIndex, 1) & " AND ItemID='" & strItemID & "'")
    txtInHandTo = Round(dQty, 2)
    
End Sub

Private Sub cmbStore_Change()
    Call cmbStore_Click
End Sub

Private Sub cmbStore_Click()

    On Error GoTo err
    
    If cmbStore.MatchFound = False Then Exit Sub
    Call AddToCombo(cmbRack, "RackNo", "StoreRacks", " WHERE Store_RefID=" & cmbStore.List(cmbStore.ListIndex, 1), , "EntryID")
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdCancel_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()

    If cmbShelf.MatchFound = False Then
        MsgBox "Please Select Location.", vbInformation
        Exit Sub
    End If
        
    lShelf_RefID = cmbShelf.List(cmbShelf.ListIndex, 1)
    bCancelled = False
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    bCancelled = True
    Call AddToCombo(cmbStore, "StoreName", "Stores", " ORDER BY StoreName", , "EntryID")
    
End Sub

Public Function GetShelf(p_strItemID As String, ByRef p_lShelf_RefID As Long) As Boolean
    
    strItemID = p_strItemID
    txtItem = strItemID & " -  " & GetSingleStringValue("ItemName", "Items", " WHERE ItemID='" & strItemID & "'")
    
    Me.Show 1
     
    p_lShelf_RefID = lShelf_RefID
    GetShelf = Not bCancelled
    
End Function
