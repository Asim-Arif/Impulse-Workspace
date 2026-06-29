VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmPrintCatalog_Frgz 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Catalog"
   ClientHeight    =   3510
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbGroup 
      Height          =   285
      Left            =   45
      TabIndex        =   2
      Top             =   375
      Width           =   1740
      _ExtentX        =   3069
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
      Appearance      =   0
   End
   Begin ComboList.Usercontrol1 cmbItemCat 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   375
      Width           =   1740
      _ExtentX        =   3069
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
      Appearance      =   0
   End
   Begin ComboList.Usercontrol1 cmbStore 
      Height          =   285
      Left            =   45
      TabIndex        =   11
      Top             =   945
      Width           =   3495
      _ExtentX        =   6165
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
      Appearance      =   0
   End
   Begin ComboList.Usercontrol1 cmbAdditionalGroup 
      Height          =   285
      Left            =   60
      TabIndex        =   12
      Top             =   1515
      Width           =   3495
      _ExtentX        =   6165
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
      Appearance      =   0
   End
   Begin VB.CheckBox chkShowLocationwise 
      Caption         =   "Show Locationwise stock"
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
      Height          =   255
      Left            =   2790
      TabIndex        =   9
      Top             =   3255
      Width           =   2535
   End
   Begin VB.CheckBox chkZero 
      Caption         =   "Print Zero Qty"
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
      Height          =   255
      Left            =   2790
      TabIndex        =   8
      Top             =   2910
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CheckBox chkInactive 
      Caption         =   "Inactive Articles"
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
      Height          =   255
      Left            =   2790
      TabIndex        =   7
      Top             =   2550
      Width           =   2535
   End
   Begin VB.OptionButton OptSort 
      Caption         =   "Sort By Item Name"
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
      Height          =   285
      Index           =   1
      Left            =   45
      TabIndex        =   6
      Top             =   2910
      Width           =   2280
   End
   Begin VB.OptionButton OptSort 
      Caption         =   "Sort By Catalog Code"
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
      Height          =   285
      Index           =   0
      Left            =   45
      TabIndex        =   5
      Top             =   2550
      Value           =   -1  'True
      Width           =   2280
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Additional Group"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   13
      Top             =   1275
      Width           =   3495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Store"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   0
      Left            =   45
      TabIndex        =   10
      Top             =   705
      Width           =   3495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Item Group                  Item Category"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   7
      Left            =   45
      TabIndex        =   3
      Top             =   105
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   4395
      Picture         =   "frmPrintCatalog_Frgz.frx":0000
      Top             =   1110
      Width           =   480
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   3630
      TabIndex        =   1
      Top             =   555
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Cancel       "
      PicturePosition =   327683
      Size            =   "3096;635"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdPrint 
      Height          =   360
      Left            =   3630
      TabIndex        =   0
      Top             =   105
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Print           "
      PicturePosition =   327683
      Size            =   "3096;635"
      Picture         =   "frmPrintCatalog_Frgz.frx":05EF
      Accelerator     =   71
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmPrintCatalog_Frgz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    
    
    If Not (cmbGroup.MatchFound And cmbItemCat.MatchFound And cmbAdditionalGroup.MatchFound) Then
        MsgBox "Invalid Group Or Category Selected.", vbInformation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    'Dim rpt As New rptCompanyCatalog
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\rptCompanyCatalog.rpt") '
    
    
    Dim Cond As String
    Cond = " WHERE 1=1"
    If cmbGroup.ID <> "0" Then
        Cond = Cond & " AND VItems.GroupID=" & cmbGroup.ID
    End If
    
    If cmbAdditionalGroup.ID <> "0" Then
        Cond = Cond & " AND VItems.MainGroupID=" & cmbAdditionalGroup.ID
    End If
    
    If cmbItemCat.ID <> "0" Then
        Cond = Cond & " AND VItems.CatID=" & cmbItemCat.ID
    End If

    If chkInactive.Value = vbChecked Then
        Cond = Cond & " AND VItems.InActive=TRUE"
    End If

    If chkZero.Value = vbUnchecked Then
        Cond = Cond & " AND VItems.InHand>0"
    End If

    If cmbStore.ID <> "0" Then
        Cond = Cond & " AND VItems.ItemID IN(SELECT ItemID FROM VItemsWithShelfWiseStock WHERE Store_RefID=" & cmbStore.ID & " AND NetQty>0)"
    End If
'    rpt.RecordSelectionFormula = Cond
'
    If OptSort(0).Value Then
        'Dim CRXDatabaseField As CRAXDRT.DatabaseFieldDefinition
        'Set CRXDatabaseField = rpt.Database.Tables.Item(1).Fields.Item(3)
        'rpt.RecordSortFields.Item(1).Field = CRXDatabaseField
        Cond = Cond & " ORDER BY VItems.ItemID"
    Else
        Cond = Cond & " ORDER BY VItems.ItemName"
    End If
    
    If chkShowLocationwise.Value = vbChecked Then
        rpt.FormulaFields.GetItemByName("ShowLocationWiseStock").Text = True
        rpt.FormulaFields.GetItemByName("StoreEntryID").Text = cmbStore.ID
    Else
        rpt.FormulaFields.GetItemByName("ShowLocationWiseStock").Text = False
    End If
    
    Me.Hide
    Load frmPrevRpt
    frmPrevRpt.ShowReport "", rpt, "SELECT VItems.*,ItemCatagories.* FROM VItems INNER JOIN ItemCatagories ON ItemCatagories.CatID=VItems.CatID " & Cond
    Unload Me
    Screen.MousePointer = vbDefault

    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
        
    cmbGroup.AddItem "<All Groups>", "0"
    Call cmbGroup.AddVals(con, "Description", "ItemGroups", "ID")
    cmbGroup.ID = "0"
    
    cmbAdditionalGroup.AddItem "<All>", "0"
    Call cmbAdditionalGroup.AddVals(con, "MainGroupName", "ItemGroupsMain", "MainGroupID")
    cmbAdditionalGroup.ID = "0"
    
    cmbGroup.ListHeight = 1200
    cmbItemCat.ListHeight = 1200
    
    cmbItemCat.AddItem "<All Categories>", "0"
    Call cmbItemCat.AddVals(con, "Description", "ItemCatagories", "CatID")
    cmbItemCat.ID = "0"
    
    cmbStore.ListHeight = 1200
    cmbStore.AddItem "<All Stores>", "0"
    cmbStore.AddVals con, "StoreName", "Stores", "EntryID"
    cmbStore.ID = "0"
    
End Sub
