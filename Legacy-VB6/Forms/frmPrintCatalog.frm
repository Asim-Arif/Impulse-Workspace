VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmPrintCatalog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Catalog"
   ClientHeight    =   3255
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkValuationRackwise 
      Caption         =   "Valuation Report (Rackwise)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2790
      TabIndex        =   12
      Top             =   2865
      Width           =   2715
   End
   Begin VB.CheckBox chkPrintCost 
      Caption         =   "Valuation Report"
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
      TabIndex        =   11
      Top             =   2505
      Width           =   2535
   End
   Begin VB.CheckBox chkGroupByLocation 
      Caption         =   "Group by Location"
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
      TabIndex        =   10
      Top             =   2175
      Width           =   2535
   End
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
      Left            =   2775
      TabIndex        =   9
      Top             =   1845
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
      Top             =   1500
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
      Top             =   1140
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
      Top             =   1500
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
      Top             =   1140
      Value           =   -1  'True
      Width           =   2280
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
      Left            =   2235
      Picture         =   "frmPrintCatalog.frx":0000
      Top             =   675
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
      Left            =   3615
      TabIndex        =   0
      Top             =   105
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Print           "
      PicturePosition =   327683
      Size            =   "3096;635"
      Picture         =   "frmPrintCatalog.frx":05EF
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
Attribute VB_Name = "frmPrintCatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    On Error GoTo err
    
    
    If Not (cmbGroup.MatchFound And cmbItemCat.MatchFound) Then
        MsgBox "Invalid Group Or Category Selected.", vbInformation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    'Dim rpt As New rptCompanyCatalog
    Dim rpt As CRAXDDRT.Report
    
    If chkValuationRackwise.Value = vbChecked Then
        Dim strLocation As String, frmLoc As New frmGetSingleTextValue
        If frmLoc.getTextValue(strLocation) = False Then Exit Sub
        
        Set rpt = rptApp.OpenReport(RptPath & "\CompanyCatalogValuationRackwise.rpt")
    ElseIf chkPrintCost.Value = vbChecked Then
        Set rpt = rptApp.OpenReport(RptPath & "\CompanyCatalogValuation.rpt") '
    ElseIf chkGroupByLocation.Value = vbChecked Then
        Set rpt = rptApp.OpenReport(RptPath & "\CompanyCatalogLocationwise.rpt") '
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\rptCompanyCatalog.rpt") '
    End If
    
    Dim Cond As String
    Cond = ""
    If cmbGroup.ID <> "0" Then
        Cond = "{VEmps.GroupID}=" & cmbGroup.ID
    End If
    
    If cmbItemCat.ID <> "0" Then
        If Cond <> "" Then
            Cond = Cond & " And "
        End If
        Cond = Cond & "{VEmps.CatID}=" & cmbItemCat.ID
    End If
     
    If chkInactive.Value = vbChecked Then
        If Cond <> "" Then
            Cond = Cond & " And "
        End If
        Cond = Cond & "{VEmps.InActive}=TRUE"
    End If
    If chkZero.Value = vbUnchecked Then
        If Cond <> "" Then
            Cond = Cond & " And "
        End If
        If chkGroupByLocation.Value = vbChecked Then
            Cond = Cond & "{VItemsWithLocationwiseStock.NetQty}>0"
        Else
            Cond = Cond & "{VItemsInHandQty.InHand}>0"
        End If
    End If
    
    If strLocation <> "" Then
        If Cond <> "" Then
            Cond = Cond & " AND "
        End If
        Cond = Cond & " {VItemsWithLocationwiseStock_Rackwise.Rack}='" & strLocation & "'"
    End If
    
    rpt.RecordSelectionFormula = Cond
    
    If OptSort(0).Value Then
        Dim CRXDatabaseField As CRAXDRT.DatabaseFieldDefinition
        Set CRXDatabaseField = rpt.Database.Tables.Item(1).Fields.Item(3)
        rpt.RecordSortFields.Item(1).Field = CRXDatabaseField
    End If
    
    If chkGroupByLocation.Value = vbUnchecked Then
        If chkShowLocationwise.Value = vbChecked Then
            rpt.FormulaFields.GetItemByName("ShowLocationWiseStock").Text = True
        Else
            rpt.FormulaFields.GetItemByName("ShowLocationWiseStock").Text = False
        End If
    End If
    
    Me.Hide
    Load frmPrevRpt
    frmPrevRpt.ShowReport Cond, rpt
    Unload Me
    Screen.MousePointer = vbDefault

    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
        
    cmbGroup.AddItem "<All Groups>", "0"
    Call cmbGroup.AddVals(con, "Description", "ItemGroups", "ID")
    
    cmbGroup.ListHeight = 1200
    cmbItemCat.ListHeight = 1200
    
    cmbItemCat.AddItem "<All Categories>", "0"
    Call cmbItemCat.AddVals(con, "Description", "ItemCatagories", "CatID")
    
End Sub
