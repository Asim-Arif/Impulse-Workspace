VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmPrintProcessWiseStockReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Stock Report"
   ClientHeight    =   8025
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   13470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrintProcessWiseStockReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   13470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbItemGroup 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   300
      Width           =   3165
      _ExtentX        =   5583
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
   End
   Begin ComboList.Usercontrol1 cmbItems 
      Height          =   285
      Left            =   3285
      TabIndex        =   2
      Top             =   300
      Width           =   10125
      _ExtentX        =   17859
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
   End
   Begin ComboList.Usercontrol1 cmbItemSearch 
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   1110
      Width           =   13290
      _ExtentX        =   23442
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
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Combined"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   345
      Index           =   2
      Left            =   2520
      TabIndex        =   9
      Top             =   7500
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.OptionButton Option1 
      Caption         =   "R2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   345
      Index           =   1
      Left            =   1260
      TabIndex        =   6
      Top             =   7500
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "R1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   345
      Index           =   0
      Left            =   150
      TabIndex        =   5
      Tag             =   "2"
      Top             =   7500
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ListView LV 
      Height          =   6045
      Left            =   120
      TabIndex        =   4
      Top             =   1425
      Width           =   13290
      _ExtentX        =   23442
      _ExtentY        =   10663
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
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S #"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Item"
         Object.Width           =   21167
      EndProperty
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   12180
      TabIndex        =   1
      Top             =   7605
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&Print"
      Height          =   375
      Left            =   10890
      TabIndex        =   0
      Top             =   7605
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Search Grid"
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
      Left            =   120
      TabIndex        =   11
      Top             =   885
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Group"
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
      TabIndex        =   8
      Top             =   60
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Item"
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
      Left            =   3285
      TabIndex        =   3
      Top             =   60
      Width           =   420
   End
End
Attribute VB_Name = "frmPrintProcessWiseStockReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean
Dim strReturn As String

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub cmbItemGroup_matched()
    
    cmbItemSearch.ClearVals
    cmbItems.ClearVals
    cmbItems.DropDownOffLine = True
    cmbItemSearch.DropDownOffLine = True
    
    If cmbItemGroup.ID = "0" Then
        cmbItems.AddVals con, "'{' + ItemID + '} ' + ItemName + ' {' + ItemGroup + '}'", "VItems", "ItemID"
    Else
        cmbItems.AddVals con, "'{' + ItemID + '} ' + ItemName + ' {' + ItemGroup + '}'", "VItems", "ItemID", " WHERE GroupID=" & cmbItemGroup.ID
    End If
    
    Dim rs As New ADODB.Recordset, i As Integer
    Dim ITM As ListItem
    With rs
        If cmbItemGroup.ID = "0" Then
            .Open "SELECT ItemID,'{' + ItemID + '} ' + ItemName + ' {' + ItemGroup + '}' AS ItemName FROM VItems INNER JOIN (SELECT CompItemCode,SUM((Qty-ShippedQty)) AS OrderQty FROM VFOrderItemswithShippedQty WHERE (Qty-ShippedQty)>0 GROUP BY CompItemCode) T1 ON  VItems.ItemID=T1.CompItemCode ORDER BY OrderQty", con, adOpenForwardOnly, adLockReadOnly
            cmbItemSearch.AddVals con, "'{' + ItemID + '} ' + ItemName + ' {' + ItemGroup + '}'", "VItems INNER JOIN (SELECT CompItemCode,SUM((Qty-ShippedQty)) AS OrderQty FROM VFOrderItemswithShippedQty WHERE (Qty-ShippedQty)>0 GROUP BY CompItemCode) T1 ON  VItems.ItemID=T1.CompItemCode", "ItemID", " ORDER BY OrderQty"
        Else
            .Open "SELECT ItemID,'{' + ItemID + '} ' + ItemName + ' {' + ItemGroup + '}' AS ItemName FROM VItems INNER JOIN (SELECT CompItemCode,SUM((Qty-ShippedQty)) AS OrderQty FROM VFOrderItemswithShippedQty WHERE (Qty-ShippedQty)>0 GROUP BY CompItemCode) T1 ON  VItems.ItemID=T1.CompItemCode WHERE GroupID=" & cmbItemGroup.ID & " ORDER BY OrderQty", con, adOpenForwardOnly, adLockReadOnly
            cmbItemSearch.AddVals con, "'{' + ItemID + '} ' + ItemName + ' {' + ItemGroup + '}'", "VItems INNER JOIN (SELECT CompItemCode,SUM((Qty-ShippedQty)) AS OrderQty FROM VFOrderItemswithShippedQty WHERE (Qty-ShippedQty)>0 GROUP BY CompItemCode) T1 ON  VItems.ItemID=T1.CompItemCode", "ItemID", " WHERE GroupID=" & cmbItemGroup.ID & " ORDER BY OrderQty"
        End If
        LV.ListItems.Clear
        
        Do Until .EOF
            Set ITM = LV.ListItems.add(, !ITemID & "'", LV.ListItems.count + 1)
            ITM.Tag = !ITemID
            ITM.ListSubItems.add , , !ItemName & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Dim lCount As Long
    If LV.ListItems.count >= 20 Then
        lCount = 20
    Else
        lCount = LV.ListItems.count
    End If
    For i = 1 To lCount
        LV.ListItems(i).Checked = True
    Next
    
End Sub

Private Sub cmbitems_matched()
    
    On Error GoTo err
    
    Dim ITM As ListItem
    Set ITM = LV.ListItems.add(, cmbItems.ID & "'", LV.ListItems.count + 1)
    ITM.Tag = cmbItems.ID
    ITM.ListSubItems.add , , cmbItems.Text
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbItemSearch_matched()
    LV.ListItems(cmbItemSearch.ID & "'").Selected = True
    LV.ListItems(cmbItemSearch.ID & "'").EnsureVisible
    LV.SetFocus
End Sub

Private Sub Form_Load()
    strReturn = ""
    
    cmbItemGroup.ListHeight = 6000
    cmbItemGroup.AddItem "<All Groups>", "0"
    cmbItemGroup.AddVals con, "Description", "ItemGroups", "ID"
    cmbItemGroup.ID = "0"
    
    cmbItems.ListHeight = 6000
    
    
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    If LV.ListItems.count > 0 Then
        If KeyCode = 46 Then
            LV.ListItems.Remove LV.SelectedItem.Index
        End If
    End If
End Sub

Private Sub OKButton_Click()
    
'    If LV.ListItems.count > 20 Then
'        MsgBox "You Can't Have more than 20 Items.", vbInformation
'        Exit Sub
'    End If
    
    Dim lCount As Long, i As Integer
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            lCount = lCount + 1
        End If
    Next
    
    If lCount > 20 Then
        MsgBox "You Can't Have more than 20 Items.", vbInformation
        Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    If Option1(0).Tag = 2 Then
        Set rpt = rptApp.OpenReport(RptPath & "\ProcessWiseStockReportR1R2Combined.rpt")
    Else
        Set rpt = rptApp.OpenReport(RptPath & "\ProcessWiseStockReport.rpt")
    End If
    
    rpt.ParameterFields.GetItemByName("@R1").SetCurrentValue CLng(Option1(0).Tag)
    Dim iCount As Integer
    iCount = 0
    For i = 1 To LV.ListItems.count
        If LV.ListItems(i).Checked Then
            iCount = iCount + 1
            rpt.ParameterFields.GetItemByName("@Item" & iCount).SetCurrentValue LV.ListItems(i).Tag
        End If
    Next
    
    For i = lCount + 1 To 20
        rpt.ParameterFields.GetItemByName("@Item" & i).SetCurrentValue ""
    Next
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    Me.Hide
    f.ShowReport "", rpt, "", False, , Me
    
End Sub

Private Sub txtComments_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call OKButton_Click
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
    Option1(0).Tag = Index
End Sub
