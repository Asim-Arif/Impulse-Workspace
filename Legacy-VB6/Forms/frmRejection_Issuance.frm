
VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmRejection_Issuance 
   ClientHeight    =   7560
   ClientLeft      =   1110
   ClientTop       =   60
   ClientWidth     =   11670
   ControlBox      =   0   'False
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
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   11670
   WindowState     =   2  'Maximized
   Begin ComboList.Usercontrol1 cmbLocation 
      Height          =   285
      Left            =   3165
      TabIndex        =   3
      Top             =   1425
      Width           =   7980
      _ExtentX        =   14076
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
      Left            =   405
      TabIndex        =   2
      Top             =   1425
      Width           =   2745
      _ExtentX        =   4842
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
      Enabled         =   0   'False
      Appearance      =   0
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   405
      TabIndex        =   4
      Top             =   7020
      Width           =   5325
   End
   Begin MSComCtl2.DTPicker DT 
      Height          =   285
      Left            =   9645
      TabIndex        =   14
      Top             =   885
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   163840003
      CurrentDate     =   38535
   End
   Begin VB.TextBox txtIssuedTo 
      Height          =   300
      Left            =   405
      TabIndex        =   1
      Top             =   885
      Width           =   9240
   End
   Begin VB.TextBox txtIssNo 
      Height          =   300
      Left            =   9885
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   345
      Visible         =   0   'False
      Width           =   1725
   End
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   5850
      TabIndex        =   8
      Top             =   6810
      Width           =   5280
      Begin MSForms.CommandButton cmdSave 
         Height          =   360
         Left            =   75
         TabIndex        =   11
         Top             =   195
         Width           =   1665
         Caption         =   "Save & New   "
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   110
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdCancel 
         CausesValidation=   0   'False
         Height          =   360
         Left            =   3510
         TabIndex        =   10
         Top             =   195
         Width           =   1665
         Caption         =   "Cancel            "
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   67
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdSaveClose 
         Height          =   360
         Left            =   1785
         TabIndex        =   9
         Top             =   180
         Width           =   1665
         Caption         =   "Save & Close  "
         PicturePosition =   327683
         Size            =   "2937;635"
         Accelerator     =   83
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
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
      TabIndex        =   5
      Top             =   0
      Width           =   10365
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rejection Store Issuance"
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
         Left            =   3720
         TabIndex        =   6
         Top             =   135
         Width           =   3780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rejection Store Issuance"
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
         Left            =   3735
         TabIndex        =   7
         Top             =   150
         Width           =   3780
      End
   End
   Begin MSComctlLib.ListView LV 
      Height          =   5085
      Left            =   405
      TabIndex        =   12
      Top             =   1725
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   8969
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   15199215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S No."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Bin No."
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Lot No."
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Item Code"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   4
         Text            =   "Balance"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Issd. Qty"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSForms.Label Label5 
      Height          =   240
      Left            =   405
      TabIndex        =   16
      Top             =   6795
      Width           =   5325
      BackColor       =   11517387
      Caption         =   " Remarks"
      Size            =   "9393;423"
      BorderColor     =   16512
      BorderStyle     =   1
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label4_1 
      Height          =   240
      Left            =   405
      TabIndex        =   15
      Top             =   1200
      Width           =   10740
      BackColor       =   11517387
      Caption         =   " Store                                                  Racks"
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
   Begin MSForms.Label Label3 
      Height          =   240
      Left            =   405
      TabIndex        =   13
      Top             =   645
      Width           =   10740
      BackColor       =   11517387
      Caption         =   $"frmRejection_Issuance.frx":0000
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
Attribute VB_Name = "frmRejection_Issuance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbLocation_matched()

    Dim rs As New ADODB.Recordset
    Dim ITM As ListItem
    With rs
        .Open "SELECT EntryNo,LotNo,ShelfNo,ItemCode,QtyRcvd,TotalIssued,Shelf_RefID FROM VRejection_Receiving WHERE Rack_RefID=" & cmbLocation.ID & " AND QtyRcvd-TotalIssued>0 ORDER BY EntryNo", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , LV.ListItems.count + 1)
            ITM.Tag = Val(!EntryNo & "")
            ITM.ListSubItems.add(, , !ShelfNo & "").Tag = Val(!Shelf_RefID & "")
            ITM.ListSubItems.add , , !LotNo & ""
            ITM.ListSubItems.add , , !ItemCode & ""
            ITM.ListSubItems.add , , Val(!QtyRcvd & "") - Val(!TotalIssued & "")
            ITM.ListSubItems.add , , ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub cmbStore_matched()

    cmbLocation.ClearVals
    cmbLocation.AddVals con, "RackNo", "StoreRacks", "EntryID", " WHERE Store_RefID=" & cmbStore.ID
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Saved Then
        Unload Me
'        LV.ListItems.Clear
'        Call DT_Change
        Me.Show
    End If
End Sub

Private Sub cmdSaveClose_Click()

    If Saved Then Unload Me
    
End Sub

Private Function Saved() As Boolean

    On Error GoTo err
     
'    If DT > GetServerDate Then
'        MsgBox "Sorry, You cannot issue in future dates.", vbCritical
'        Exit Function
'    End If
    
    If Trim(txtIssuedTo) = "" Then
        MsgBox "Please Enter Issued To.", vbInformation
        Exit Function
    End If
    
    Dim lRefID As Long, i As Integer
    
    Call DT_Change   'To ReAssign Issuance No.
    Call StartTrans(con)
    
    con.Execute "INSERT INTO Rejection_Issuance(IssNo,DT,IssuedTo,StoreRacks_RefID,UserName,MachineName,Remarks) VALUES('" & _
     txtIssNo & "','" & DT & "','" & txtIssuedTo & "'," & cmbLocation.ID & ",'" & _
     CurrentUserName & "','" & strComputerName & "','" & txtRemarks & "')"
    
    lRefID = con.Execute("SELECT MAX(EntryID) FROM Rejection_Issuance").Fields(0).Value & ""
    
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            If .Checked Then
                con.Execute "INSERT INTO Rejection_Issuance_Detail(R_I_RefID,R_R_RefID,Qty,Shelf_RefID) VALUES(" & _
                 lRefID & "," & Val(.Tag) & "," & Val(.SubItems(5)) & "," & Val(.ListSubItems(1).Tag) & ")"
            End If
        End With
    Next
    con.CommitTrans
    Saved = True
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub DT_Change()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT MAX(CAST(RIGHT(IssNo,3) AS INT)) FROM IssItemsSimple WHERE DT='" & DT & "'", con, adOpenForwardOnly, adLockReadOnly
        
        txtIssNo = "RIS-" & Format(DT, "yyMMdd-") & Format(Val(.Fields(0) & "") + 1, "000")
        
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()

    'DT = Date
    'Call DT_Change
    'Call RestrictIssRcvDates(DT)
    DT = Date
    Call DT_Change
    If CurrentUserName = "Administrator" Then
        DT.Enabled = True
    Else
        DT.Enabled = False
    End If
    
    cmbLocation.ListHeight = 2400
'    cmbMaterial.DropDownOffLine = True
'    cmbMaterial.AddVals con, "ItemName + ' {' + ItemID + '}'", "Items", "ItemID", " Order By ItemID"
    
    'cmbDept.AddVals con, "Name", "Departments", "DeptID", " Where Active=1 Order By Name"
    'AddToCombo cmbPrinter, "PrinterName", "Printers", , , "EntryID"
    
    '82
    cmbStore.ListHeight = 4000
    'cmbStore.AddVals con, "DISTINCT StoreName", "VStoreShelfs", "Store_RefID", " WHERE Store_RefID IN(SELECT Store_RefID FROM Users_Stores WHERE UserID=" & UserID & ") AND EntryID IN(SELECT Shelf_RefID FROM VItemsWithShelfWiseStock WHERE ItemID='" & cmbMaterial.ID & "' AND NetQty>0)"
    cmbStore.AddVals con, "StoreName", "Stores", "EntryID", " WHERE EntryID=82"
    cmbStore.ID = 82
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    fTitle.Move (Me.ScaleWidth - fTitle.Width) / 2
    
'    Dim TitleBottom As Long
'
'    TitleBottom = cmbMaterial.Top + cmbMaterial.Height + 50 '+ fTitle.Height
 
End Sub

Private Sub LV_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Item.Checked Then
        Item.SubItems(5) = Val(Item.SubItems(4))
    Else
        Item.SubItems(5) = ""
    End If
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Do you want to Remove ?", vbQuestion + vbYesNo) = vbYes Then
            LV.ListItems.Remove LV.SelectedItem.Index
        End If
    End If
End Sub
