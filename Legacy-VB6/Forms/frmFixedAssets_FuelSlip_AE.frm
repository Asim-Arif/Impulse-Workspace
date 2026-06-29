VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmFixedAssets_FuelSlip_AE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fuel Slip"
   ClientHeight    =   6300
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7530
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   795
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2175
      Width           =   7245
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   135
      TabIndex        =   5
      Top             =   1575
      Width           =   7245
      _ExtentX        =   12779
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
   Begin MSComctlLib.ListView LVHistory 
      Height          =   2490
      Left            =   135
      TabIndex        =   16
      Top             =   3210
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   4392
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
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "S #"
         Object.Width           =   1147
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2364
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Meter Reading"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Fuel"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Qty"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Emp ID"
         Object.Width           =   1764
      EndProperty
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   6105
      TabIndex        =   4
      Top             =   990
      Width           =   1275
   End
   Begin VB.CommandButton cmdSaveNext 
      Caption         =   "Sa&ve && New"
      Height          =   495
      Left            =   3525
      TabIndex        =   7
      Top             =   5760
      Width           =   1215
   End
   Begin ComboList.Usercontrol1 cmbAsset 
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   405
      Width           =   7245
      _ExtentX        =   12779
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
   Begin VB.TextBox txtMeterReading 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1695
      TabIndex        =   2
      Top             =   990
      Width           =   1485
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save && Close"
      Height          =   495
      Left            =   4845
      TabIndex        =   8
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   6150
      TabIndex        =   9
      Top             =   5760
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker DT 
      Height          =   315
      Left            =   135
      TabIndex        =   1
      Top             =   990
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   116064259
      CurrentDate     =   37055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   135
      TabIndex        =   18
      Top             =   1980
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Short History:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   3015
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Employee"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   135
      TabIndex        =   15
      Top             =   1365
      Width           =   825
   End
   Begin MSForms.ComboBox cmbFuel 
      Height          =   300
      Left            =   3195
      TabIndex        =   3
      Top             =   990
      Width           =   2895
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5106;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fuel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   3195
      TabIndex        =   14
      Top             =   810
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   6105
      TabIndex        =   13
      Top             =   795
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Meter Reading"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   1710
      TabIndex        =   12
      Top             =   795
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   165
      TabIndex        =   11
      Top             =   795
      Width           =   405
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Asset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   165
      TabIndex        =   10
      Top             =   210
      Width           =   480
   End
End
Attribute VB_Name = "frmFixedAssets_FuelSlip_AE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean

Private Sub cmbAsset_matched()

    Dim ITM As ListItem
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT TOP 10 * FROM FixedAssets_FuelSlips WHERE AssetID='" & cmbAsset.ID & "' ORDER BY EntryID DESC", con, adOpenForwardOnly, adLockReadOnly
        LVHistory.ListItems.Clear
        Do Until .EOF
            Set ITM = LVHistory.ListItems.add(, , LVHistory.ListItems.count + 1)
            ITM.ListSubItems.add , , Format(!DT, "dd-MMM-yyyy")
            ITM.ListSubItems.add , , Val(!MeterReading & "")
            ITM.ListSubItems.add , , !Fuel & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , !EmpId & ""
            .MoveNEXT
        Loop
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Saved Then Unload Me
End Sub

Private Sub cmdSaveNext_Click()
    If Saved Then
        cmbAsset.ID = ""
        cmbAsset.SetFocus
    End If
End Sub

Private Function Saved() As Boolean

    On Error GoTo err
     
    If cmbAsset.MatchFound = False Then
        MsgBox "Please Select Asset.", vbInformation
        Exit Function
    End If
    
    If cmbEmp.MatchFound = False Then
        MsgBox "Select Employee.", vbInformation
        Exit Function
    End If
    
    If cmbFuel.MatchFound = False Then
        MsgBox "Select Fuel", vbInformation
        Exit Function
    End If
    
    con.Execute "INSERT INTO FixedAssets_FuelSlips(AssetID,DT,MeterReading,Qty,EmpID,UserName" & _
     ",MachineName,Fuel,Remarks) VALUES('" & cmbAsset.ID & "','" & DT.Value & "'," & Val(txtMeterReading) & _
     "," & Val(txtQty) & ",'" & cmbEmp.ID & "','" & CurrentUserName & "','" & strComputerName & "','" & cmbFuel.Text & "','" & txtRemarks & "')"
     
    Saved = True
    bCancelled = False
    
    Exit Function
err:
    MsgBox err.Description
End Function

Private Sub Form_Load()

    DT.Value = Date
     
    cmbAsset.ClearVals
    cmbAsset.ListHeight = 4000
    cmbAsset.AddVals con, "'{' + AssetID + '} ' + Description", "FixedAssets", "AssetID", " ORDER BY AssetID"
   
    bCancelled = True
    
    With cmbFuel
        .AddItem "Super"
        .AddItem "Diesel"
        .AddItem "Lubrication"
        .AddItem "Other"
    End With
        
    cmbEmp.ClearVals
    cmbEmp.ListHeight = 2300
    cmbEmp.AddVals con, "'{' + EmpID + '} ' + Name + ' {' + Designation + '} ' + '{' + DeptName + '}'", "VEmp", "EmpID"
    
End Sub

Public Function AddNew() As Boolean
    Me.Show 1
    AddNew = Not bCancelled
End Function
