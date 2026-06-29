VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmFixedAssets_RepairLog_AE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fixed Assets Repair Log"
   ClientHeight    =   4095
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
   ScaleHeight     =   4095
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtBillReference 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   6060
      TabIndex        =   5
      Top             =   990
      Width           =   1320
   End
   Begin VB.TextBox txtRepairCost 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   4725
      TabIndex        =   4
      Top             =   990
      Width           =   1320
   End
   Begin VB.CommandButton cmdSaveNext 
      Caption         =   "Sa&ve && New"
      Height          =   495
      Left            =   3525
      TabIndex        =   7
      Top             =   3480
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
      Width           =   1320
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save && Close"
      Height          =   495
      Left            =   4845
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   6150
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.TextBox txtRemarks 
      Appearance      =   0  'Flat
      Height          =   1485
      Left            =   135
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Width           =   7245
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
      Format          =   213385219
      CurrentDate     =   37055
   End
   Begin MSComCtl2.DTPicker DTSchedule 
      Height          =   315
      Left            =   3045
      TabIndex        =   3
      Top             =   990
      Width           =   1695
      _ExtentX        =   2990
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
      Format          =   213385219
      CurrentDate     =   37055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Bill Ref."
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
      Left            =   6060
      TabIndex        =   16
      Top             =   810
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Repair Cost"
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
      Left            =   4740
      TabIndex        =   15
      Top             =   795
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Next Schedule Date"
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
      Left            =   3060
      TabIndex        =   14
      Top             =   795
      Width           =   1650
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
      TabIndex        =   13
      Top             =   795
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Remarks (Enter Parts and Labor Details)"
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
      Left            =   150
      TabIndex        =   12
      Top             =   1575
      Width           =   3420
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
Attribute VB_Name = "frmFixedAssets_RepairLog_AE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If Saved Then Unload Me
End Sub

Private Sub cmdSaveNext_Click()
    If Saved Then
        txtRemarks = ""
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
    
    If DTSchedule.Value < DT.Value Then
        MsgBox "Invalid Next Schedule Date.", vbInformation
        Exit Function
    End If
    
    con.Execute "INSERT INTO FixedAssets_RepairLog(AssetID,DT,Remarks,MeterReading,ScheduleDT,CostOfRepair,BillReference,UserName" & _
     ",MachineName) VALUES('" & cmbAsset.ID & "','" & DT.Value & "','" & txtRemarks & "'," & _
     Val(txtMeterReading) & ",'" & DTSchedule.Value & "'," & Val(txtRepairCost) & ",'" & txtBillReference & "','" & CurrentUserName & "','" & strComputerName & "')"
     
    Saved = True
    bCancelled = False
    
    Exit Function
err:
    MsgBox err.Description
End Function

Private Sub Form_Load()

    DT.Value = Date
    DTSchedule.Value = DateAdd("d", 7, Date)
    
    cmbAsset.ClearVals
    cmbAsset.ListHeight = 4000
    cmbAsset.AddVals con, "'{' + AssetID + '} ' + Description", "FixedAssets", "AssetID", " ORDER BY AssetID"
   
    bCancelled = True
    
End Sub

Public Function AddNew() As Boolean
    Me.Show 1
    AddNew = Not bCancelled
End Function
