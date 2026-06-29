VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.1#0"; "ComboList.ocx"
Begin VB.Form frmNewEmpProductivityEntry 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Employee Productivity ..."
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8310
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   8310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   360
      Left            =   165
      TabIndex        =   0
      Top             =   435
      Width           =   8040
      _ExtentX        =   14182
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
   Begin VB.CommandButton cmdSaveNext 
      Caption         =   "Save && &New"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2625
      TabIndex        =   6
      Top             =   3870
      Width           =   1815
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   6615
      TabIndex        =   4
      Top             =   1800
      Width           =   1590
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save && Close"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4530
      TabIndex        =   7
      Top             =   3870
      Width           =   1815
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6420
      TabIndex        =   8
      Top             =   3870
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker DT 
      Height          =   360
      Left            =   165
      TabIndex        =   5
      Top             =   2985
      Width           =   1845
      _ExtentX        =   3254
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   92274691
      CurrentDate     =   39651
   End
   Begin ComboList.Usercontrol1 cmbArticle 
      Height          =   360
      Left            =   165
      TabIndex        =   1
      Top             =   1125
      Width           =   8040
      _ExtentX        =   14182
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
   Begin ComboList.Usercontrol1 cmbProcess 
      Height          =   360
      Left            =   165
      TabIndex        =   2
      Top             =   1800
      Width           =   4575
      _ExtentX        =   8070
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   4
      Left            =   4785
      TabIndex        =   14
      Top             =   1545
      Width           =   525
   End
   Begin MSForms.ComboBox cmbType 
      Height          =   360
      Left            =   4755
      TabIndex        =   3
      Top             =   1800
      Width           =   1860
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "3281;635"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Process"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   3
      Left            =   165
      TabIndex        =   13
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Article"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   1
      Left            =   165
      TabIndex        =   12
      Top             =   885
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Employee"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   0
      Left            =   180
      TabIndex        =   11
      Top             =   180
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   5
      Left            =   6630
      TabIndex        =   10
      Top             =   1545
      Width           =   390
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   2
      Left            =   180
      TabIndex        =   9
      Top             =   2730
      Width           =   510
   End
End
Attribute VB_Name = "frmNewEmpProductivityEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lEntryID As Long
Dim bSaved As Boolean

Private Sub cmbArticle_matched()
    cmbProcess.ClearVals
    cmbProcess.AddVals con, "'{' + Code + '} ' + Description", "VItemProcesses", "ProcessID", " WHERE IsExist='" & cmbArticle.ID & "' ORDER BY ItemSNo"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function Saved() As Boolean

    On Error GoTo err
    
    If cmbEmp.MatchFound = False Then
        MsgBox "Please Select Employee.", vbInformation
        Exit Function
    ElseIf cmbArticle.MatchFound = False Then
        MsgBox "Please Select Article.", vbInformation
        Exit Function
    ElseIf cmbProcess.MatchFound = False Then
        MsgBox "Please Select Process.", vbInformation
        Exit Function
    ElseIf cmbType.ListIndex = -1 Then
        MsgBox "Please Select Type.", vbInformation
        Exit Function
    ElseIf Val(txtQty) <= 0 Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Function
    End If
    
     
    If lEntryID = 0 Then
        con.Execute "INSERT INTO EmpProductivityReport(EmpID,ItemID,ProcessID,DT,Qty,iType)" & _
         " VALUES('" & cmbEmp.ID & "','" & cmbArticle.ID & "'," & Val(cmbProcess.ID) & ",'" & DT & "'," & Val(txtQty) & "," & cmbType.ListIndex & ")"
    Else
'        con.Execute "UPDATE Vehicles SET RegNo='" & txtRegNo & "',RegDate='" & DTReg & "',RegName='" & txtRegName & _
'         "',EngineNo='" & txtEngineNo & "',ChasisNo='" & txtChasisNo & "',MakerName='" & txtMakerName & _
'         "',TypeOfBody='" & txtTypeOfBody & "',ManufactureDT='" & DTManufacture & "',HoursePower=" & Val(txtHoursePower) & _
'         ",SeatingCapacity=" & Val(txtSeatingCapacity) & ",PurchaseDT='" & DTPurchase & "',PurchaseValue=" & _
'         Val(txtPurchaseValue) & ",PurchaseFrom='" & txtPurchaseFrom & "',InsuranceCompany='" & txtInsuranceCompany & _
'         "',UserName='" & txtUserName & "' WHERE EntryID=" & lEntryID
    End If
    
    bSaved = True
    Saved = True
    
    Exit Function
err:
    MsgBox err.Description
    
End Function

Private Sub cmdSave_Click()
    If Saved Then Unload Me
End Sub


Private Sub cmdSaveNext_Click()
    If Saved Then
        Unload Me
        Me.Show 1
    End If
End Sub

Private Sub Form_Load()

    lEntryID = 0
    DT = Date
    
    bSaved = False
    
    cmbType.AddItem "Normal"
    cmbType.AddItem "Over Time"
    
    cmbEmp.ListHeight = 2400
    cmbEmp.AddVals con, "'{' + EmpID + '} ' + Name", "VEmp", "EmpID", " WHERE Active=1 ORDER BY EmpID"
    
    cmbArticle.ListHeight = 2400
    cmbArticle.AddVals con, "'{' + ItemID + '} ' + ItemName", "Items", "ItemID", " WHERE InActive=0 ORDER BY ItemID"
    
    cmbProcess.ListHeight = 1800
    
    
    
End Sub

Public Function ShowMe(strVehicleClass As String) As Boolean
    
    Me.Show 1
    ShowMe = bSaved
    
End Function
