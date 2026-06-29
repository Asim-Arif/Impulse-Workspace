VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.0#0"; "ComboList.ocx"
Begin VB.Form frmContEmpWorkDone 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Temp. Employee's Daily Work Done"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5100
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
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5100
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCountry 
      Height          =   285
      Left            =   1800
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3120
   End
   Begin VB.TextBox txtCustCode 
      Height          =   285
      Left            =   1800
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1470
      Width           =   3120
   End
   Begin ComboList.Usercontrol1 cmbDept 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   45
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Top             =   390
      Width           =   3120
      _ExtentX        =   5503
      _ExtentY        =   503
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   735
      Width           =   3120
      _ExtentX        =   5503
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
      CustomFormat    =   "dddd dd-MMM-yyyy"
      Format          =   20643843
      CurrentDate     =   37384
   End
   Begin MSForms.ComboBox cmbOrderNo 
      Height          =   285
      Left            =   1800
      TabIndex        =   3
      Top             =   1125
      Width           =   3120
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5503;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Order No.                  :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   9
      Left            =   165
      TabIndex        =   20
      Top             =   1155
      Width           =   1590
   End
   Begin MSForms.ComboBox cmbItem 
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   2130
      Width           =   3120
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "5503;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Article #                     :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   8
      Left            =   150
      TabIndex        =   19
      Top             =   2160
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Country                     :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   7
      Left            =   180
      TabIndex        =   18
      Top             =   1830
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Customer                   :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   6
      Left            =   165
      TabIndex        =   17
      Top             =   1500
      Width           =   1605
   End
   Begin MSForms.TextBox txtDescription 
      Height          =   285
      Left            =   1800
      TabIndex        =   9
      Top             =   2445
      Width           =   3120
      VariousPropertyBits=   746604575
      Size            =   "5503;503"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description                 :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   3
      Left            =   150
      TabIndex        =   16
      Top             =   2460
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Per Piece Rate           :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   165
      TabIndex        =   15
      Top             =   2820
      Width           =   1605
   End
   Begin MSForms.TextBox txtRate 
      Height          =   315
      Left            =   1800
      TabIndex        =   10
      Top             =   2790
      Width           =   1020
      VariousPropertyBits=   746604571
      Size            =   "1799;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Date                :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   165
      TabIndex        =   14
      Top             =   780
      Width           =   1605
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   180
      X2              =   4880
      Y1              =   3270
      Y2              =   3270
   End
   Begin MSForms.TextBox txtQty 
      Height          =   315
      Left            =   3900
      TabIndex        =   5
      Top             =   2790
      Width           =   1020
      VariousPropertyBits=   746604571
      Size            =   "1799;556"
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Worked Qty :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   5
      Left            =   2880
      TabIndex        =   13
      Top             =   2835
      Width           =   975
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   360
      Left            =   1800
      TabIndex        =   7
      Top             =   3375
      Width           =   1710
      Caption         =   "Save & Close"
      PicturePosition =   327683
      Size            =   "3016;635"
      Picture         =   "frmContEmpWorkDone.frx":0000
      Accelerator     =   83
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdNext 
      Height          =   360
      Left            =   150
      TabIndex        =   6
      Top             =   3375
      Width           =   1590
      Caption         =   "Save & New    "
      PicturePosition =   327683
      Size            =   "2805;635"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   360
      Left            =   3555
      TabIndex        =   8
      Top             =   3375
      Width           =   1380
      Caption         =   "Cancel    "
      PicturePosition =   327683
      Size            =   "2434;635"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Employee        :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   195
      TabIndex        =   12
      Top             =   435
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select Department    :"
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   4
      Left            =   195
      TabIndex        =   11
      Top             =   105
      Width           =   1575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   180
      X2              =   4880
      Y1              =   3285
      Y2              =   3285
   End
End
Attribute VB_Name = "frmContEmpWorkDone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LongTerm As Boolean

Private Sub PrintReport()

    Screen.MousePointer = vbHourglass

    Dim f As New frmPrevRpt, rpt As CRAXDDRT.Report, ID As Integer
    Set rpt = rptApp.OpenReport(rptPayrollPath & "\rptAdvancesPayslip.rpt")
    ID = Val(con.Execute("Select MAX(EntryID) FROM Advances").Fields(0) & "")

    Dim FormulaFields As CRAXDDRT.FormulaFieldDefinitions
    Dim FormulaField As CRAXDDRT.FormulaFieldDefinition

    Set FormulaFields = rpt.FormulaFields

    For Each FormulaField In FormulaFields
        If FormulaField.Name = "{@Company}" Then
            FormulaField.Text = "'" & company & "'"
        ElseIf FormulaField.Name = "{@FromTo}" Then
            FormulaField.Text = "'From " & Format(DTFrom, "dd-MMM-yyyy") & " To " & Format(DTTo, "dd-MMM-yyyy") & "'"
        End If
    Next

    Me.Hide
    f.ShowReport "{VAdvances.EntryID}=" & ID & " ", rpt

    Screen.MousePointer = vbDefault

End Sub

Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select ArticleNo,ItemName,ItemID From VFCustomerCatalogActive Where CustCode='" & cmbCustCode.ID & "' AND Country='" & cmbCountry & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbItem.Clear
        Do Until .EOF
            cmbItem.AddItem !ArticleNo & ""
            cmbItem.List(.AbsolutePosition - 1, 1) = !ItemName & ""
            cmbItem.List(.AbsolutePosition - 1, 2) = !ItemID & ""
            .MoveNext
        Loop
        .Close
    End With
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbCustCode_matched()

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    
    With rs
        .Open "Select Country From ForeignCustomers Where CustCode='" & cmbCustCode.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        
        cmbCountry.Clear
        
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            .MoveNext
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbdept_matched()

    cmbEmp.ClearVals
    
    cmbEmp.AddVals con, "Name", "Employees", "EmpID", "Where DeptID='" & cmbDept.ID & "' AND Active=1 AND PerPieceRate=1"
        
End Sub

Private Function Saved() As Boolean

    Saved = False
    
    On Error GoTo err
    
    Dim SQL As String
    
    If Not cmbEmp.MatchFound Then
        MsgBox "Please Select Valid Employee From The List.", vbInformation
        Exit Function
    End If
    
    If IsEmpPresent(cmbEmp.ID, DTPicker1) = False Then
        MsgBox "Employee Is Not Present On Selected Date.", vbInformation
        Exit Function
    End If
    
    If Val(txtQty) <= 0 Then
        MsgBox "Invalid Qty.", vbInformation
        Exit Function
    End If
        
    'con.Execute "Update EmpWork Set QtyWorked=" & Val(txtOT) & " where EmpID='" & cmbEmp.ID & "' and dt='" & Format(DTPicker1, "dd-MMM-yyyy") & "'", a
        
    con.Execute "Insert Into EmpPerPieceWorkDone(EmpID,DT,CustCode,Country,ItemID,Rate,Qty,Description,OrderNo) Values ('" & _
     cmbEmp.ID & "','" & Format(DTPicker1, "dd-MMM-yyyy") & "','" & txtCustCode & "','" & txtCountry & "','" & _
     cmbItem.List(cmbItem.ListIndex, 2) & "'," & Val(txtRate) & _
     "," & Val(txtQty) & ",'" & txtDescription & "','" & cmbOrderNo.Text & "')"
        
    Saved = True
   
    SQL = ""
    Exit Function
    
err:
    MsgBox err.Description

End Function


Private Sub cmbEmp_matched()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Rate From Employees Where EmpID='" & cmbEmp.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        txtRate = Val(!Rate & "")
        .Close
    End With
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub cmbItem_Change()
    Call cmbItem_Click
End Sub

Private Sub cmbItem_Click()
    If cmbItem.MatchFound = False Then Exit Sub
    txtDescription = cmbItem.List(cmbItem.ListIndex, 1)
End Sub

Private Sub cmbOrderNo_Change()
    Call cmbOrderNo_Click
End Sub

Private Sub cmbOrderNo_Click()

    On Error GoTo err
    
    If cmbOrderNo.MatchFound Then
        txtCustCode = cmbOrderNo.List(cmbOrderNo.ListIndex, 1)
        txtCountry = cmbOrderNo.List(cmbOrderNo.ListIndex, 2)
        
        AddToCombo cmbItem, "ArticleNo,ItemName,ItemID", "VFOrderItems", " Where OrderNo='" & cmbOrderNo & "'"
    End If
    
    Exit Sub
    
err:
    MsgBox err.Description
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()

    If Saved Then
        
        txtQty = ""
        cmbDept.SetFocus
        
    End If
    
End Sub

Private Sub cmdOK_Click()
    If Saved Then Unload Me
End Sub


Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}"
    End If
End Sub

Private Sub Form_Load()

    cmbDept.ListHeight = 1600
    cmbEmp.ListHeight = 1250
    
    cmbDept.ClearVals
    
    cmbDept.AddVals con, "Name", "Departments", "DeptID", " Where TempDept=1"
    
    
    AddToCombo cmbOrderNo, "OrderNo,CustCode,Country", "FCustomerOrders"
    
    DTPicker1 = GetServerDate

End Sub

Private Sub txtAmt_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = OnlyNumber(KeyAscii.value, True)
End Sub

Private Sub txtDeduct_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = OnlyNumber(KeyAscii.value, True)
End Sub

'To Check The Long Term Settings That Are Made For Departments
Private Function IsNotLongApplicable() As Boolean

    Dim Rec As New ADODB.Recordset
    'dim Item as New
    
    With Rec
        .CursorLocation = adUseClient
        .Open "SELECT JoinDate,LongTermTime,LongTermLimit FROM VOnlyEmps WHERE EmpID='" & cmbEmp.ID & "'", con, adOpenKeyset, adLockReadOnly
        If DateDiff("m", !JoinDate, DTPicker1) < !longtermtime Then
            MsgBox "Could Not Give Long Term Loan To Employee." & vbNewLine & "It Can Be Given After " & !longtermtime & " Months", vbInformation
            IsNotLongApplicable = True
        ElseIf Val(txtAmt) > !LongTermLimit Then
            MsgBox "Could Not Give Long Term Loan To Employee. Because Amount " & vbNewLine & " Is Greater Than Long Term Limit.", vbInformation
            IsNotLongApplicable = True
        End If
    End With

End Function

