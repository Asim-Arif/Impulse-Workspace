VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmMiscIssuanceAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7665
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
   ScaleHeight     =   6450
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin ComboList.Usercontrol1 cmbCountry 
      Height          =   405
      Left            =   5385
      TabIndex        =   2
      Top             =   525
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   714
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextBackColor   =   12640511
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin ComboList.Usercontrol1 cmbCustCode 
      Height          =   405
      Left            =   1710
      TabIndex        =   1
      Top             =   525
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   714
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextBackColor   =   12640511
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin ComboList.Usercontrol1 cmbEmp 
      Height          =   405
      Left            =   270
      TabIndex        =   3
      Top             =   1260
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   714
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextBackColor   =   12640511
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.TextBox txtPakRs 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   6225
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2730
      Width           =   1065
   End
   Begin VB.TextBox txtExchRate 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   5160
      TabIndex        =   9
      Top             =   2730
      Width           =   1065
   End
   Begin VB.TextBox txtCourier 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   4080
      TabIndex        =   8
      Top             =   2730
      Width           =   1065
   End
   Begin VB.TextBox txtAWBNo 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   2322
      TabIndex        =   7
      Top             =   2730
      Width           =   1740
   End
   Begin VB.TextBox txtWeight 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   1296
      TabIndex        =   6
      Top             =   2730
      Width           =   1020
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      Height          =   405
      Left            =   270
      TabIndex        =   5
      Top             =   2730
      Width           =   1020
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   405
      Left            =   270
      TabIndex        =   0
      Top             =   525
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   714
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   590872579
      CurrentDate     =   41041
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   465
      Left            =   4155
      TabIndex        =   11
      Top             =   5805
      Width           =   1635
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   465
      Left            =   5850
      TabIndex        =   12
      Top             =   5805
      Width           =   1635
   End
   Begin VB.TextBox txtRemarks 
      Height          =   405
      Left            =   270
      TabIndex        =   4
      Top             =   1965
      Width           =   7020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pak Rs."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   10
      Left            =   6240
      TabIndex        =   23
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Exch. Rate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   9
      Left            =   5160
      TabIndex        =   22
      Top             =   2520
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Net Courier"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   4080
      TabIndex        =   21
      Top             =   2520
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "AWB No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   2325
      TabIndex        =   20
      Top             =   2520
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Weight"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   1290
      TabIndex        =   19
      Top             =   2520
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Country"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   5400
      TabIndex        =   18
      Top             =   300
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   1725
      TabIndex        =   17
      Top             =   315
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Qty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   270
      TabIndex        =   16
      Top             =   2520
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   285
      TabIndex        =   15
      Top             =   1740
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sent By"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   285
      TabIndex        =   14
      Top             =   1035
      Width           =   645
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   285
      TabIndex        =   13
      Top             =   300
      Width           =   405
   End
End
Attribute VB_Name = "frmMiscIssuanceAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lEntryID As Long
Dim bReturnValue As Boolean

Private Sub cmbCustCode_matched()

    cmbCountry.ClearVals
    cmbCountry.AddVals con, "Country", "ForeignCustomers", "Country", " WHERE CustCode='" & cmbCustCode.id & "'"
        
    Dim strCountry As String
    strCountry = GetSingleStringValue("Country", "FOreignCustomers", " WHERE CustCode='" & cmbCustCode.id & "'")
    
    cmbCountry.id = strCountry
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    On Error GoTo err
    
    If cmbCustCode.Text = "" Then
        MsgBox "Please Enter Customer.", vbInformation
        Exit Sub
    End If
    
    If lEntryID = 0 Then
        con.Execute "INSERT INTO ItemIssuanceMisc(DT,UserName,MachineName,SentBy,SentTo,Remarks,Qty,Country,Weight,AWBNo,NetCourierCharges,ExchRate) VALUES('" & _
         DTPicker1.Value & "','" & CurrentUserName & "','" & strComputerName & "','" & cmbEmp.Text & "','" & _
         cmbCustCode.Text & "','" & txtRemarks & "'," & Val(txtQty) & ",'" & cmbCountry.Text & "'," & Val(txtWeight) & _
         ",'" & txtAWBNo & "'," & Val(txtCourier) & "," & Val(txtExchRate) & ")"
    Else
        con.Execute "UPDATE ItemIssuanceMisc SET DT='" & DTPicker1 & "',SentBy='" & cmbEmp.Text & _
         "',SentTo='" & cmbCustCode.Text & "',Remarks='" & txtRemarks & "',Qty=" & Val(txtQty) & ",Country='" & cmbCountry.Text & "',Weight=" & _
         Val(txtWeight) & ",AWBNo='" & txtAWBNo.Text & "',NetCourierCharges=" & Val(txtCourier) & ",ExchRate=" & Val(txtExchRate) & " WHERE EntryID=" & lEntryID
    End If
    
    bReturnValue = True
    
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    lEntryID = 0
    DTPicker1.Value = Date
    bReturnValue = False
    
    cmbEmp.AddVals con, "Name", "Employees", "EmpID", " WHERE Active=1"
    cmbCustCode.AddVals con, "DISTINCT CustCode", "ForeignCustomers", "CustCode"
    
End Sub

Public Function ShowMe(Optional p_lEntryID As Long) As Boolean

    lEntryID = p_lEntryID
    If lEntryID > 0 Then
        Dim rs As New ADODB.Recordset
        With rs
            .Open "SELECT * FROM ItemIssuanceMisc WHERE EntryID=" & lEntryID, con, adOpenForwardOnly, adLockReadOnly
            DTPicker1 = !DT & ""
            cmbEmp.Text = !SentBy & ""
            cmbCustCode.Text = !SentTo & ""
            txtRemarks = !Remarks & ""
            txtQty = Val(!Qty & "")
            
            cmbCountry.id = !Country & ""
            txtWeight = Val(!Weight & "")
            txtAWBNo = !AWBNo & ""
            txtCourier = Val(!NetCourierCharges & "")
            txtExchRate = Val(!ExchRate & "")
            
            .Close
        End With
    End If
    Me.Show 1
    ShowMe = bReturnValue
    
End Function

Private Sub txtCourier_Change()
    txtPakRs = Val(txtCourier) * Val(txtExchRate)
End Sub

Private Sub txtExchRate_Change()
    txtPakRs = Val(txtCourier) * Val(txtExchRate)
End Sub
