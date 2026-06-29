VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmFCustomerOrders_ShippingDates 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Shipping Dates"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Calibri"
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
   ScaleHeight     =   3825
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComboList.Usercontrol1 cmbCustomer 
      Height          =   285
      Left            =   210
      TabIndex        =   0
      Top             =   375
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   503
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin ComboList.Usercontrol1 cmbOrderNo 
      Height          =   285
      Left            =   1950
      TabIndex        =   1
      Top             =   375
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   503
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
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6135
      TabIndex        =   9
      Top             =   3225
      Width           =   1350
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   7560
      TabIndex        =   10
      Top             =   3225
      Width           =   1350
   End
   Begin MSComCtl2.DTPicker DTDelivery 
      Height          =   345
      Left            =   5430
      TabIndex        =   2
      Top             =   375
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   115605507
      CurrentDate     =   41010
   End
   Begin MSComCtl2.DTPicker DT1 
      Height          =   345
      Left            =   210
      TabIndex        =   4
      Top             =   1065
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   115605507
      CurrentDate     =   41010
   End
   Begin MSComCtl2.DTPicker DT2 
      Height          =   345
      Left            =   1950
      TabIndex        =   5
      Top             =   1065
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   101187587
      CurrentDate     =   41010
   End
   Begin MSComCtl2.DTPicker DT3 
      Height          =   345
      Left            =   3690
      TabIndex        =   6
      Top             =   1065
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   101187587
      CurrentDate     =   41010
   End
   Begin MSComCtl2.DTPicker DT4 
      Height          =   345
      Left            =   5430
      TabIndex        =   7
      Top             =   1065
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   101187587
      CurrentDate     =   41010
   End
   Begin MSComCtl2.DTPicker DT5 
      Height          =   345
      Left            =   7170
      TabIndex        =   8
      Top             =   1065
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   101187587
      CurrentDate     =   41010
   End
   Begin MSComCtl2.DTPicker DTCurrentDelivery 
      Height          =   345
      Left            =   7170
      TabIndex        =   3
      Top             =   375
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   101253123
      CurrentDate     =   41010
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Current Delivery"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   7170
      TabIndex        =   19
      Top             =   105
      Width           =   1680
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ship. Date 5"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   7170
      TabIndex        =   18
      Top             =   750
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ship. Date 4"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   5445
      TabIndex        =   17
      Top             =   765
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ship. Date 3"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   3690
      TabIndex        =   16
      Top             =   765
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ship. Date 2"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   1950
      TabIndex        =   15
      Top             =   765
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ship. Date 1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   210
      TabIndex        =   14
      Top             =   765
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Original Delivery"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   5430
      TabIndex        =   13
      Top             =   105
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Order No."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   12
      Top             =   105
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   165
      TabIndex        =   11
      Top             =   105
      Width           =   975
   End
End
Attribute VB_Name = "frmFCustomerOrders_ShippingDates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbCustomer_matched()
    cmbOrderNo.ClearVals
    cmbOrderNo.AddVals con, "InternalRefNo+' {'+OrderNo+'}'", "FCustomerOrders", "OrderNo", " WHERE CustCode='" & cmbCustomer.id & "'"
End Sub

Private Sub cmbOrderNo_matched()

    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM VFCustomerOrders_ShippingDates WHERE OrderNo='" & cmbOrderNo.id & "'", con, adOpenForwardOnly, adLockReadOnly
        
        DTCurrentDelivery = !DeliveryDT
        
        DT1.Value = !DT1 & ""
        DT2.Value = !DT2 & ""
        DT3.Value = !DT3 & ""
        DT4.Value = !DT4 & ""
        DT5.Value = !DT5 & ""
         
        If IsNull(!OriginalDelivery) Then
            DTDelivery = !DeliveryDT
        Else
            DTDelivery = !OriginalDelivery
        End If
        
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()

    Dim iRecAffect As Integer
    Dim strDT1 As String, strDT2 As String, strDT3 As String, strDT4 As String, strDT5 As String
    
    If IsNull(DT1.Value) Then
        strDT1 = "NULL"
    Else
        strDT1 = "'" & Format(DT1.Value, "dd-MMM-yyyy") & "'"
    End If
     
    If IsNull(DT2.Value) Then
        strDT2 = "NULL"
    Else
        strDT2 = "'" & Format(DT2.Value, "dd-MMM-yyyy") & "'"
    End If
    
    If IsNull(DT3.Value) Then
        strDT3 = "NULL"
    Else
        strDT3 = "'" & Format(DT3.Value, "dd-MMM-yyyy") & "'"
    End If
    
    If IsNull(DT4.Value) Then
        strDT4 = "NULL"
    Else
        strDT4 = "'" & Format(DT4.Value, "dd-MMM-yyyy") & "'"
    End If
    
    If IsNull(DT5.Value) Then
        strDT5 = "NULL"
    Else
        strDT5 = "'" & Format(DT5.Value, "dd-MMM-yyyy") & "'"
    End If
    
    Call StartTrans(con)
    con.Execute "UPDATE FCustomerOrders_ShippingDates SET OriginalDelivery='" & DTDelivery & "',DT1=" & strDT1 & ",DT2=" & strDT2 & ",DT3=" & strDT3 & ",DT4=" & strDT4 & ",DT5=" & strDT5 & " WHERE OrderNo='" & cmbOrderNo.id & "'", iRecAffect
    If iRecAffect = 0 Then
        con.Execute "INSERT INTO FCustomerOrders_ShippingDates(OrderNo,DT1,DT2,DT3,DT4,DT5,OriginalDelivery) VALUES('" & cmbOrderNo.id & "'," & strDT1 & "," & strDT2 & "," & strDT3 & "," & strDT4 & "," & strDT5 & ",'" & DTDelivery & "')"
    End If
    con.Execute "UPDATE FCustomerOrders SET DeliveryDT='" & DTCurrentDelivery & "' WHERE OrderNo='" & cmbOrderNo.id & "'"
    con.CommitTrans
    
    
End Sub

Private Sub Form_Load()

    cmbCustomer.ListHeight = 3000
    cmbCustomer.AddVals con, "DISTINCT CustCode", "ForeignCustomers", "CustCode"
    
    cmbOrderNo.ListHeight = 3000
    
    DT1.Value = Date
    DT1.Value = Null
    
    DT2.Value = Date
    DT2.Value = Null
    
    DT3.Value = Date
    DT3.Value = Null
    
    DT4.Value = Date
    DT4.Value = Null
    
    DT5.Value = Date
    DT5.Value = Null
    
End Sub


Public Sub ShowMe(strCustCode As String, strOrderNo As String)

    cmbCustomer.id = strCustCode
    cmbOrderNo.id = strOrderNo
    
    Me.Show 1
    
End Sub
