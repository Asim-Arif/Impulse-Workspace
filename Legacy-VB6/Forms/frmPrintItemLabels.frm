VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmPrintItemLabels 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Item Labels"
   ClientHeight    =   4335
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6690
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbCustomer 
      Height          =   285
      Left            =   240
      TabIndex        =   10
      Top             =   420
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   556
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextBackColor   =   12640511
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin ComboList.Usercontrol1 cmbItem 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   1065
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   556
      BeginProperty ListFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextBackColor   =   12640511
      BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ListIndex       =   -1
   End
   Begin VB.TextBox txtQty 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4230
      TabIndex        =   2
      Top             =   1740
      Width           =   1140
   End
   Begin VB.TextBox txtNo 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5385
      TabIndex        =   3
      Top             =   1740
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Customer"
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
      Left            =   240
      TabIndex        =   11
      Top             =   195
      Width           =   825
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
      Index           =   3
      Left            =   4260
      TabIndex        =   9
      Top             =   1515
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "No. of Labels"
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
      Left            =   5385
      TabIndex        =   8
      Top             =   1515
      Width           =   1050
   End
   Begin MSForms.ComboBox cmbOrderNo 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   1740
      Visible         =   0   'False
      Width           =   3975
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "7011;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   178
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Order No."
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
      Left            =   240
      TabIndex        =   7
      Top             =   1515
      Visible         =   0   'False
      Width           =   780
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
      Height          =   195
      Index           =   0
      Left            =   225
      TabIndex        =   6
      Top             =   840
      Width           =   420
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   4725
      TabIndex        =   5
      Top             =   3810
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Cancel       "
      PicturePosition =   327683
      Size            =   "3096;635"
      Picture         =   "frmPrintItemLabels.frx":0000
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
      Left            =   4725
      TabIndex        =   4
      Top             =   3360
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Print           "
      PicturePosition =   327683
      Size            =   "3096;635"
      Picture         =   "frmPrintItemLabels.frx":0112
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
Attribute VB_Name = "frmPrintItemLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbCust_matched()
    
    On Error GoTo err
    Dim rs As New ADODB.Recordset
    With rs
        .Open "Select Country,Curr From ForeignCustomers Where CustCode='" & cmbCust.ID & "'", con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            cmbCountry.List(cmbCountry.ListCount - 1, 1) = ![Curr] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    
    Set rs = Nothing
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


Private Sub cmbCustomInvoice_Change()
    Call cmbCustomInvoice_Click
End Sub

Private Sub cmbCustomInvoice_Click()
    On Error GoTo err
    If cmbCustomInvoice.MatchFound = False Then Exit Sub
    
    Dim rs As New ADODB.Recordset
    Dim TotalCartons As Long
    With rs
        .Open "Select TotalCartons From VTotalCartons Where CustomInvoice='" & cmbCustomInvoice & "'", con, adOpenForwardOnly, adLockReadOnly
        
        If .RecordCount = 0 Then
            TotalCartons = 0
        Else
            TotalCartons = Val(.Fields(0) & "")
        End If
        .Close
    End With
    Set rs = Nothing
    
    cmbFrom.Clear
    cmbTo.Clear
    If TotalCartons = 0 Then
        MsgBox "Packing List Not Generated For The Selected Invoice.", vbInformation
        Exit Sub
    Else
    
        For i = 1 To TotalCartons
            cmbFrom.AddItem i
            cmbTo.AddItem i
        Next
        
        cmbFrom.ListIndex = 0
        cmbTo.ListIndex = cmbTo.ListCount - 1
    End If
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmbCustomer_matched()

    cmbItem.ClearVals
    cmbItem.AddVals con, "'{'+Items.ItemID+'} ' + Items.ItemName + ' {'+FCustomerCatalog.ItemID+'}'", "Items INNER JOIN FCustomerCatalog ON Items.ItemID=FCustomerCatalog.CompItemID", "Items.ItemID", " WHERE CustCode='" & cmbCustomer.ID & "'"
        
End Sub

Private Sub cmbItem_matched()
    Call AddToCombo(cmbOrderNo, "DISTINCT OrderNo", "FOrderItems", " WHERE CompItemCode='" & cmbItem.ID & "'")
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub




Private Sub cmdPrint_Click()

    On Error GoTo err
    If cmbCustomer.MatchFound = False Then
        MsgBox "Please Select Customer.", vbInformation
        Exit Sub
    ElseIf cmbItem.MatchFound = False Then
        MsgBox "Please Select Item.", vbInformation
        Exit Sub
    End If
    
    If Val(txtNo) <= 0 Then
        MsgBox "Invalid No.", vbInformation
        Exit Sub
    End If
    
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ItemLabel_Customer.rpt")
    
    rpt.EnableParameterPrompting = False
    rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue cmbCustomer.ID
    rpt.ParameterFields.GetItemByName("@ItemID").SetCurrentValue cmbItem.ID
    rpt.ParameterFields.GetItemByName("@PONo").SetCurrentValue cmbOrderNo.Text
    rpt.ParameterFields.GetItemByName("@No").SetCurrentValue Val(txtNo.Text)
    rpt.ParameterFields.GetItemByName("@Qty").SetCurrentValue Val(txtQty)
    
    Dim FRM As New frmPrevRpt
    
    Me.Hide
    FRM.ShowReport "", rpt, "", False
    Unload Me
    
    Exit Sub
err:
    MsgBox err.Description
 End Sub

Private Sub cmdPrintItemLabel_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ItemLabel.rpt")
    
    rpt.EnableParameterPrompting = False
    rpt.ParameterFields.GetItemByName("@InvoiceNo").SetCurrentValue cmbCustomInvoice.Text
    Dim FRM As New frmPrevRpt
    
    Me.Hide
    FRM.ShowReport "", rpt, "", False
    Unload Me
    
End Sub

Private Sub Form_Load()
            
    cmbCustomer.ListHeight = 3000
    cmbCustomer.AddVals con, "CustCode", "ForeignCustomers", "CustCode"
    
    cmbItem.ListHeight = 3000
    cmbItem.DropDownOffLine = True
    
    
End Sub
Private Sub cmbCountry_Change()
    Call cmbCountry_Click
End Sub

Private Sub cmbCountry_Click()
    
    On Error GoTo err
    If cmbCountry.MatchFound = False Then Exit Sub
    cmbCountry.Tag = cmbCust.Text & cmbCountry
    txtCurrency = cmbCountry.List(cmbCountry.ListIndex, 1)
    
    Dim rs As New ADODB.Recordset
    
    cmbCustomInvoice.Clear
    With rs
        .Open "Select CustomInvoice From VrptCustomInvoice Where CustCode='" & cmbCust.ID & "' and Country='" & cmbCountry & "' ORDER BY CustomInvoice", con, adOpenForwardOnly, adLockReadOnly
        Do Until .EOF
            cmbCustomInvoice.AddItem .Fields(0) & ""
            
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub


