VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmPushDataTomySQL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Push Data to Web"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9300
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
   ScaleHeight     =   8475
   ScaleWidth      =   9300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUpdateDeliveryDates 
      Caption         =   "Update Delivery Dates Only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   315
      TabIndex        =   24
      Top             =   2385
      Width           =   4170
   End
   Begin VB.TextBox txtTotalRecords 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   315
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   6585
      Width           =   2820
   End
   Begin VB.TextBox txtSuccess 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3180
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   6585
      Width           =   2820
   End
   Begin VB.TextBox txtLeft 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6045
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   6585
      Width           =   2820
   End
   Begin ComboList.Usercontrol1 cmbOrderCustomer 
      Height          =   285
      Left            =   360
      TabIndex        =   7
      Top             =   4440
      Width           =   2040
      _ExtentX        =   3598
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
      Left            =   2430
      TabIndex        =   9
      Top             =   4455
      Width           =   2040
      _ExtentX        =   3598
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
   Begin ComboList.Usercontrol1 cmbInvoiceCustomer 
      Height          =   285
      Left            =   4740
      TabIndex        =   12
      Top             =   4455
      Width           =   2040
      _ExtentX        =   3598
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
   Begin ComboList.Usercontrol1 cmbInvoiceNo 
      Height          =   285
      Left            =   6810
      TabIndex        =   14
      Top             =   4455
      Width           =   2040
      _ExtentX        =   3598
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
   Begin VB.CommandButton cmdPushPackingList 
      Caption         =   "Push Packing List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4710
      TabIndex        =   15
      Top             =   2385
      Width           =   4170
   End
   Begin VB.CommandButton cmdUpdateInvoiceOrders 
      Caption         =   "Update Order against Invoice"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4710
      TabIndex        =   10
      Top             =   3375
      Width           =   4170
   End
   Begin VB.CommandButton cmdPushSF 
      Caption         =   "Push Semi Finish Stock"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4710
      TabIndex        =   5
      Top             =   1380
      Width           =   4170
   End
   Begin VB.CommandButton cmdPustFS 
      Caption         =   "Push Finished Stock"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4710
      TabIndex        =   4
      Top             =   315
      Width           =   4170
   End
   Begin VB.CommandButton cmdDeleteCustomerData 
      BackColor       =   &H000000FF&
      Caption         =   "Delete All Customer's Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   315
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1380
      Width           =   4170
   End
   Begin VB.CommandButton cmdUpdateOrder 
      Caption         =   "Update Order"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   315
      TabIndex        =   2
      Top             =   3375
      Width           =   4170
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4710
      TabIndex        =   1
      Top             =   7605
      Width           =   4170
   End
   Begin VB.CommandButton cmdPushCustomerOrderData 
      Caption         =   "Push Customer Orders Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   315
      TabIndex        =   0
      Top             =   315
      Width           =   4170
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   435
      Left            =   315
      TabIndex        =   16
      Top             =   5490
      Width           =   8550
      _ExtentX        =   15081
      _ExtentY        =   767
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Total Records"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   855
      TabIndex        =   23
      Top             =   6255
      Width           =   1620
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Success"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   4080
      TabIndex        =   22
      Top             =   6255
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Left"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   7140
      TabIndex        =   21
      Top             =   6255
      Width           =   465
   End
   Begin VB.Label lblImportStatus 
      Alignment       =   2  'Center
      Caption         =   "Uploading"
      Height          =   225
      Left            =   315
      TabIndex        =   17
      Top             =   5970
      Width           =   8550
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Invoice No."
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
      Left            =   6810
      TabIndex        =   13
      Top             =   4230
      Width           =   930
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
      Index           =   2
      Left            =   4725
      TabIndex        =   11
      Top             =   4230
      Width           =   825
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
      Left            =   2430
      TabIndex        =   8
      Top             =   4215
      Width           =   780
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
      Index           =   0
      Left            =   345
      TabIndex        =   6
      Top             =   4215
      Width           =   825
   End
End
Attribute VB_Name = "frmPushDataTomySQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbInvoiceCustomer_matched()

    cmbInvoiceNo.ClearVals
    
    cmbInvoiceNo.AddVals con, "CustomInvoice", "CustomInvoice", "CustomInvoice", " WHERE CustCode='" & cmbInvoiceCustomer.ID & "'"
    
End Sub

Private Sub cmbOrderCustomer_matched()

    cmbOrderNo.ClearVals
    
    cmbOrderNo.AddVals con, "OrderNo", "FCustomerOrders", "OrderNo", " WHERE CustCode='" & cmbOrderCustomer.ID & "'"
    
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdDeleteCustomerData_Click()
    
    If getDBPassword("DeleteCustomerDataPushtoWeb") = False Then
        Exit Sub
    End If
    If MsgBox("Are you sure?", vbQuestion + vbYesNo) = vbYes Then
        Dim mySQLCon As New ADODB.Connection
        mySQLCon.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=204.93.174.24;Port=3306;Database=frigzint_frigz;User=frigzint_frigz; Password=r3h@n;Option=3;"
        mySQLCon.Open
        
        mySQLCon.Execute "DELETE FROM tbl_CustomerOrders"
        mySQLCon.Execute "DELETE FROM tbl_customerorder_shippeddetail"
        
        con.Execute "UPDATE PushToWebStatus SET CustomerOrders=0,ShippedDetails=0,DTEntry=getDate(),UserName='" & CurrentUserName & "',MachineName='" & strComputerName & "'"
        MsgBox "Done."
    End If
     
End Sub

Private Sub cmdPushCustomerOrderData_Click()
    Call PushCustomerData
End Sub

Private Sub cmdPushPackingList_Click()

    Dim mySQLCon As New ADODB.Connection
    mySQLCon.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=204.93.174.24;Port=3306;Database=frigzint_frigz;User=frigzint_frigz; Password=r3h@n;Option=3;"
    mySQLCon.Open
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM VCustompList", con, adOpenForwardOnly, adLockReadOnly
        mySQLCon.Execute "DELETE FROM tbl_packinglist"
        
        If .EOF = False Then
            Call UpdateProgressBar(.RecordCount)
        End If
        Do Until .EOF
            mySQLCon.Execute "INSERT INTO tbl_packinglist(invoiceno,pinvoiceno,orderno,itemid,itemname,cartonfrom,cartonto,qtypercarton) VALUES('" & !CustomInvoice & _
             "','" & !PInvoice & "','" & !OrderNo & "','" & !ITemID & "','" & !ItemName & "'," & Val(!CartonFrom & "") & "," & Val(!CartonTo & "") & _
             "," & Val(!Qty & "") & ")"
            
            Call UpdateProgressBar(, .AbsolutePosition)
            .MoveNEXT
        Loop
        .Close
    End With

    mySQLCon.Close
    Set mySQLCon = Nothing
   
    MsgBox "Done."

End Sub

Private Sub cmdPushSF_Click()
            
    Dim mySQLCon As New ADODB.Connection
    mySQLCon.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=204.93.174.24;Port=3306;Database=frigzint_frigz;User=frigzint_frigz; Password=r3h@n;Option=3;"
    mySQLCon.Open
     
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT Items.ItemID,Items.ItemName,VStockOrderOpening.Description,SUM((Qty-IssdQty)) AS TotalStock FROM VStockOrderOpening INNER JOIN (SELECT Opening_RefID FROM VendRcvdDetail GROUP BY Opening_RefID) T1 ON T1.Opening_RefID=VStockOrderOpening.EntryID INNER JOIN Items ON VStockOrderOpening.ItemID=Items.ItemID WHERE Qty>IssdQty GROUP BY Items.ItemID,Items.ItemName,VStockOrderOpening.Description", con, adOpenForwardOnly, adLockReadOnly
        mySQLCon.Execute "DELETE FROM tbl_semifinishedstock"
        
        If .EOF = False Then
            Call UpdateProgressBar(.RecordCount)
        End If
        
        Do Until .EOF
            mySQLCon.Execute "INSERT INTO tbl_semifinishedstock(itemid,itemname,process,stock) VALUES('" & !ITemID & "','" & !ItemName & _
             "','" & !Description & "'," & Val(!TotalStock & "") & ")"
                 
            Call UpdateProgressBar(, .AbsolutePosition)
            
            .MoveNEXT
        Loop
        .Close
    End With

    mySQLCon.Close
    Set mySQLCon = Nothing
   
    MsgBox "Done."

End Sub

Private Sub cmdPustFS_Click()
        
    Dim mySQLCon As New ADODB.Connection
    mySQLCon.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=204.93.174.24;Port=3306;Database=frigzint_frigz;User=frigzint_frigz; Password=r3h@n;Option=3;"
    mySQLCon.Open
  
    Dim lLastCustomerOrders As Long, lLastShippedDetails As Long
   
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT Items.ItemID,Items.ItemName,SUM(NetQty) AS TotalStock FROM VItemsWithShelfWiseStock INNER JOIN Items ON VItemsWithShelfWiseStock.ItemID=Items.ItemID WHERE NetQty>0 GROUP BY Items.ItemID,Items.ItemName", con, adOpenForwardOnly, adLockReadOnly
        mySQLCon.Execute "DELETE FROM tbl_finishedstock"
        
        If .EOF = False Then
            Call UpdateProgressBar(.RecordCount)
        End If
        
        Do Until .EOF
            mySQLCon.Execute "INSERT INTO tbl_finishedstock(itemid,itemname,stock) VALUES('" & !ITemID & "','" & !ItemName & _
             "'," & Val(!TotalStock & "") & ")"
                  
            Call UpdateProgressBar(, .AbsolutePosition)
            
            .MoveNEXT
        Loop
        .Close
    End With

    mySQLCon.Close
    Set mySQLCon = Nothing
   
    MsgBox "Done."
    
End Sub

Private Sub cmdUpdateDeliveryDates_Click()
            
    Dim mySQLCon As New ADODB.Connection
    mySQLCon.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=204.93.174.24;Port=3306;Database=frigzint_frigz;User=frigzint_frigz; Password=r3h@n;Option=3;"
    mySQLCon.Open
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT FOrderItems_DeliveryDates_PTW.*,FOrderItems.DeliveryDT FROM FOrderItems_DeliveryDates_PTW INNER JOIN FOrderItems ON FOrderItems.ID=FOrderItems_DeliveryDates_PTW.FOI_RefID WHERE Updated=0 ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        
        If .EOF = False Then
            Call UpdateProgressBar(.RecordCount)
        End If
        
        Do Until .EOF
        
            mySQLCon.Execute "UPDATE tbl_CustomerOrders SET DeliveryDT='" & Format(!DeliveryDT, "yyyy-MM-dd") & "' WHERE EntryID=" & Val(!FOI_RefID & "")
           con.Execute "UPDATE FOrderItems_DeliveryDates_PTW SET Updated=1 WHERE EntryID=" & Val(!EntryID & "")
           
            Call UpdateProgressBar(, .AbsolutePosition)
            
            .MoveNEXT
        Loop
        .Close
    End With

    mySQLCon.Close
    Set mySQLCon = Nothing
   
    MsgBox "Done."

End Sub

Private Sub cmdUpdateInvoiceOrders_Click()

    If cmbInvoiceNo.MatchFound = False Then
        MsgBox "Please Select Invoice.", vbInformation
        Exit Sub
    End If
    
    Call PushCustomerData("", cmbInvoiceNo.ID)
    
End Sub

Private Sub cmdUpdateOrder_Click()

    If cmbOrderNo.MatchFound = False Then
        MsgBox "Please Select Order No.", vbInformation
        Exit Sub
    End If
    
    Call PushCustomerData(cmbOrderNo.ID)
    
End Sub


Private Sub PushCustomerData(Optional strOrderNo As String = "", Optional strInvoiceNo As String = "")

    Dim mySQLCon As New ADODB.Connection
    mySQLCon.ConnectionString = "Driver={MySQL ODBC 5.1 Driver};Server=204.93.174.24;Port=3306;Database=frigzint_frigz;User=frigzint_frigz; Password=r3h@n;Option=3;"
    mySQLCon.Open
    
    Dim lLastCustomerOrders As Long, lLastShippedDetails As Long, strOrderIDs As String
    
    Dim rs As New ADODB.Recordset
    With rs
        '.Open "SELECT * FROM VFCustomerOrders_PTW WHERE OrderNo NOT IN(SELECT OrderNo FROM VUnshippedOrderList) ORDER BY CustCode,OrderNo,ID", con, adOpenForwardOnly, adLockReadOnly
        If strOrderNo = "" And strInvoiceNo = "" Then
            .Open "SELECT MAX(CustomerOrders) AS Customerorders,MAX(ShippedDetails) AS ShippedDetails FROM PushToWebStatus", con, adOpenForwardOnly, adLockReadOnly
            If .EOF Then
                lLastCustomerOrders = 0
                lLastShippedDetails = 0
            Else
                lLastCustomerOrders = Val(!CustomerOrders & "")
                lLastShippedDetails = Val(!ShippedDetails & "")
            End If
            .Close
        End If
        Dim strOrderList As String
        If strOrderNo = "" And strInvoiceNo = "" Then
            .Open "SELECT * FROM VFCustomerOrders_PTW WHERE ID>" & lLastCustomerOrders & " ORDER BY ID", con, adOpenForwardOnly, adLockReadOnly
        ElseIf strOrderNo <> "" Then
            strOrderList = "'" & strOrderNo & "'"
            .Open "SELECT * FROM VFCustomerOrders_PTW WHERE OrderNo IN(" & strOrderList & ") ORDER BY ID", con, adOpenForwardOnly, adLockReadOnly
            
            mySQLCon.Execute "DELETE FROM tbl_customerorder_shippeddetail WHERE refid in(SELECT EntryID FROM tbl_CustomerOrders WHERE OrderNo IN(" & strOrderList & "))"
            mySQLCon.Execute "DELETE FROM tbl_CustomerOrders WHERE OrderNo IN(" & strOrderList & ")"
        ElseIf strInvoiceNo <> "" Then
            Dim strOrderEntryIDs As String
            strOrderEntryIDs = ""
            
            strOrderList = ""
            .Open "SELECT DISTINCT OrderNo FROM VCustomInvoice WHERE CustomInvoice='" & strInvoiceNo & "'", con, adOpenForwardOnly, adLockReadOnly
            Do Until .EOF
                strOrderList = strOrderList & "'" & !OrderNo & "',"
                
                .MoveNEXT
            Loop
            .Close
            
            If strOrderList <> "" Then
                strOrderList = Left(strOrderList, Len(strOrderList) - 1)
            End If
            
            .Open "SELECT ID FROM FOrderItems WHERE OrderNo IN(" & strOrderList & ")"
            strOrderIDs = ""
            Do Until .EOF
                strOrderIDs = strOrderIDs & Val(!ID & "") & ","
                .MoveNEXT
            Loop
            .Close
            If strOrderIDs <> "" Then
                strOrderIDs = Left(strOrderIDs, Len(strOrderIDs) - 1)
            End If
            
'            .Open "SELECT OrderEntryID FROM VCustomInvoice WHERE CustomInvoice='" & strInvoiceNo & "'", con, adOpenForwardOnly, adLockReadOnly
'            Do Until .EOF
'                strOrderEntryIDs = strOrderEntryIDs & !OrderEntryID & ","
'                .MoveNEXT
'            Loop
'            .Close
'            If strOrderEntryIDs <> "" Then
'                strOrderEntryIDs = Left(strOrderEntryIDs, Len(strOrderEntryIDs) - 1)
'            End If
            
            
            .Open "SELECT * FROM VFCustomerOrders_PTW WHERE OrderNo IN(" & strOrderList & ") ORDER BY ID", con, adOpenForwardOnly, adLockReadOnly
                        
            mySQLCon.Execute "DELETE FROM tbl_customerorder_shippeddetail WHERE refid in(SELECT EntryID FROM tbl_CustomerOrders WHERE OrderNo IN(" & strOrderList & "))"
            
            mySQLCon.Execute "DELETE FROM tbl_CustomerOrders WHERE OrderNo IN(" & strOrderList & ")"
            mySQLCon.Execute "DELETE FROM tbl_CustomerOrders WHERE EntryID IN(" & strOrderIDs & ")"
'            mySQLCon.Execute "DELETE FROM tbl_CustomerOrders WHERE OrderEntryID IN(" & strOrderEntryIDs & ")"
        End If
        'mySQLCon.Execute "DELETE FROM tbl_CustomerOrders"
        If .EOF = False Then
            Call UpdateProgressBar(.RecordCount)
        End If
        
        Do Until .EOF
             
            mySQLCon.Execute "INSERT INTO tbl_CustomerOrders(EntryID,OrderNo,CustCode,CompItemCode,ItemName" & _
             ",CustItemCode,DeliveryDT,Qty,Price,ShippedQty,order_date) VALUES(" & Val(!ID & "") & ",'" & !OrderNo & "','" & !CustCode & _
             "','" & !CompItemCode & "','" & !ItemName & "','" & !ItemCode & "','" & Format(!DeliveryDT, "yyyy-MM-dd") & "'," & _
             Val(!Qty & "") & "," & Val(!Price & "") & "," & Val(!ShippedQty & "") & ",'" & Format(!DT, "yyyy-MM-dd") & "')"
            
            If strOrderNo = "" And strInvoiceNo = "" Then
                con.Execute "UPDATE PushToWebStatus SET CustomerOrders=" & Val(!ID & "")
            End If
            Call UpdateProgressBar(, .AbsolutePosition)
            .MoveNEXT
        Loop
        .Close
        
        If strOrderNo = "" And strInvoiceNo = "" Then
            .Open "SELECT CustomInvoice.CustomInvoice,CustomInvoice.DT,CustomInvoiceItems.Qty,FOrderItems.ID FROM CustomInvoice INNER JOIN CustomInvoiceItems " & _
             " ON CustomInvoice.CustomInvoice=CustomInvoiceItems.CustomInvoice INNER JOIN FProformaOrders ON CustomInvoiceItems.RefID=FProformaOrders.EntryID " & _
             " INNER JOIN FOrderItems ON FProformaOrders.OrderEntryID=FOrderItems.ID WHERE FOrderItems.ID>" & lLastShippedDetails & " ORDER BY FOrderITems.ID", con, adOpenForwardOnly, adLockReadOnly
        Else
            .Open "SELECT CustomInvoice.CustomInvoice,CustomInvoice.DT,CustomInvoiceItems.Qty,FOrderItems.ID FROM CustomInvoice INNER JOIN CustomInvoiceItems " & _
             " ON CustomInvoice.CustomInvoice=CustomInvoiceItems.CustomInvoice INNER JOIN FProformaOrders ON CustomInvoiceItems.RefID=FProformaOrders.EntryID " & _
             " INNER JOIN FOrderItems ON FProformaOrders.OrderEntryID=FOrderItems.ID WHERE FOrderItems.OrderNo IN(" & strOrderList & ") ORDER BY FOrderITems.ID", con, adOpenForwardOnly, adLockReadOnly
             
            'mySQLCon.Execute "DELETE FROM tbl_customerorder_shippeddetail WHERE refid in(SELECT EntryID FROM tbl_CustomerOrders WHERE OrderNo IN(" & strOrderList & "))"
        End If
        
        If .EOF = False Then
            Call UpdateProgressBar(.RecordCount)
        End If
        Do Until .EOF
            mySQLCon.Execute "INSERT INTO tbl_customerorder_shippeddetail(refid,invoicenumber,dt,quantity) VALUES(" & Val(!ID & "") & _
             ",'" & !CustomInvoice & "','" & Format(!DT, "yyyy-MM-dd") & "'," & Val(!Qty & "") & ")"
             
            If strOrderNo = "" And strInvoiceNo <> "" Then
                con.Execute "UPDATE PushToWebStatus SET ShippedDetails=" & Val(!ID & "")
            End If
            Call UpdateProgressBar(, .AbsolutePosition)
            .MoveNEXT
        Loop
            
        .Close
    End With

    mySQLCon.Close
    Set mySQLCon = Nothing
    
    MsgBox "Done."
    
End Sub

Private Sub Form_Load()

    cmbOrderCustomer.ClearVals
    cmbOrderCustomer.ListHeight = 2400
    cmbOrderCustomer.AddVals con, "DISTINCT CustCode", "ForeignCustomers", "CustCode", " WHERE CustCode NOT IN('Stock')"
    
    cmbInvoiceCustomer.ClearVals
    cmbInvoiceCustomer.ListHeight = 2400
    cmbInvoiceCustomer.AddVals con, "DISTINCT CustCode", "ForeignCustomers", "CustCode", " WHERE CustCode NOT IN('Stock')"
    
    cmbOrderNo.ListHeight = 2400
    cmbInvoiceNo.ListHeight = 2400
End Sub

Private Sub UpdateProgressBar(Optional lTotalRecords As Long = 0, Optional lTotalDone As Long = 0)
    
    If lTotalRecords > 0 Then
        txtTotalRecords = lTotalRecords
        PB1.Min = 1
        PB1.Max = lTotalRecords + 1
    End If
         
    Dim i As Long, lTotalSuccess As Long, lTotalFail As Long, iStatus As Integer
    
    If lTotalDone > 0 Then
        lblImportStatus.Caption = "Uploading " & (lTotalDone + 1) & " of " & Val(txtTotalRecords)
        PB1.Value = lTotalDone
        txtSuccess = lTotalDone
        txtLeft = Val(txtTotalRecords) - lTotalDone
        Me.Refresh
        DoEvents
    End If
    
    If Val(txtTotalRecords) = lTotalDone Then
        lblImportStatus.Caption = "Done"
    End If
    
End Sub
