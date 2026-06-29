VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmCustomerFeedbackAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Feedback"
   ClientHeight    =   9360
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9165
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   9165
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog CD1 
      Left            =   1680
      Top             =   -150
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "PDF Files Only|*.pdf"
   End
   Begin VB.TextBox txtFilePath 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   315
      TabIndex        =   29
      Top             =   8265
      Width           =   8040
   End
   Begin VB.CommandButton cmdBrowseBarcode 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   8385
      TabIndex        =   28
      Top             =   8250
      Width           =   405
   End
   Begin ComboList.Usercontrol1 cmbItems 
      Height          =   285
      Left            =   2400
      TabIndex        =   10
      Top             =   5625
      Width           =   3675
      _ExtentX        =   6482
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
   Begin ComboList.Usercontrol1 cmbInvoiceNo 
      Height          =   285
      Left            =   315
      TabIndex        =   9
      Top             =   5625
      Width           =   2040
      _ExtentX        =   3598
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
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   300
      TabIndex        =   0
      Tag             =   "CustCode"
      Top             =   360
      Width           =   2310
      _ExtentX        =   4075
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
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   330
      Left            =   7080
      TabIndex        =   3
      Tag             =   "DT"
      Top             =   360
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   108068867
      CurrentDate     =   40698
   End
   Begin MSComctlLib.ListView LV 
      Height          =   1995
      Left            =   315
      TabIndex        =   22
      Top             =   5940
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   3519
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      ForeColor       =   8388608
      BackColor       =   16777215
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Invoice No."
         Object.Width           =   3572
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Product"
         Object.Width           =   6456
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Qty"
         Object.Width           =   2011
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ref."
         Object.Width           =   2752
      EndProperty
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   3
      Left            =   4575
      TabIndex        =   5
      Tag             =   "Designation"
      Top             =   945
      Width           =   4215
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "7435;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Designation"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   3
      Left            =   4575
      TabIndex        =   32
      Top             =   720
      Width           =   840
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   2
      Left            =   315
      TabIndex        =   4
      Tag             =   "ComplaintBy"
      Top             =   945
      Width           =   4215
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "7435;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Received by"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   2
      Left            =   345
      TabIndex        =   31
      Top             =   720
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "PDF Path to Attach"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   330
      TabIndex        =   30
      Top             =   7980
      Width           =   2040
   End
   Begin MSForms.TextBox TBox 
      Height          =   1065
      Index           =   1
      Left            =   315
      TabIndex        =   6
      Tag             =   "Description"
      Top             =   1575
      Width           =   8475
      VariousPropertyBits=   -1467987941
      BorderStyle     =   1
      Size            =   "14949;1879"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   1
      Left            =   345
      TabIndex        =   27
      Top             =   1335
      Width           =   795
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Reference"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   3
      Left            =   7260
      TabIndex        =   26
      Top             =   5400
      Width           =   750
   End
   Begin MSForms.TextBox txtEmailRef 
      Height          =   315
      Left            =   7230
      TabIndex        =   12
      Top             =   5610
      Width           =   1560
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "2752;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Qty"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   2
      Left            =   6060
      TabIndex        =   25
      Top             =   5400
      Width           =   270
   End
   Begin MSForms.TextBox txtQty 
      Height          =   315
      Left            =   6075
      TabIndex        =   11
      Top             =   5610
      Width           =   1140
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "2011;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Product"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   1
      Left            =   2385
      TabIndex        =   24
      Top             =   5400
      Width           =   555
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Invoice No."
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   345
      TabIndex        =   23
      Top             =   5400
      Width           =   825
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Action Requested"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   345
      TabIndex        =   21
      Top             =   4035
      Width           =   1275
   End
   Begin MSForms.TextBox TBox 
      Height          =   1065
      Index           =   0
      Left            =   315
      TabIndex        =   8
      Tag             =   "ActionRequested"
      Top             =   4275
      Width           =   8475
      VariousPropertyBits=   -1467987941
      BorderStyle     =   1
      Size            =   "14949;1879"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Country"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   7
      Left            =   2670
      TabIndex        =   20
      Top             =   150
      Width           =   585
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Customer"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   4
      Left            =   345
      TabIndex        =   19
      Top             =   150
      Width           =   690
   End
   Begin MSForms.ComboBox cmbCountry 
      Height          =   315
      Left            =   2655
      TabIndex        =   1
      Tag             =   "Country"
      Top             =   360
      Width           =   1845
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3254;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Reason for Feedback"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   19
      Left            =   345
      TabIndex        =   18
      Top             =   2655
      Width           =   1530
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Feedback No."
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   73
      Left            =   4515
      TabIndex        =   17
      Top             =   150
      Width           =   990
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   14
      Left            =   4530
      TabIndex        =   2
      Tag             =   "ComplaintNo"
      Top             =   360
      Width           =   2520
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "4445;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.TextBox TBox 
      Height          =   1065
      Index           =   15
      Left            =   315
      TabIndex        =   7
      Tag             =   "Reason"
      Top             =   2895
      Width           =   8475
      VariousPropertyBits=   -1467987941
      BorderStyle     =   1
      Size            =   "14949;1879"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Date"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   5
      Left            =   7125
      TabIndex        =   16
      Top             =   150
      Width           =   345
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   420
      Left            =   7215
      TabIndex        =   15
      Top             =   8805
      Width           =   1575
      Caption         =   "Cancel            "
      PicturePosition =   327683
      Size            =   "2778;741"
      Accelerator     =   67
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton cmdOk 
      Height          =   420
      Left            =   5580
      TabIndex        =   14
      Top             =   8805
      Width           =   1575
      Caption         =   " Save & Close"
      PicturePosition =   327683
      Size            =   "2778;741"
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
      Height          =   420
      Left            =   3945
      TabIndex        =   13
      Top             =   8805
      Visible         =   0   'False
      Width           =   1575
      Caption         =   "  Save & New"
      PicturePosition =   327683
      Size            =   "2778;741"
      Accelerator     =   78
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmCustomerFeedbackAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lEntryID As Long

Public Sub EditItem(p_lEntryID As Long)

    On Error GoTo err
    lEntryID = p_lEntryID
    
    Call ShowData(lEntryID)
    
    Me.Show 1
    
    Exit Sub
err:
    MsgBox err.Description
    
End Sub

Private Sub FillCmbs()

    On Error GoTo 0
    
    cmbCust.ListHeight = 2500
    Call cmbCust.AddVals(con, "DISTINCT CustCode", "ForeignCustomers", "CustCode")
    
    cmbInvoiceNo.ListHeight = 2500
    cmbItems.ListHeight = 2500
    
    
    
End Sub

Private Sub cmbCust_matched()

    On Error GoTo err
    
    Dim rs As New ADODB.Recordset
    Dim SQL As String
    
    SQL = "SELECT DISTINCT Country FROM ForeignCustomers "
    
    If cmbCust.ID <> "0" Then
        SQL = SQL & " WHERE CustCode='" & cmbCust.ID & "'"
    End If
    
    With rs
        .Open SQL, con, adOpenForwardOnly, adLockReadOnly
        cmbCountry.Clear
        Do Until .EOF
            cmbCountry.AddItem ![Country] & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    If cmbCountry.ListCount > 0 Then cmbCountry.ListIndex = 0
    Set rs = Nothing
    
    
    cmbInvoiceNo.ClearVals
    cmbInvoiceNo.AddItem "<All Items>", "0"
    cmbInvoiceNo.AddVals con, "CustomInvoice", "CustomInvoice", "CustomInvoice", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'"
    cmbInvoiceNo.ID = "0"
    
    
    Exit Sub
err:
    MsgBox err.Description

End Sub

Private Sub cmbInvoiceNo_matched()

    cmbItems.DropDownOffLine = True
    cmbItems.ClearVals
     
    If cmbInvoiceNo.ID = "0" Then
        If cmbCust.MatchFound = False Then
            cmbItems.AddVals con, "ItemID", "Items", "ItemID"
        Else
            cmbItems.AddVals con, "'{' + CompItemID + '} ' + ItemID", "VFCustomerCatalog", "CompItemID", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'"
        End If
    Else
        cmbItems.AddVals con, "ItemID + ' {'+CustItemID+'}'", "VCustomInvoiceItems", "ItemID", " WHERE CustomInvoice='" & cmbInvoiceNo.Text & "'"
    End If
    
End Sub

Private Sub cmbPenalty_Change()
    Call cmbPenalty_Click
End Sub

Private Sub cmbPenalty_Click()

    Tbox(4).Enabled = False
    Tbox(5).Enabled = False
    Tbox(6).Enabled = False
'    If cmbPenalty.ListIndex = 1 Then
'        TBox(4).Enabled = True
'    ElseIf cmbPenalty.ListIndex = 2 Then
'        TBox(5).Enabled = True
'        TBox(6).Enabled = True
'    End If
    
End Sub

Private Sub cmdBrowseBarcode_Click()

    On Error Resume Next
    Dim Picfile As String
    CD1.ShowSave
    
    If Not (err.Number = 32755) Then
        Picfile = CD1.FileName
        txtFilePath = Picfile
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


Private Sub cmdNext_Click()
    If SaveItem Then
        Unload Me
        
        Me.Show 1
    End If
End Sub

Private Sub cmdOK_Click()
    If SaveItem Then Unload Me
End Sub

Private Sub DTPicker2_Change()
    Tbox(14) = GetNewComplaintNo(DTPicker2.Value)
End Sub

Private Sub Form_Load()
    
    Me.KeyPreview = True
    DTPicker2.Value = Date
    Call DTPicker2_Change
    
    Call FillCmbs
    
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err
    
    
    
    Dim SQL As String
    
    Call StartTrans(con)
    
    Dim lmyEntryID As Long
    If lEntryID = 0 Then
        SQL = CreateInsertSQL("CustomerComplaints")
    Else
        SQL = CreateUpdateSQL("CustomerComplaints")
        SQL = SQL & " WHERE EntryID=" & lEntryID
    End If
    
    con.Execute SQL
    
    
    If lEntryID = 0 Then
        lmyEntryID = GetSingleLongValue("MAX(EntryID)", "CustomerComplaints")
    Else
        lmyEntryID = lEntryID
    End If
    
    con.Execute "UPDATE CustomerComplaints SET Type=1 WHERE EntryID=" & lmyEntryID
    con.Execute "DELETE FROM CustomerComplaints_Detail WHERE RefID=" & lmyEntryID
    
    Dim i As Integer
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
            con.Execute "INSERT INTO CustomerComplaints_Detail(RefID,InvoiceNo,ItemID,Qty,EmailReference) VALUES(" & lmyEntryID & _
             ",'" & .Text & "','" & .ListSubItems(1).Tag & "'," & Val(.SubItems(2)) & ",'" & .SubItems(3) & "')"
        End With
    Next
    con.CommitTrans
     
    
    Call AttachPDF(lmyEntryID)
   
    
SavePic:

'    If (Image1.Picture = 0 Or Trim(Image1.Tag) = "") Then GoTo TheEnd
'
'    Dim Rec As New ADODB.Recordset
'    Dim rsPic As New ADODB.Stream
'    Dim rsThumb As New ADODB.Stream
'
'    Rec.Open "SELECT EntryID,AssetPic FROM FixedAssets WHERE EntryID=" & lmyEntryID, con, adOpenStatic, adLockOptimistic
'    If Image1.Picture <> 0 And Trim(Image1.Tag) <> "" Then
'        rsPic.Open
'        rsPic.Type = adTypeBinary
'        rsPic.LoadFromFile Image1.Tag
'        Rec.Fields(1) = rsPic.Read
'        rsPic.Close
'        Rec.Update
'    End If

TheEnd:
    lEntryID = 0
    SaveItem = True
    
    
    Exit Function
err:
    MsgBox err.Description
    
End Function



Private Function CreateInsertSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "INSERT INTO " & strTable & "("
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is MSForms.CheckBox) Or (TypeOf c Is ComboList.Usercontrol1) Or (TypeOf c Is DTPicker) Then
            If c.Tag <> "" Then
                RetString = RetString & c.Tag & ","
            End If
        End If
    Next
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ") Values("
    
    
    'Now  Get The Second Part...
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        
        If c.Tag <> "" Then
            If (TypeOf c Is MSForms.ComboBox) Then
                If c.Name = "cmbCountry" Then
                    RetString = RetString & "'" & c.Text & "" & "',"
                ElseIf c.Name = "cmbPenalty" Then
                    RetString = RetString & c.ListIndex & ","
                Else
                    RetString = RetString & "'" & c.List(c.ListIndex, 1) & "',"
                End If
            ElseIf (TypeOf c Is MSForms.TextBox) Then
                If c.Name = "NBox" Then
                    RetString = RetString & Val(c.Text & "") & ","
                Else
                    RetString = RetString & "'" & c.Text & "',"
                End If
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                RetString = RetString & Abs(c.Value) & ","
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                RetString = RetString & "'" & c.ID & "',"
            ElseIf (TypeOf c Is DTPicker) Then
                RetString = RetString & "'" & c.Value & "',"
            End If
        End If
    Next
    
    '---------------------------------------------------------------------------
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    RetString = RetString & ")"
    CreateInsertSQL = RetString
    
End Function
Private Function CreateUpdateSQL(strTable As String) As String
    
    Dim c As Control
    Dim RetString As String
    
    'Create The First Part Of Insert Query...
    RetString = "Update " & strTable & " Set "
    '---------------------------------------------------------------------------
    For Each c In Controls
        
        If c.Tag <> "" Then
            If (TypeOf c Is MSForms.ComboBox) Then
                If c.Name = "cmbCountry" Then
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
                ElseIf c.Name = "cmbPenalty" Then
                    RetString = RetString & c.Tag & "=" & c.ListIndex & ","
                Else
                    RetString = RetString & c.Tag & "='" & c.List(c.ListIndex, 1) & "',"
                End If
            ElseIf (TypeOf c Is MSForms.TextBox) Then   '
                If c.Name = "cmbCatID" Or c.Name = "cmbItemType" Then
                    RetString = RetString & c.Tag & "=" & Val(c.List(c.ListIndex, 1)) & ","
                ElseIf c.Name = "NBox" Then
                    RetString = RetString & c.Tag & "=" & Val(c.Text) & ","
                ElseIf c.Name = "cmbSteel" Or c.Name = "cmbItemGroup" Then
                    RetString = RetString & c.Tag & "=" & c.List(c.ListIndex, 1) & ","
                ElseIf c.Name = "cmbGuage" Then
                    RetString = RetString & c.Tag & "='" & c.List(c.ListIndex, 1) & "',"
                Else
                    RetString = RetString & c.Tag & "='" & c.Text & "',"
                End If
            ElseIf (TypeOf c Is MSForms.CheckBox) Then
                RetString = RetString & c.Tag & "=" & Abs(c.Value) & ","
            ElseIf (TypeOf c Is ComboList.Usercontrol1) Then
                RetString = RetString & c.Tag & "='" & c.ID & "',"
            ElseIf (TypeOf c Is DTPicker) Then
                RetString = RetString & c.Tag & "='" & c.Value & "',"
            End If
        End If
    Next
    
    
    
    RetString = Left$(RetString, Len(RetString) - 1) 'Get Rid Of The Last ","
    
    CreateUpdateSQL = RetString
    
End Function



Private Sub ShowData(p_lEntryID As Long, Optional bCopy As Boolean = False)

    On Error GoTo err
    Dim rs As New ADODB.Recordset
    Dim c As Control
    With rs
        .Open "SELECT * FROM CustomerComplaints WHERE EntryID=" & p_lEntryID, con, adOpenForwardOnly, adLockReadOnly
        
        For Each c In Controls
            If c.Tag <> "" Then
                If TypeOf c Is MSForms.CheckBox Then
                    c.Value = .Fields(c.Tag)
                ElseIf TypeOf c Is ComboList.Usercontrol1 Then
                    'C.Text = .Fields(C.Tag & "Text") & ""
                    c.ID = .Fields(c.Tag) & ""
                ElseIf TypeOf c Is DTPicker Then
                    If Not IsNull(.Fields(c.Tag)) Then
                        c.Value = .Fields(c.Tag)
                    End If
                ElseIf c.Name = "cmbPenalty" Then
                    c.ListIndex = Val(.Fields(c.Tag) & "")
                Else
                    c.Text = .Fields(c.Tag) & ""
                End If
            End If
        Next
        
'        If Not IsNull(![AssetPic]) Then
'            Set Image1.DataSource = rs
'            Image1.DataField = "AssetPic"
'            lblPic.Visible = False
'        Else
'            lblPic.Visible = True
'        End If
        
        .Close
        Dim ITM As ListItem
        .Open "SELECT * FROM VCustomerComplaints_Detail WHERE RefID=" & lEntryID & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        Do Until .EOF
            Set ITM = LV.ListItems.add(, , !InvoiceNo & "")
            ITM.Tag = !EntryID
            ITM.ListSubItems.add(, , !ITemID & " {" & !ItemCode & "}").Tag = !ITemID & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , !EmailReference & ""
            .MoveNEXT
        Loop
        .Close
    End With
    
    Set rs = Nothing
    
    Exit Sub
err:
    MsgBox err.Description
End Sub


Public Sub AddNew()
    
    Tbox(14) = GetNewComplaintNo(DTPicker2.Value)
    Me.Show 1
    
End Sub

Private Function GetNewComplaintNo(myDT As Date) As String
    'CC-1/2011
    Dim lComplaintNo As Long
    lComplaintNo = GetSingleLongValue("MAX(CAST(SUBSTRING(ComplaintNo,4,CHARINDEX('/',ComplaintNo)-4) AS INT))", "CustomerComplaints", " WHERE Type=1")
    lComplaintNo = lComplaintNo + 1
    
    GetNewComplaintNo = "FB-" & lComplaintNo & "/" & year(myDT)
    
End Function

Private Sub Image1_DblClick()
    Call lblPic_DblClick
End Sub

Private Sub lblPic_DblClick()

    On Error Resume Next
    Dim Picfile As String
'    CD1.ShowSave
'
'    If Not (err.Number = 32755) Then
'        Picfile = CD1.FileName
'        Image1.Picture = LoadPicture(Picfile)
'        Image1.Tag = Picfile
'        lblPic.Visible = False
'    End If
    
End Sub

Private Sub txtEmailRef_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        Dim ITM As ListItem
        Set ITM = LV.ListItems.add(, , cmbInvoiceNo.Text)
        ITM.ListSubItems.add(, , cmbItems.Text).Tag = cmbItems.ID
        ITM.ListSubItems.add , , Val(txtQty)
        ITM.ListSubItems.add , , txtEmailRef
    End If
End Sub


Private Sub AttachPDF(p_lEntryID As Long)
    
    On Error Resume Next
    Dim strPDFPath As String
    strPDFPath = txtFilePath
    If strPDFPath <> "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EntryID,AttachmentPDF FROM CustomerComplaints WHERE EntryID=" & p_lEntryID, con, adOpenStatic, adLockOptimistic
        If Rec.RecordCount > 0 Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile strPDFPath
            Rec.Fields(1) = rsPic.Read
            rsPic.Close
            Rec.Update
        End If
        
        Rec.Close
        
        Set Rec = Nothing
        Set rsPic = Nothing
        
    End If

End Sub

