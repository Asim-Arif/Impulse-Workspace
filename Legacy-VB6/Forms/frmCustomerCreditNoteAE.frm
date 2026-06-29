VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmCustomerCreditNoteAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Complaints"
   ClientHeight    =   10590
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10560
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
   ScaleHeight     =   10590
   ScaleWidth      =   10560
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtPrice 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   4815
      TabIndex        =   41
      Top             =   3015
      Visible         =   0   'False
      Width           =   1275
   End
   Begin MSComctlLib.ListView LV 
      Height          =   7905
      Left            =   315
      TabIndex        =   24
      Top             =   1335
      Width           =   10110
      _ExtentX        =   17833
      _ExtentY        =   13944
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
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Price"
         Object.Width           =   2117
      EndProperty
   End
   Begin ComboList.Usercontrol1 cmbCompany 
      Height          =   285
      Left            =   315
      TabIndex        =   38
      Tag             =   "Company_RefID"
      Top             =   315
      Visible         =   0   'False
      Width           =   10110
      _ExtentX        =   17833
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
      Left            =   315
      TabIndex        =   0
      Tag             =   "CustCode"
      Top             =   285
      Visible         =   0   'False
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
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   1170
      Left            =   315
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Tag             =   "ActionRequested"
      Top             =   4860
      Visible         =   0   'False
      Width           =   10110
   End
   Begin VB.TextBox txtReason 
      Appearance      =   0  'Flat
      Height          =   1140
      Left            =   315
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Tag             =   "Reason"
      Top             =   3480
      Visible         =   0   'False
      Width           =   10110
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      Height          =   1035
      Left            =   315
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Tag             =   "Description"
      Top             =   2145
      Visible         =   0   'False
      Width           =   10110
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   780
      Top             =   10005
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
      TabIndex        =   31
      Top             =   9555
      Visible         =   0   'False
      Width           =   9690
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
      Left            =   10020
      TabIndex        =   30
      Top             =   9540
      Visible         =   0   'False
      Width           =   405
   End
   Begin ComboList.Usercontrol1 cmbItems 
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Top             =   6915
      Visible         =   0   'False
      Width           =   3660
      _ExtentX        =   6456
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
      TabIndex        =   10
      Top             =   6915
      Visible         =   0   'False
      Width           =   2070
      _ExtentX        =   3651
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
      Top             =   975
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   108068867
      CurrentDate     =   40698
   End
   Begin MSComDlg.CommonDialog CD2 
      Left            =   330
      Top             =   10005
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Please Select Employee's Picture"
      Filter          =   "All Picture Files|*.bmp;*.jpg;*.gif;*.jpeg"
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   0
      Left            =   2850
      TabIndex        =   40
      Tag             =   "CreditNoteNo"
      Top             =   975
      Width           =   2520
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "4445;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Credit Note No."
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   10
      Left            =   2865
      TabIndex        =   39
      Top             =   750
      Width           =   1125
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Company"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   9
      Left            =   330
      TabIndex        =   37
      Top             =   105
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Unit"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   8
      Left            =   8745
      TabIndex        =   36
      Top             =   765
      Visible         =   0   'False
      Width           =   285
   End
   Begin MSForms.TextBox txtItemReason 
      Height          =   315
      Left            =   7230
      TabIndex        =   13
      Top             =   6900
      Visible         =   0   'False
      Width           =   1620
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      Size            =   "2857;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Reason"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   6
      Left            =   7245
      TabIndex        =   35
      Top             =   6690
      Visible         =   0   'False
      Width           =   540
   End
   Begin MSForms.ComboBox cmbUnit 
      Height          =   315
      Left            =   8745
      TabIndex        =   4
      Tag             =   "Unit"
      Top             =   975
      Visible         =   0   'False
      Width           =   1650
      VariousPropertyBits=   746604571
      DisplayStyle    =   3
      Size            =   "2910;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbCc 
      Height          =   315
      Left            =   5385
      TabIndex        =   9
      Tag             =   "EmailCC"
      Top             =   6285
      Visible         =   0   'False
      Width           =   5040
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "8890;556"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbTo 
      Height          =   315
      Left            =   315
      TabIndex        =   8
      Tag             =   "EmailTo"
      Top             =   6285
      Visible         =   0   'False
      Width           =   5040
      VariousPropertyBits=   746604571
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "8890;556"
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
      Caption         =   "Cc"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   9
      Left            =   5385
      TabIndex        =   34
      Top             =   6060
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   8
      Left            =   315
      TabIndex        =   33
      Top             =   6060
      Visible         =   0   'False
      Width           =   180
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
      TabIndex        =   32
      Top             =   9270
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   1
      Left            =   345
      TabIndex        =   29
      Top             =   1950
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Reference"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   3
      Left            =   8895
      TabIndex        =   28
      Top             =   6690
      Visible         =   0   'False
      Width           =   750
   End
   Begin MSForms.TextBox txtEmailRef 
      Height          =   315
      Left            =   8865
      TabIndex        =   14
      Top             =   6900
      Visible         =   0   'False
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
      TabIndex        =   27
      Top             =   6690
      Visible         =   0   'False
      Width           =   270
   End
   Begin MSForms.TextBox txtQty 
      Height          =   315
      Left            =   6075
      TabIndex        =   12
      Top             =   6900
      Visible         =   0   'False
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
      TabIndex        =   26
      Top             =   6690
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Invoice No."
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   0
      Left            =   345
      TabIndex        =   25
      Top             =   6690
      Visible         =   0   'False
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
      TabIndex        =   23
      Top             =   4650
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Country"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   7
      Left            =   2670
      TabIndex        =   22
      Top             =   75
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Customer"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   4
      Left            =   345
      TabIndex        =   21
      Top             =   75
      Visible         =   0   'False
      Width           =   690
   End
   Begin MSForms.ComboBox cmbCountry 
      Height          =   315
      Left            =   2655
      TabIndex        =   1
      Tag             =   "Country"
      Top             =   285
      Visible         =   0   'False
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
      Caption         =   "Reason for Complaint"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   19
      Left            =   345
      TabIndex        =   20
      Top             =   3270
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Complaint No."
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   73
      Left            =   315
      TabIndex        =   19
      Top             =   765
      Width           =   1005
   End
   Begin MSForms.TextBox TBox 
      Height          =   315
      Index           =   14
      Left            =   315
      TabIndex        =   2
      Tag             =   "ComplaintNo"
      Top             =   975
      Width           =   2520
      VariousPropertyBits=   679495711
      BorderStyle     =   1
      Size            =   "4445;556"
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label label1 
      AutoSize        =   -1  'True
      Caption         =   "Complaint Date"
      ForeColor       =   &H00400000&
      Height          =   195
      Index           =   5
      Left            =   7125
      TabIndex        =   18
      Top             =   765
      Visible         =   0   'False
      Width           =   1095
   End
   Begin MSForms.CommandButton cmdCancel 
      Height          =   420
      Left            =   8865
      TabIndex        =   17
      Top             =   10095
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
      Left            =   7230
      TabIndex        =   16
      Top             =   10095
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
      Left            =   5595
      TabIndex        =   15
      Top             =   10095
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
Attribute VB_Name = "frmCustomerCreditNoteAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lEntryID As Long
Dim strDelList As String
Dim iColNo As Integer

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
    
'    cmbCompany.ListHeight = 6000
'    cmbCompany.AddVals con, "CompanyName", "Companies", "EntryID", " Order By CompanyName"
'    cmbCompany.id = 1
'
'    cmbCust.ListHeight = 2500
'    Call cmbCust.AddVals(con, "DISTINCT CustCode", "ForeignCustomers", "CustCode")
'
'    Call AddToCombo(cmbTo, "ComplaintFormEMailList", "ComplaintFormEMailList", , , "ComplaintFormEMailList")
'    Call AddToCombo(cmbCc, "ComplaintFormEMailList", "ComplaintFormEMailList", , , "ComplaintFormEMailList")
'
'    cmbInvoiceNo.ListHeight = 2500
'    cmbItems.ListHeight = 2500
'
'    With cmbPenalty
'        .AddItem "None"
'        .AddItem "Free Placement"
'        .AddItem "Credit Note"
'        .AddItem "Investigation"
'    End With
'
'    cmbUnit.AddItem "Unit 1"
'    cmbUnit.List(0, 1) = "Unit 1"
'    cmbUnit.AddItem "Unit 2"
'    cmbUnit.List(1, 1) = "Unit 2"
    
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
    TBox(14) = GetNewComplaintNo(DTPicker2.Value)
End Sub

Private Sub Form_Load()
    
    Me.KeyPreview = True
    DTPicker2.Value = Date
    strDelList = ""
    Call DTPicker2_Change
    
    Call FillCmbs
    
End Sub

Private Function SaveItem() As Boolean

    On Error GoTo err
    
     
    Call StartTrans(con)
   
    con.Execute "UPDATE CustomerComplaints SET CreditNoteNo='" & TBox(0) & "' WHERE EntryID=" & lEntryID
   
    
    Dim i As Integer
    Dim lCC_D_EntryID As Long
    For i = 1 To LV.ListItems.count
        With LV.ListItems(i)
        
            If Val(.Tag) = 0 Then
                'con.Execute "INSERT INTO CustomerComplaints_Detail(RefID,InvoiceNo,ItemID,Qty,ItemReason,EmailReference) VALUES(" & lmyEntryID & _
                 ",'" & .Text & "','" & .ListSubItems(1).Tag & "'," & Val(.SubItems(2)) & ",'" & .SubItems(3) & "','" & .SubItems(4) & "')"
                 
                lCC_D_EntryID = GetSingleLongValue("MAX(EntryID)", "CustomerComplaints_Detail")
            Else
                lCC_D_EntryID = Val(.Tag)
                con.Execute "UPDATE CustomerComplaints_Detail SET Price=" & Val(.SubItems(3)) & " WHERE EntryID=" & lCC_D_EntryID
            End If
             
        End With
    Next
    con.CommitTrans
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
        
        If (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is MSForms.ComboBox) Or (TypeOf c Is MSForms.CheckBox) Or (TypeOf c Is ComboList.Usercontrol1) Or (TypeOf c Is DTPicker) Or (TypeOf c Is VB.TextBox) Then
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
            ElseIf (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is VB.TextBox) Then
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
            ElseIf (TypeOf c Is MSForms.TextBox) Or (TypeOf c Is VB.TextBox) Then   '
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
        
'        For Each c In Controls
'            If c.Tag <> "" Then
'                If TypeOf c Is MSForms.CheckBox Then
'                    c.Value = .Fields(c.Tag)
'                ElseIf TypeOf c Is ComboList.Usercontrol1 Then
'                    'C.Text = .Fields(C.Tag & "Text") & ""
'                    c.id = .Fields(c.Tag) & ""
'                ElseIf TypeOf c Is DTPicker Then
'                    If Not IsNull(.Fields(c.Tag)) Then
'                        c.Value = .Fields(c.Tag)
'                    End If
'                ElseIf c.Name = "cmbPenalty" Then
'                    c.ListIndex = Val(.Fields(c.Tag) & "")
'                Else
'                    c.Text = .Fields(c.Tag) & ""
'                End If
'            End If
'        Next
        TBox(14) = !ComplaintNo & ""
        TBox(0) = !CreditNoteNo & ""
        .Close
        
        If TBox(0) = "" Then
            TBox(0) = GetNewCreditNoteNo(Date)
        End If
        
        Dim ITM As ListItem
        Dim strPicAttach As String
        .Open "SELECT * FROM VCustomerComplaints_Detail WHERE RefID=" & lEntryID & " ORDER BY EntryID", con, adOpenForwardOnly, adLockReadOnly
        LV.ListItems.Clear
        
        Do Until .EOF
        
            Set ITM = LV.ListItems.add(, , !InvoiceNo & "")
            ITM.Tag = !EntryID
            ITM.ListSubItems.add(, , !ITemID & " {" & !ItemCode & "}").Tag = !ITemID & ""
            ITM.ListSubItems.add , , Val(!Qty & "")
            ITM.ListSubItems.add , , Val(!Price & "")
             
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
    
    TBox(14) = GetNewComplaintNo(DTPicker2.Value)
    Me.Show 1
    
End Sub

Private Function GetNewComplaintNo(myDT As Date) As String
    'CC-1/2011
    Dim lComplaintNo As Long
    lComplaintNo = GetSingleLongValue("MAX(CAST(SUBSTRING(ComplaintNo,4,CHARINDEX('/',ComplaintNo)-4) AS INT))", "CustomerComplaints", " WHERE Type=0")
    lComplaintNo = lComplaintNo + 1
    
    GetNewComplaintNo = "CC-" & lComplaintNo & "/" & year(myDT)
    
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

Private Sub LV_DblClick()

    If LV.ListItems.count = 0 Then Exit Sub
    iColNo = 4
    With txtPrice
        .Move LV.Left + LV.ColumnHeaders(iColNo).Left, LV.Top + LV.SelectedItem.Top, LV.ColumnHeaders(iColNo).Width
        .Text = Val(LV.SelectedItem.SubItems(iColNo - 1))
        .SelStart = 0
        .SelLength = Len(.Text)
        .Visible = True
        .SetFocus
    End With
End Sub

Private Sub LV_KeyDown(KeyCode As Integer, Shift As Integer)
        
    Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    If KeyCode = 46 Then
        If MsgBox("Are you sure to Delete?", vbQuestion + vbYesNo) = vbYes Then
            If Val(LV.SelectedItem.Tag) > 0 Then
                strDelList = strDelList & Val(LV.SelectedItem.Tag) & ","
            End If
            LV.ListItems.Remove LV.SelectedItem.Index
        End If
    End If
End Sub

Private Sub LV_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button <> 2 Then Exit Sub
    If LV.ListItems.count = 0 Then Exit Sub
    
    Dim strRet As String
    If LV.SelectedItem.SubItems(5) = "" Or LV.SelectedItem.SubItems(5) = "Remove" Then 'Not Attached
        strRet = ShowPopUpMenu1("Attach Pic")
    Else
        strRet = ShowPopUpMenu1("View Pic", "-", "Remove Pic")
    End If
    
    If strRet = "Attach Pic" Then
        
        On Error Resume Next
        Dim strPicfile As String
        CD2.ShowSave
    
        If Not (err.Number = 32755) Then
            strPicfile = CD2.FileName
            LV.SelectedItem.SubItems(5) = "Add"
            LV.SelectedItem.ListSubItems(5).Tag = strPicfile
        End If
        
    ElseIf strRet = "Remove Pic" Then
        LV.SelectedItem.SubItems(5) = "Remove"
        LV.SelectedItem.ListSubItems(5).Tag = ""
    ElseIf strRet = "View Pic" Then
        'View Pic...
        Dim f As New frmViewPictureFSExt
        Load f
        f.ShowFromDBField "ItemPic", "CustomerComplaints_Detail", " WHERE EntryID=" & Val(LV.SelectedItem.Tag)
    End If
     
End Sub

Private Sub txtEmailRef_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    If KeyCode.Value = 13 Then
        Dim ITM As ListItem
        Set ITM = LV.ListItems.add(, , cmbInvoiceNo.Text)
        ITM.ListSubItems.add(, , cmbItems.Text).Tag = cmbItems.ID
        ITM.ListSubItems.add , , Val(txtQty)
        ITM.ListSubItems.add , , txtItemReason
        ITM.ListSubItems.add , , txtEmailRef
        ITM.ListSubItems.add , , ""
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

Private Sub AttachPic(p_lEntryID As Long, strFilePath As String)
    
    On Error Resume Next
    If strFilePath <> "" Then
        Dim Rec As New ADODB.Recordset
        Dim rsPic As New ADODB.Stream
        Rec.Open "SELECT EntryID,ItemPic FROM CustomerComplaints_Detail WHERE EntryID=" & p_lEntryID, con, adOpenStatic, adLockOptimistic
        
        If Rec.RecordCount > 0 Then
            rsPic.Open
            rsPic.Type = adTypeBinary
            rsPic.LoadFromFile strFilePath
            Rec.Fields(1) = rsPic.Read
            rsPic.Close
            Rec.Update
        End If
        
        Rec.Close
        
        Set Rec = Nothing
        Set rsPic = Nothing
        
    End If

End Sub


Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LV.SelectedItem.SubItems(iColNo - 1) = Val(txtPrice)
        txtPrice.Visible = False
    ElseIf KeyAscii = 27 Then
        txtPrice.Visible = False
    End If
End Sub

Private Sub txtPrice_LostFocus()
    txtPrice.Visible = False
End Sub

Private Function GetNewCreditNoteNo(myDT As Date) As String
    'CR-1/2011
    Dim lCreditNoteNo As Long
    lCreditNoteNo = GetSingleLongValue("MAX(CAST(SUBSTRING(CreditNoteNo,4,CHARINDEX('/',CreditNoteNo)-4) AS INT))", "CustomerComplaints", " WHERE ISNULL(CreditNoteNo,'')<>''")
    lCreditNoteNo = lCreditNoteNo + 1
    
    GetNewCreditNoteNo = "CR-" & lCreditNoteNo & "/" & year(myDT)
    
End Function

