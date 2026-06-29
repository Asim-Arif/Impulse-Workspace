VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmPackingLabels 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Packing Outer Labels"
   ClientHeight    =   2115
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPrintInner 
      Caption         =   "Print Inner"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3645
      TabIndex        =   11
      Top             =   1035
      Visible         =   0   'False
      Width           =   1290
   End
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   45
      TabIndex        =   2
      Top             =   375
      Width           =   1740
      _ExtentX        =   3069
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
   Begin VB.TextBox txtCurrency 
      Appearance      =   0  'Flat
      DataField       =   "RptTime"
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   45
      MaxLength       =   150
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1740
   End
   Begin MSForms.CommandButton cmdPrintItemLabel 
      Height          =   375
      Left            =   3645
      TabIndex        =   12
      Top             =   1665
      Visible         =   0   'False
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Item Label"
      PicturePosition =   327683
      Size            =   "3096;661"
      Accelerator     =   71
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.ComboBox cmbTo 
      Height          =   285
      Left            =   1815
      TabIndex        =   10
      Top             =   1755
      Width           =   1725
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3043;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin MSForms.ComboBox cmbFrom 
      Height          =   285
      Left            =   45
      TabIndex        =   9
      Top             =   1755
      Width           =   1740
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3069;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " From Carton No.          To Caron No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   1
      Left            =   45
      TabIndex        =   8
      Top             =   1470
      Width           =   3495
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Currency                      Invoice No"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   0
      Left            =   45
      TabIndex        =   7
      Top             =   810
      Width           =   3495
   End
   Begin MSForms.ComboBox cmbCountry 
      Height          =   285
      Left            =   1815
      TabIndex        =   6
      Top             =   375
      Width           =   1725
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3043;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C5D1DC&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Customer                     Country"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Index           =   7
      Left            =   45
      TabIndex        =   5
      Top             =   105
      Width           =   3495
   End
   Begin MSForms.ComboBox cmbCustomInvoice 
      Height          =   285
      Left            =   1815
      TabIndex        =   4
      Top             =   1080
      Width           =   1725
      VariousPropertyBits=   679495707
      BorderStyle     =   1
      DisplayStyle    =   3
      Size            =   "3043;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      SpecialEffect   =   0
      FontName        =   "Tahoma"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
   End
   Begin VB.Image Image2 
      Height          =   435
      Left            =   4965
      Picture         =   "frmPackingLabels.frx":0000
      Top             =   1020
      Width           =   480
   End
   Begin MSForms.CommandButton cmdClose 
      Height          =   360
      Left            =   3645
      TabIndex        =   1
      Top             =   555
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Cancel       "
      PicturePosition =   327683
      Size            =   "3096;635"
      Picture         =   "frmPackingLabels.frx":05EF
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
      Left            =   3645
      TabIndex        =   0
      Top             =   105
      Width           =   1755
      ForeColor       =   0
      Caption         =   "Print           "
      PicturePosition =   327683
      Size            =   "3096;635"
      Picture         =   "frmPackingLabels.frx":0701
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
Attribute VB_Name = "frmPackingLabels"
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

Private Sub cmdClose_Click()
    Unload Me
End Sub




Private Sub cmdPrint_Click()

    On Error GoTo err
    Dim strInvoice As String
    
    If Not (cmbFrom.MatchFound And cmbTo.MatchFound) Then
        MsgBox "Invalid Carton No. Selected.", vbInformation
        Exit Sub
    End If
    
    strInvoice = cmbCustomInvoice.Text
    
    Screen.MousePointer = vbHourglass
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT ID,CartonFrom,CartonTo FROM CustomPList WHERE CustomInvoice='" & strInvoice & "' ORDER BY CartonFrom,ID", con, adOpenForwardOnly, adLockReadOnly
        con.Execute "DELETE FROM PrintPackingLabels"
        Do Until .EOF
            For i = !CartonFrom To !CartonTo
                con.Execute "INSERT INTO PrintPackingLabels(RefID,CartonNo) VALUES(" & !ID & "," & i & ")"
            Next
            .MoveNEXT
        Loop
        .Close
    End With
        
    Set rs = Nothing
    Me.Hide
    Load frmPrevRpt
    Dim rpt As CRAXDDRT.Report
    If chkPrintInner.Value = vbChecked Then
        Dim strInnerFileName As String
        strInnerFileName = GetSingleStringValue("InnerPackingLabel", "ForeignCustomers", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'")
        Set rpt = rptApp.OpenReport(RptPath & "\" & strInnerFileName) '
    Else
        Dim strOuterFileName As String
        strOuterFileName = GetSingleStringValue("OuterPackingLabel", "ForeignCustomers", " WHERE CustCode='" & cmbCust.ID & "' AND Country='" & cmbCountry.Text & "'")
        Set rpt = rptApp.OpenReport(RptPath & "\" & strOuterFileName)  '
    End If
    frmPrevRpt.ShowReport "{VrptProformas.CustomInvoice}='" & strInvoice & "' AND {CustomPList.CartonFrom} In (" & cmbFrom & " to " & cmbTo & ") ", rpt
    Unload Me
    Screen.MousePointer = vbDefault
    
    
    Exit Sub
err:
    MsgBox err.Description
End Sub

Private Sub cmdPrintItemLabel_Click()

    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ItemLabel.rpt")
    
    rpt.EnableParameterPrompting = False
    rpt.ParameterFields.GetItemByName("@InvoiceNo").SetCurrentValue cmbCustomInvoice.Text
    Dim frm As New frmPrevRpt
    
    Me.Hide
    frm.ShowReport "", rpt, "", False
    Unload Me
End Sub

Private Sub Form_Load()
        
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode", " Where Active=1")
 
    cmbCust.ListHeight = 2500
    
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


