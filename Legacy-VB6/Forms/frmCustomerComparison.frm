VERSION 5.00
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "combolist.ocx"
Begin VB.Form frmCustomerComparison 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Comparison"
   ClientHeight    =   7215
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5940
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustomerComparison.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbCust 
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Top             =   330
      Width           =   5520
      _ExtentX        =   9737
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
   Begin ComboList.Usercontrol1 cmbCompareTo 
      Height          =   285
      Left            =   150
      TabIndex        =   4
      Top             =   915
      Width           =   5520
      _ExtentX        =   9737
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
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4575
      TabIndex        =   1
      Top             =   6705
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3330
      TabIndex        =   0
      Top             =   6705
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Compare to (Customer)"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Customer"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   135
      Width           =   690
   End
End
Attribute VB_Name = "frmCustomerComparison"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean

Dim strItemGroups As String, strCustCode As String

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub


Private Sub Form_Load()
      
    cmbCust.ListHeight = 4000
    cmbCust.Functionality = 0
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode")
    
    '-----------------------------------------------------------------------
    cmbCompareTo.ListHeight = 4000
    cmbCompareTo.Functionality = 1
    
    Call cmbCompareTo.AddVals(con, "CustCode", "ForeignCustomers", "CustCode")
   
End Sub

Private Sub OKButton_Click()
    
    If cmbCust.MatchFound = False Then
        MsgBox "Please Select Customer.", vbInformation
        Exit Sub
    End If
    
    If cmbCompareTo.MatchFound = False Then
        MsgBox "Please Select Compare to.", vbInformation
        Exit Sub
    End If
     
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\CustomerComparison.rpt")
    rpt.ParameterFields.GetItemByName("@CustCode").SetCurrentValue cmbCust.ID
    rpt.ParameterFields.GetItemByName("@CompareCustomers").SetCurrentValue cmbCompareTo.ID
    
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    Me.Hide
    f.ShowReport "", rpt, "", True, , Me
    
End Sub

Public Function getData(ByRef p_strCustCode As String, ByRef p_strItemGroups As String) As Boolean
     
    Me.Show 1
    
    p_strCustCode = strCustCode
    p_strItemGroups = strItemGroups
    
    getData = Not bCancelled
    
End Function
