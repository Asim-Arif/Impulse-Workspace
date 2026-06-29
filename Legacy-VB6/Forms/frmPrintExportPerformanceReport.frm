VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmPrintExportPerformanceReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print Export Performance"
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
   Icon            =   "frmPrintExportPerformanceReport.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbCustomer 
      Height          =   285
      Left            =   150
      TabIndex        =   12
      Top             =   315
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
   Begin ComboList.Usercontrol1 cmbItemGroup 
      Height          =   285
      Left            =   150
      TabIndex        =   4
      Top             =   900
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
   Begin ComboList.Usercontrol1 cmbCatagory 
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Top             =   1440
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
   Begin ComboList.Usercontrol1 cmbAdditionalGroup 
      Height          =   285
      Left            =   135
      TabIndex        =   6
      Top             =   2070
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
   Begin MSComCtl2.DTPicker DTFrom 
      Height          =   330
      Left            =   150
      TabIndex        =   9
      Top             =   2955
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   98828291
      CurrentDate     =   40925
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
   Begin MSComCtl2.DTPicker DTTo 
      Height          =   330
      Left            =   1590
      TabIndex        =   11
      Top             =   2955
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   98828291
      CurrentDate     =   40925
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Customer"
      Height          =   195
      Index           =   5
      Left            =   150
      TabIndex        =   13
      Top             =   120
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date To"
      Height          =   195
      Index           =   4
      Left            =   1605
      TabIndex        =   10
      Top             =   2715
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Date from"
      Height          =   195
      Index           =   3
      Left            =   165
      TabIndex        =   8
      Top             =   2715
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Additional Group"
      Height          =   195
      Index           =   2
      Left            =   165
      TabIndex        =   7
      Top             =   1875
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Item Groups"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   705
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Catagory"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   1245
      Width           =   675
   End
End
Attribute VB_Name = "frmPrintExportPerformanceReport"
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
      
    cmbCatagory.ListHeight = 4000
    cmbCatagory.Functionality = 0
    cmbCatagory.AddItem "<All Catagories>", "0"
    Call cmbCatagory.AddVals(con, "Description", "ItemCatagories", "CatID")
    cmbCatagory.id = "0"
    
    '-----------------------------------------------------------------------
    cmbItemGroup.ListHeight = 4000
    cmbItemGroup.Functionality = 0
    cmbItemGroup.AddItem "<All Groups>", "0"
    Call cmbItemGroup.AddVals(con, "Description", "ItemGroups", "ID")
    cmbItemGroup.id = "0"
    
    cmbAdditionalGroup.ListHeight = 4000
    cmbAdditionalGroup.Functionality = 0
    cmbAdditionalGroup.AddItem "<All Groups>", "0"
    Call cmbAdditionalGroup.AddVals(con, "MainGroupName", "ItemGroupsMain", "MainGroupID")
    cmbAdditionalGroup.id = "0"
    
    cmbCustomer.ListHeight = 4000
    cmbCustomer.Functionality = 0
    cmbCustomer.AddItem "<All Customers>", "0"
    Call cmbCustomer.AddVals(con, "DISTINCT CustCode", "ForeignCustomers", "CustCode")
    cmbCustomer.id = "0"
    
    DtFrom.Value = DateAdd("m", -6, Date)
    DtTo.Value = Date
    
End Sub

Private Sub OKButton_Click()
    
    If cmbItemGroup.MatchFound = False Then
        MsgBox "Please Select Group.", vbInformation
        Exit Sub
    End If
    
    If cmbCatagory.MatchFound = False Then
        MsgBox "Please Select Item Catagory.", vbInformation
        Exit Sub
    End If
    
    If cmbAdditionalGroup.MatchFound = False Then
        MsgBox "Please Select Additional Group.", vbInformation
        Exit Sub
    End If
     
     
    Dim rpt As CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\ExportPerformanceReport.rpt")
    
    rpt.ParameterFields.GetItemByName("@DTFrom").SetCurrentValue DtFrom.Value
    rpt.ParameterFields.GetItemByName("@DTTo").SetCurrentValue DtTo.Value
    
    If cmbItemGroup.id <> "0" Then
        rpt.ParameterFields.GetItemByName("@ItemGroup").SetCurrentValue CLng(cmbItemGroup.id)
    Else
        rpt.ParameterFields.GetItemByName("@ItemGroup").SetCurrentValue 0
    End If
    
    If cmbCatagory.id <> "0" Then
        rpt.ParameterFields.GetItemByName("@CatID").SetCurrentValue cmbCatagory.id
    Else
        rpt.ParameterFields.GetItemByName("@CatID").SetCurrentValue ""
    End If
    
    If cmbAdditionalGroup.id <> "0" Then
        rpt.ParameterFields.GetItemByName("@AdditionalGroup").SetCurrentValue CLng(cmbAdditionalGroup.id)
    Else
        rpt.ParameterFields.GetItemByName("@AdditionalGroup").SetCurrentValue 0
    End If
    
    If cmbCustomer.id <> "0" Then
        rpt.ParameterFields.GetItemByName("@Customers").SetCurrentValue cmbCustomer.id
    Else
        rpt.ParameterFields.GetItemByName("@Customers").SetCurrentValue ""
    End If
    
    rpt.EnableParameterPrompting = False
    
    Dim f As New frmPrevRpt
    
    Me.Hide
    f.ShowReport "", rpt, "", False, , Me
    
End Sub

Public Function getData(ByRef p_strCustCode As String, ByRef p_strItemGroups As String) As Boolean
     
    Me.Show 1
    
    p_strCustCode = strCustCode
    p_strItemGroups = strItemGroups
    
    getData = Not bCancelled
    
End Function
