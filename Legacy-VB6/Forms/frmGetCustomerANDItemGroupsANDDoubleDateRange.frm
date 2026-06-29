VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmGetCustomerANDItemGroupsANDDoubleDateRange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Customer & Item Groups"
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
   Icon            =   "frmGetCustomerANDItemGroupsANDDoubleDateRange.frx":0000
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
      Top             =   300
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
      Top             =   840
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
   Begin VB.Frame FraLower 
      Caption         =   "Order"
      Height          =   1170
      Left            =   150
      TabIndex        =   11
      Top             =   3210
      Width           =   5520
      Begin MSComCtl2.DTPicker DTFrom2 
         Height          =   315
         Left            =   270
         TabIndex        =   12
         Top             =   570
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dddd, dd-MMM-yyyy"
         Format          =   114884611
         CurrentDate     =   39372
      End
      Begin MSComCtl2.DTPicker DTTo2 
         Height          =   315
         Left            =   2745
         TabIndex        =   13
         Top             =   570
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dddd, dd-MMM-yyyy"
         Format          =   114884611
         CurrentDate     =   39372
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&To Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   5
         Left            =   2730
         TabIndex        =   15
         Top             =   345
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&From Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   285
         TabIndex        =   14
         Top             =   330
         Width           =   885
      End
   End
   Begin VB.Frame FraUpper 
      Caption         =   "Packing"
      Height          =   1170
      Left            =   150
      TabIndex        =   6
      Top             =   1455
      Width           =   5520
      Begin MSComCtl2.DTPicker DTFrom1 
         Height          =   315
         Left            =   270
         TabIndex        =   7
         Top             =   570
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dddd, dd-MMM-yyyy"
         Format          =   114884611
         CurrentDate     =   39372
      End
      Begin MSComCtl2.DTPicker DTTo1 
         Height          =   315
         Left            =   2745
         TabIndex        =   8
         Top             =   570
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dddd, dd-MMM-yyyy"
         Format          =   114884611
         CurrentDate     =   39372
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&From Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   3
         Left            =   285
         TabIndex        =   10
         Top             =   330
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "&To Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   2730
         TabIndex        =   9
         Top             =   345
         Width           =   660
      End
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
      Caption         =   "Item Groups"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   645
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Customer"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   105
      Width           =   690
   End
End
Attribute VB_Name = "frmGetCustomerANDItemGroupsANDDoubleDateRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean

Dim strItemGroups As String, strCustCode As String
Dim myDTFrom1 As Date, myDTTo1 As Date
Dim myDTFrom2 As Date, myDTTo2 As Date

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub


Private Sub Form_Load()
      
    cmbCust.ListHeight = 4000
    cmbCust.Functionality = 1
    cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode")
    cmbCust.id = "0"
    
    '-----------------------------------------------------------------------
    cmbItemGroup.ListHeight = 4000
    cmbItemGroup.Functionality = 1
    cmbItemGroup.AddItem "<All Groups>", "0"
    Call cmbItemGroup.AddVals(con, "Description", "ItemGroups", "ID")
    cmbItemGroup.id = "0"
    
    
    DTFrom1.Value = Date
    DTFrom2.Value = Date
    DTTo1.Value = Date
    DTTo2.Value = Date
    
End Sub

Private Sub OKButton_Click()
    
    If cmbCust.MatchFound = False Then
        MsgBox "Please Select Customer.", vbInformation
        Exit Sub
    End If
    
    If cmbItemGroup.MatchFound = False Then
        MsgBox "Please Select Item Group.", vbInformation
        Exit Sub
    End If
    
    bCancelled = False
    
    strCustCode = cmbCust.id
    strItemGroups = cmbItemGroup.id
    
    myDTFrom1 = DTFrom1
    myDTFrom2 = DTFrom2
    myDTTo1 = DTTo1
    myDTTo2 = DTTo2
    
    Unload Me
    
End Sub

Public Function getData(ByRef p_strCustCode As String, ByRef p_strItemGroups As String, ByRef p_DTFrom1 As Date, ByRef p_DTTo1 As Date, p_DTFrom2 As Date, p_DTTo2 As Date) As Boolean
     
    Me.Show 1
    
    p_strCustCode = strCustCode
    p_strItemGroups = strItemGroups
        
    p_DTFrom1 = myDTFrom1
    p_DTFrom2 = myDTFrom2
    
    p_DTTo1 = myDTTo1
    p_DTTo2 = myDTTo2
    
    getData = Not bCancelled
    
End Function
