VERSION 5.00
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmGetCustomerANDItemGroupsANDFinYear 
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
   Icon            =   "frmGetCustomerANDItemGroupsANDFinYear.frx":0000
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
      Top             =   270
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
      Top             =   810
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
   Begin VB.ComboBox cmbFinYear 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1395
      Width           =   5520
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
      Caption         =   "Year"
      Height          =   195
      Index           =   2
      Left            =   150
      TabIndex        =   7
      Top             =   1170
      Width           =   330
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Item Groups"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   615
      Width           =   885
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Customer"
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   3
      Top             =   75
      Width           =   690
   End
End
Attribute VB_Name = "frmGetCustomerANDItemGroupsANDFinYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean

Dim strItemGroups As String, strCustCode As String
Dim strYear As String

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub


Private Sub Form_Load()
      
    cmbCust.ListHeight = 4000
    cmbCust.Functionality = 0
    cmbCust.AddItem "<All Customers>", "0"
    Call cmbCust.AddVals(con, "CustCode", "ForeignCustomers", "CustCode")
    cmbCust.id = "0"
    
    '-----------------------------------------------------------------------
    cmbItemGroup.ListHeight = 4000
    cmbItemGroup.Functionality = 0
    cmbItemGroup.AddItem "<All Groups>", "0"
    Call cmbItemGroup.AddVals(con, "Description", "ItemGroups", "ID")
    cmbItemGroup.id = "0"
        
        
    cmbFinYear.Clear
    Dim iStartYear As Integer
    Dim rs As New ADODB.Recordset
    
    iStartYear = 2010
    Dim i As Integer
    For i = iStartYear To year(Date)
        cmbFinYear.AddItem i & " - " & i + 1 & " [ " & Format(DateSerial(i, 7, 1), "dd-MMM-yyyy") & " to " & Format(DateSerial(i + 1, 6, 30), "dd-MMM-yyyy") & " ]"
        cmbFinYear.ItemData(cmbFinYear.ListCount - 1) = i
    Next i
    cmbFinYear.ListIndex = 0
    
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
    
    If cmbFinYear.ListIndex = -1 Then
        MsgBox "Please Select Year.", vbInformation
        Exit Sub
    End If
    
    bCancelled = False
    
    If cmbCust.Functionality = 0 And cmbCust.id <> "0" Then
        strCustCode = "'" & cmbCust.id & "'"
    Else
        strCustCode = cmbCust.id
    End If
    
    If cmbItemGroup.Functionality = 0 And cmbItemGroup.id <> "0" Then
        strItemGroups = "'" & cmbItemGroup.id & "'"
    Else
        strItemGroups = cmbItemGroup.id
    End If
    
    strYear = cmbFinYear.Text
    
    Unload Me
    
End Sub

Public Function getData(ByRef p_strCustCode As String, ByRef p_strItemGroups As String, ByRef p_strYear As String) As Boolean
     
    Me.Show 1
    
    p_strCustCode = strCustCode
    p_strItemGroups = strItemGroups
    p_strYear = strYear
    getData = Not bCancelled
    
End Function
