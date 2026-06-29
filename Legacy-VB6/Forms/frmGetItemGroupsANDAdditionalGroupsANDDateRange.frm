VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmGetItemGroupsANDAdditionalGroupsANDDateRange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Item Groups"
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
   Icon            =   "frmGetItemGroupsANDAdditionalGroupsANDDateRange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbItemGroup 
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Top             =   585
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
      Left            =   150
      TabIndex        =   9
      Top             =   1275
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
   Begin VB.Frame FraUpper 
      Height          =   1170
      Left            =   150
      TabIndex        =   4
      Top             =   2295
      Width           =   5520
      Begin MSComCtl2.DTPicker DTFrom1 
         Height          =   315
         Left            =   270
         TabIndex        =   5
         Top             =   570
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dddd, dd-MMM-yyyy"
         Format          =   143523843
         CurrentDate     =   39372
      End
      Begin MSComCtl2.DTPicker DTTo1 
         Height          =   315
         Left            =   2745
         TabIndex        =   6
         Top             =   570
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dddd, dd-MMM-yyyy"
         Format          =   143523843
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
         TabIndex        =   8
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
         TabIndex        =   7
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
      Caption         =   "Additional Group"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   10
      Top             =   1080
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Item Groups"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   390
      Width           =   885
   End
End
Attribute VB_Name = "frmGetItemGroupsANDAdditionalGroupsANDDateRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean

Dim strItemGroups As String, strCustCode As String, strAdditionalGroups As String
Dim myDTFrom1 As Date, myDTTo1 As Date

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub


Private Sub Form_Load()
    
    '-----------------------------------------------------------------------
    cmbItemGroup.ListHeight = 4000
    cmbItemGroup.Functionality = 1
    cmbItemGroup.AddItem "<All Groups>", "0"
    Call cmbItemGroup.AddVals(con, "Description", "ItemGroups", "ID")
    cmbItemGroup.ID = "0"
    
    cmbAdditionalGroup.ListHeight = 4000
    cmbAdditionalGroup.Functionality = 1
    cmbAdditionalGroup.AddItem "<All Groups>", "0"
    Call cmbAdditionalGroup.AddVals(con, "MainGroupName", "ItemGroupsMain", "MainGroupID")
    cmbAdditionalGroup.ID = "0"
    
    DTFrom1.Value = Date
    DTTo1.Value = Date
    
End Sub

Private Sub OKButton_Click()
    
    
    If cmbItemGroup.MatchFound = False Then
        MsgBox "Please Select Item Group.", vbInformation
        Exit Sub
    ElseIf cmbAdditionalGroup.MatchFound = False Then
        MsgBox "Please Select Additional Group.", vbInformation
        Exit Sub
    End If
    
    bCancelled = False
    
 
    strItemGroups = cmbItemGroup.ID
    strAdditionalGroups = cmbAdditionalGroup.ID
    
    myDTFrom1 = DTFrom1
   
    myDTTo1 = DTTo1
  
    
    Unload Me
    
End Sub

Public Function getData(ByRef p_strItemGroups As String, ByRef p_strAdditionalGroups As String, ByRef p_DTFrom1 As Date, ByRef p_DTTo1 As Date) As Boolean
     
    Me.Show 1
    

    p_strItemGroups = strItemGroups
    p_strAdditionalGroups = strAdditionalGroups
    
    p_DTFrom1 = myDTFrom1
    
    p_DTTo1 = myDTTo1

    
    getData = Not bCancelled
    
End Function
