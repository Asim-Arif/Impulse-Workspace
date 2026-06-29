VERSION 5.00
Object = "{4FC0C304-31EA-42CE-83E4-E42C3EB6A6AC}#4.2#0"; "ComboList.ocx"
Begin VB.Form frmPrintRMList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print RM List"
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
   Icon            =   "frmPrintRMList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComboList.Usercontrol1 cmbStoreName 
      Height          =   285
      Left            =   150
      TabIndex        =   6
      Top             =   435
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
   Begin ComboList.Usercontrol1 cmbRMGroup 
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Top             =   1245
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
   Begin ComboList.Usercontrol1 cmbRM 
      Height          =   285
      Left            =   150
      TabIndex        =   4
      Top             =   2205
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
   Begin VB.CheckBox chkLocationwiseStock 
      Caption         =   "Print Locationwise Stock"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   150
      TabIndex        =   9
      Top             =   3900
      Width           =   2580
   End
   Begin VB.CheckBox chkPrintZeroQty 
      Caption         =   "Print Zero Qty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   150
      TabIndex        =   8
      Top             =   3045
      Width           =   1755
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4575
      TabIndex        =   1
      Top             =   6705
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3330
      TabIndex        =   0
      Top             =   6705
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Store Name"
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
      Left            =   165
      TabIndex        =   7
      Top             =   240
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "RM Name"
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
      Left            =   180
      TabIndex        =   5
      Top             =   2010
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "RM Group"
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
      Left            =   180
      TabIndex        =   3
      Top             =   1050
      Width           =   825
   End
End
Attribute VB_Name = "frmPrintRMList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bCancelled As Boolean

Private Sub CancelButton_Click()
    bCancelled = True
    Unload Me
End Sub

Private Sub Form_Load()
      
    cmbStoreName.ListHeight = 4000
    cmbStoreName.Functionality = 0
    cmbStoreName.AddItem "<All Stores>", "0"
    Call cmbStoreName.AddVals(con, "StoreName", "Stores", "EntryID")
    cmbStoreName.ID = "0"
    
    '-----------------------------------------------------------------------
    cmbRMGroup.ListHeight = 4000
    cmbRMGroup.Functionality = 0
    cmbRMGroup.AddItem "<All Groups>", "0"
    Call cmbRMGroup.AddVals(con, "Description", "RMGroups", "ID")
    cmbRMGroup.ID = "0"
        
    cmbRM.ListHeight = 4000
    cmbRM.Functionality = 0
    cmbRM.AddItem "<All>", "0"
    Call cmbRM.AddVals(con, "RMID1+' '+RMName", "RM", "RMID")
    cmbRM.ID = "0"
    
End Sub

Private Sub cmdPrint_Click()
    
    If cmbStoreName.MatchFound = False Then
        MsgBox "Please Select Store.", vbInformation
        Exit Sub
    End If
    
    If cmbRMGroup.MatchFound = False Then
        MsgBox "Please Select  Group.", vbInformation
        Exit Sub
    End If
    
    If cmbRM.MatchFound = False Then
        MsgBox "Please Select RM.", vbInformation
        Exit Sub
    End If
    
    Dim strSelection As String
    strSelection = "1=1"
    If cmbStoreName.ID <> "0" Then
        strSelection = strSelection & " AND {VMaterialLocationWiseBalances.Store_RefID}=" & cmbStoreName.ID
    End If
    If cmbRMGroup.ID <> "0" Then
        strSelection = strSelection & " AND {VRM.GroupID}=" & cmbRMGroup.ID
    End If
    If cmbRM.ID <> "0" Then
        strSelection = strSelection & " AND {VRM.RMID}=" & cmbRM.ID
    End If
    
    If chkPrintZeroQty.Value = 0 Then
        strSelection = strSelection & " AND {VRM.QtyInStock}>0"
    End If
    
    Dim rpt As New CRAXDDRT.Report
    Set rpt = rptApp.OpenReport(RptPath & "\RMList.rpt")
    rpt.FormulaFields.GetItemByName("HideLocations").Text = Not (chkLocationwiseStock.Value = 1)
    
    Dim f As New frmPrevRpt
    Me.Hide
    f.ShowReport strSelection, rpt
    Unload Me
    
End Sub

