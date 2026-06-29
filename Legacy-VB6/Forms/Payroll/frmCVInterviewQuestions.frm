VERSION 5.00
Begin VB.Form frmCVInterviewQuestions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CV Interview Questions"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPrev 
      Caption         =   "&Prev"
      Height          =   495
      Left            =   4860
      TabIndex        =   8
      Top             =   5085
      Width           =   1215
   End
   Begin VB.TextBox txtAnswer 
      Height          =   1470
      Left            =   210
      MultiLine       =   -1  'True
      TabIndex        =   7
      Top             =   3495
      Width           =   7245
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next"
      Height          =   495
      Left            =   6180
      TabIndex        =   5
      Top             =   5085
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   495
      Left            =   6270
      TabIndex        =   4
      Top             =   7590
      Width           =   1215
   End
   Begin VB.TextBox txtCVNo 
      Height          =   360
      Left            =   210
      TabIndex        =   3
      Top             =   465
      Width           =   2190
   End
   Begin VB.TextBox txtCVName 
      Height          =   360
      Left            =   2415
      TabIndex        =   1
      Top             =   465
      Width           =   5085
   End
   Begin VB.Label lblQuestion 
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1635
      Left            =   210
      TabIndex        =   6
      Top             =   1695
      Width           =   7245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CV No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   225
      TabIndex        =   2
      Top             =   240
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2445
      TabIndex        =   0
      Top             =   270
      Width           =   570
   End
End
Attribute VB_Name = "frmCVInterviewQuestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strCVID As String
Dim bReturnValue As Boolean
Dim strQuestions() As String
Dim iLoadedIndex As Integer
Dim bNoQuestions As Boolean

Public Function ShowMe(p_strCVID As String) As Boolean

    strCVID = p_strCVID
    If bNoQuestions Then
        MsgBox "No Questions Found.", vbCritical
        Unload Me
        Exit Function
    End If
    
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT CVID,Name FROM CVs WHERE CVID='" & strCVID & "'", con, adOpenForwardOnly, adLockReadOnly
        txtCVNo = !CVID & ""
        txtCVName = !Name & ""
        .Close
    End With
    Set rs = Nothing
    
    Call LoadQuestion(0)
    Me.Show 1
    
End Function

Private Sub LoadQuestion(p_iIndex As Integer)

    lblQuestion.Caption = (p_iIndex + 1) & " " & strQuestions(p_iIndex, 0)
    iLoadedIndex = p_iIndex
    Call LoadAnswer(iLoadedIndex)
    
End Sub

Private Sub cmdClose_Click()

    Call SaveAnswer(iLoadedIndex, txtAnswer)
    Unload Me
    
End Sub

Private Sub cmdNext_Click()

    If (iLoadedIndex + 1) < UBound(strQuestions) Then
        'Save Answer...
        Call SaveAnswer(iLoadedIndex, txtAnswer)
        iLoadedIndex = iLoadedIndex + 1
        LoadQuestion (iLoadedIndex)
    End If
    
End Sub

Private Sub SaveAnswer(iIndex As Integer, strAnswer As String)

    If Val(txtAnswer.Tag) > 0 Then
        con.Execute "UPDATE CVInterviewQuestions SET Answer='" & strAnswer & "' WHERE EntryID=" & Val(txtAnswer.Tag)
    Else
        con.Execute "INSERT INTO CVInterviewQuestions(CVID,GIQ_RefID,Answer) VALUES('" & strCVID & "'," & _
         Val(strQuestions(iIndex, 1)) & ",'" & strAnswer & "')"
    End If
    
End Sub

Private Sub LoadAnswer(iIndex As Integer)

    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM CVInterviewQuestions WHERE CVID='" & strCVID & "' AND GIQ_RefID=" & _
         Val(strQuestions(iIndex, 1)), con, adOpenForwardOnly, adLockReadOnly
        
        If .EOF = False Then
            txtAnswer = !Answer & ""
            txtAnswer.Tag = Val(!EntryID & "")
        Else
            txtAnswer = ""
            txtAnswer.Tag = "0"
        End If
        
        .Close
    End With
    Set rs = Nothing
    
End Sub

Private Sub cmdPrev_Click()
    If iLoadedIndex > 0 Then
        Call SaveAnswer(iLoadedIndex, txtAnswer)
        iLoadedIndex = iLoadedIndex - 1
        LoadQuestion (iLoadedIndex)
    End If
End Sub

Private Sub Form_Load()

    strCVID = ""
    bReturnValue = False
    
    Dim rs As New ADODB.Recordset, iIndex As Integer
    With rs
        .Open "SELECT * FROM GeneralInterviewQuestions WHERE Active=1 ORDER BY SortNo", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            ReDim strQuestions(.RecordCount, 2)
            iIndex = 0
            Do Until .EOF
                strQuestions(iIndex, 0) = !Question & ""
                strQuestions(iIndex, 1) = !EntryID & ""
                iIndex = iIndex + 1
                .MoveNEXT
            Loop
        Else
            bNoQuestions = True
        End If
        .Close
    End With
    Set rs = Nothing
    
End Sub
