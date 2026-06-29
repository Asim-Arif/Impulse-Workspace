VERSION 5.00
Begin VB.Form frmCVInterviewQuestions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   10230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCVInterviewQuestions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lCVID As Long
Dim bReturnValue As Boolean
Dim strQuestions() As String

Public Function ShowMe(p_lCVID As Long) As Boolean
    lCVID = p_lCVID
    
    Dim rs As New ADODB.Recordset
    With rs
        .Open "SELECT * FROM "
    End With
    Set rs = Nothing
    
    Me.Show 1
    
End Function

Private Sub Form_Load()

    lCVID = 0
    bReturnValue = False
    
    Dim rs As New ADODB.Recordset, iIndex As Integer
    With rs
        .Open "SELECT * FROM GeneralInterviewQuestions WHERE Active=1 ORDER BY SortNo", con, adOpenForwardOnly, adLockReadOnly
        If .EOF = False Then
            ReDim strQuestions(.RecordCount)
            iIndex = 0
            Do Until .EOF
                strQuestions(iIndex) = !Question & ""
                iIndex = iIndex + 1
                .MoveNEXT
            Loop
        End If
        .Close
    End With
    Set rs = Nothing
    
End Sub
