VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IssueItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public ProcessID As Integer
'Public ItemCode As String
'Public CompItemCode As String
'Public ItemDescription As String
'Public lQty As Long
Public LVIndex As Long
Public OrderQty As Long
Public FreeStock As Long
Public IssueableQty As Long
Public IssQty As Long
Public IssBal As Long
Public RcvQty As Long
Public RcvBal As Long
