VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RcvAdjOrderQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Will Be Used In The frmVendRcvItems Form...
'To Adjust RcvdQty To Specific Orders...
Public OrderNo As String
Public ItemCode As String
Public lQty As Long
Public lWastage As Long
Public lOrdQty As Long
