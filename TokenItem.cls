VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TokenItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"Collection1"
Option Explicit

'interface to a token item

Private mvarValue As String
Private mvarPosition As Long
Private mvarKind As TOKEN_TYPE

Public Property Let Kind(vData As TOKEN_TYPE)
    mvarKind = vData
End Property

Public Property Get Kind() As TOKEN_TYPE
    Kind = mvarKind
End Property

Public Property Get Length() As Long
    Length = Len(mvarValue)
End Property

Public Property Let Value(ByVal vData As String)
    mvarValue = vData
End Property

Public Property Get Value() As String
Attribute Value.VB_UserMemId = 0
    Value = mvarValue
End Property



