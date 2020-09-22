VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TokenList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"TokenItem"
Attribute VB_Ext_KEY = "Member0" ,"TokenItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'standard interface to list of token items

Private mCol As Collection

Public Function Add(objNewMember As TokenItem) As TokenItem

On Error GoTo Add_Error

    mCol.Add objNewMember
    Set Add = objNewMember

   On Error GoTo 0
   Exit Function

Add_Error:

End Function

Public Property Get Item(vntIndexKey As Variant) As TokenItem
Attribute Item.VB_UserMemId = 0

    On Error GoTo Item_Error

    Set Item = mCol(vntIndexKey)

    On Error GoTo 0
    Exit Property

Item_Error:

End Property

Public Sub Remove(vntIndexKey As Variant)

    On Error GoTo Remove_Error

    mCol.Remove vntIndexKey

    On Error GoTo 0
    Exit Sub

Remove_Error:

End Sub

Public Property Get Count() As Long
    Count = mCol.Count
End Property


Sub Clear()
    Set mCol = New Collection
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

