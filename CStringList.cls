VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStringList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'---------------------------------------------------------------------------------------
' Module    : CStringList
' DateTime  : 7/5/2004 20:14
' Author    : Lewis Miller
' Purpose   : Invisible listbox
'---------------------------------------------------------------------------------------

'this cool class acts like an invisible listbox, except its ten
'times faster and more efficient... uses an array internally


'local variable(s) to hold property value(s)
Private mvarCount As Long    'local copy
Private mvarArr() As String
Private Const INDEX_OUT_OF_BOUNDS As Long = 9

'convert contents to string, seperated by Delimiter (default is space)
Public Function ToString(Optional ByVal Delimiter As String = " ") As String
    ' Author = Lewis Miller

    Dim lngLenNeeded    As Long
    Dim lngPosition     As Long
    Dim lngDelimLen     As Long
    Dim X               As Long

    If mvarCount > 1 Then
        lngDelimLen = Len(Delimiter)
        'count the length of string we need
        For X = 0 To mvarCount - 1
            lngLenNeeded = lngLenNeeded + Len(mvarArr(X))
        Next X

        'make string buffer
        ToString = Space$(lngLenNeeded + (lngDelimLen * (mvarCount - 1)))

        'insert items into string buffer
        lngPosition = 1
        For X = 0 To mvarCount - 1
            lngLenNeeded = Len(mvarArr(X))
            Mid(ToString, lngPosition, lngLenNeeded) = mvarArr(X)
            lngPosition = lngPosition + lngLenNeeded
            If (X < (mvarCount - 1)) And (lngDelimLen > 0) Then
                Mid(ToString, lngPosition, lngDelimLen) = Delimiter
                lngPosition = lngPosition + lngDelimLen
            End If
        Next X
    Else
        'we dont delimit only one item
        If mvarCount = 1 Then
            ToString = mvarArr(0)
        End If
    End If


End Function

'clear items
Sub Clear()
    ReDim mvarArr(0) As String
    mvarCount = 0
End Sub

'add an item to list
Public Function AddItem(ByVal strItem As String, Optional ByVal Index As Long = -1) As String
    ' Author = Lewis Miller

    Dim X As Long

    If Index > mvarCount Or Index < -1 Then

        'note: if your code stops here, you better add some error handling
        ' or check your index :)
        Debug.Assert False
        Err.Raise INDEX_OUT_OF_BOUNDS

    Else
        ReDim Preserve mvarArr(mvarCount) As String

        If Index = -1 Or Index = mvarCount Then
            Index = mvarCount
        Else
            X = mvarCount
            While X > Index
                mvarArr(X) = mvarArr(X - 1)
                X = X - 1
            Wend
        End If
        mvarArr(Index) = strItem
        mvarCount = mvarCount + 1
        AddItem = strItem
    End If

End Function

'remove an item from list
Public Sub RemoveItem(ByVal Index As Long)
    ' Author = Lewis Miller

    Dim X As Long

    If Index >= mvarCount Or Index < 0 Then

        'note: if your code stops here, you better add some error handling
        ' or check your index :)
        Debug.Assert False
        Err.Raise INDEX_OUT_OF_BOUNDS

    Else
        If Index < mvarCount - 1 Then
            X = Index
            While X < mvarCount - 1
                mvarArr(X) = mvarArr(X + 1)
                X = X + 1
            Wend
        End If
        mvarCount = mvarCount - 1
        ReDim Preserve mvarArr(mvarCount - 1) As String
    End If

End Sub

'(read only) number of items in list
Public Property Get ListCount() As Long
    ListCount = mvarCount
End Property

'get an item from list
Public Property Get List(ByVal Index As Long) As String

    If Index >= mvarCount Or Index < 0 Then

        'note: if your code stops here, you better add some error handling
        ' or check your index :)
        Debug.Assert False
        Err.Raise INDEX_OUT_OF_BOUNDS

    Else
        List = mvarArr(Index)
    End If

End Property

'change an item in list
Public Property Let List(ByVal Index As Long, ByVal strItem As String)

    If Index >= mvarCount Or Index < 0 Then

        'note: if your code errors here, you better add some error handling
        ' or check your index :)
        Debug.Assert False
        Err.Raise INDEX_OUT_OF_BOUNDS

    Else
        mvarArr(Index) = strItem
    End If

End Property


Private Sub Class_Initialize()
    Clear
End Sub

Private Sub Class_Terminate()
    Erase mvarArr
End Sub
