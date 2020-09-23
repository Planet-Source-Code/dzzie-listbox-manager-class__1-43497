VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Info:    This class is a wrapper to access listbox functions
'         making your code smaller and easier to read.
'
'         it also encapsulates a bunch of new features like
'         being able to filter lists, save to file, load to array
'         load from array etc.
'
'License: You are free to use this library in both private and commercial
'         products at no cost so long as this header stays intact.
'
'Author:  dzzie@yahoo.com
'
'Site:    http://sandsprite.com
'---------------------------------------------------------------

Private lst As ListBox
Attribute lst.VB_VarHelpID = -1

Public noSelectionWarningMessage As String

Property Let ListObject(l As ListBox)
    Set lst = l
End Property

Property Get Count() As Long
  Count = lst.ListCount
End Property
 
Property Get value(Index As Long)
    value = lst.List(Index)
End Property
 
Property Get SelectedText() As String
    SelectedText = lst.List(lst.ListIndex)
End Property

Property Let SelectedText(newTxt As String)
    lst.List(lst.ListIndex) = newTxt
End Property

Property Get SelectedIndex() As Long
    SelectedIndex = lst.ListIndex
End Property

Property Get HasSelection() As Boolean
    If lst.ListIndex >= 0 Then HasSelection = True
End Property

Function EnsureSelection() As Boolean
    If Not HasSelection Then
        MsgBox noSelectionWarningMessage, vbInformation
    Else
        EnsureSelection = True
    End If
End Function

Sub Remove(Index)
    lst.RemoveItem Index
End Sub

Sub Clear()
    lst.Clear
End Sub

Sub AddItem(it)
    lst.AddItem it
End Sub

Sub UpdateValue(newVal, Index)
    lst.List(Index) = newVal
End Sub

Sub FilterList(filt, Optional Likeit As Boolean = True)
    If filt = "*" Or Trim(filt) = Empty Then Exit Sub
    Dim tmp()
    tmp() = GetListToArray
    tmp() = filterArray(tmp(), filt, Likeit)
    LoadArray tmp()
End Sub

Function GetListToArray() As Variant()
    Dim tmp(), i As Long
    For i = 0 To lst.ListCount - 1
        push tmp, lst.List(i)
    Next
    GetListToArray = tmp()
End Function

Function GetListContents(Optional JoinWith = vbCrLf) As String
    Dim tmp As String, i As Long
    For i = 0 To lst.ListCount - 1
        tmp = tmp & lst.List(i) & JoinWith
    Next
    GetListContents = tmp
End Function

Sub LoadFile(fpath As String, Optional delimiter = vbCrLf, Optional AppendIt As Boolean = False)
    Dim tmp
    If FileExists(fpath) Then
        tmp = Split(ReadFile(fpath), delimiter)
        LoadArray tmp, AppendIt
    End If
End Sub

Sub LoadArray(ary, Optional AppendIt As Boolean = False)
    If Not AppendIt Then lst.Clear
    If AryIsEmpty(ary) Then: lst.AddItem "[Empty Set]": Exit Sub
    Dim i As Long
    For i = LBound(ary) To UBound(ary)
        lst.AddItem ary(i)
    Next
End Sub

Sub LoadDelimitedString(dStr, delimiter, Optional AppendIt As Boolean = False)
    Dim tmp
    tmp = Split(dStr, delimiter)
    LoadArray tmp, AppendIt
End Sub

Function filterArray(ary, filtStr, Optional Likeit As Boolean = True) As Variant()
    If AryIsEmpty(ary) Then Exit Function
    
    Dim tmp(), i As Long
    filtStr = filtStr
    'if you use lcase() on somthing not expliticly defined string
    'it returns nothign! wildcard expression always second
    For i = LBound(ary) To UBound(ary)
        If Likeit Then
            If ary(i) Like filtStr Then push tmp, ary(i)
        Else
            If Not ary(i) Like filtStr Then push tmp, ary(i)
        End If
    Next
    
    filterArray = tmp()
End Function

Private Sub Class_Initialize()
    'default message, you can change it at anytime
    noSelectionWarningMessage = "You must first select an item"
End Sub

Sub MatchSize(it As Object)
    lst.Width = it.Width
    lst.Height = it.Height
End Sub


Private Function FileExists(path) As Boolean
  If Len(path) = 0 Then Exit Function
  If Dir(path, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) <> "" Then FileExists = True
End Function

Private Function ReadFile(filename)
  Dim f As Long, temp
  f = FreeFile
  temp = ""
   Open filename For Binary As #f        ' Open file.(can be text or image)
     temp = Input(FileLen(filename), #f) ' Get entire Files data
   Close #f
   ReadFile = temp
End Function

Private Sub WriteFile(path, it)
    Dim f As Long
    f = FreeFile
    Open path For Output As #f
    Print #f, it
    Close f
End Sub

Private Sub AppendFile(path, it)
    Dim f As Long
    f = FreeFile
    Open path For Append As #f
    Print #f, it
    Close f
End Sub

Private Sub push(ary, value) 'this modifies parent ary object
    On Error GoTo init
    Dim x As Long
    x = UBound(ary) '<-throws Error If Not initalized
    ReDim Preserve ary(UBound(ary) + 1)
    ary(UBound(ary)) = value
    Exit Sub
init: ReDim ary(0): ary(0) = value
End Sub

Private Function AryIsEmpty(ary) As Boolean
  On Error GoTo oops
    Dim x As Long
    x = UBound(ary)
    AryIsEmpty = False
  Exit Function
oops: AryIsEmpty = True
End Function

