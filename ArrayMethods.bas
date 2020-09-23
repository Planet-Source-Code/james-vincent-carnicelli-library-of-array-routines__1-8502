Attribute VB_Name = "ArrayMethods"
'################################################################
' Library of Array Handler Methods
' Created 31 May 2000 by James Vincent Carnicelli
'   http://alexandria.nu/user/jcarnicelli/
'
' Notes:
' These routines work best when the arrays you want to work with
' are Variant values with single-dimension arrays in them.  To
' create a blank one, try the following:
'
' Dim MyList As Variant
' MyList = Array()
'################################################################

Option Explicit


'######################## Public Routines #######################

'Copy the first array's contents to the second.  Allows a subset
'of the array to be specified with a start point and length.
'Please note that if there are no items to copy, the array may
'still end up with one item, since not all arrays can have zero
'items.  A Length value less than zero is the same as not
'specifying a length: all items after StartAt will be copied.
Public Sub ArrayCopy(List, ListToCopyFrom, Optional ByVal StartAt As Integer = 0, Optional ByVal Length As Integer = -1)
    Dim i As Integer
    
    'Clean up the StartAt and Length values to reflect the real boundaries
    If StartAt > UBound(ListToCopyFrom) Then StartAt = UBound(ListToCopyFrom) + 1
    If StartAt < 0 Then StartAt = 0
    If Length + StartAt > UBound(ListToCopyFrom) + 1 Then Length = UBound(ListToCopyFrom) - StartAt + 1
    If Length < 0 Then Length = UBound(ListToCopyFrom) - StartAt + 1
    
    'Resize the destination array
    If Length = 0 Then
        ReDim List(0)
        ArrayCrunch List
    Else
        ReDim List(Length - 1)
        For i = 0 To Length - 1
            ArraySetItem List, i, ListToCopyFrom(i + StartAt)
        Next
    End If
End Sub

'Add a new item to the end of the array.
Public Sub ArrayAppend(List, Item)
    ReDim Preserve List(UBound(List) + 1)
    ArraySetItem List, UBound(List), Item
End Sub

'Remove an item from the array and close the gap with whatever
'was above it.  Note that while a Variant with an array can
'have zero items, other kinds of array require at least one item.
'This means that if you pass a one-item array that can't be
'shrunk down to zero, the last item will be left alone.
Public Sub ArrayRemove(List, ByVal Index As Integer)
    Dim i As Integer
    For i = Index + 1 To UBound(List)
        ArraySetItem List, i - 1, List(i)
    Next
    ArrayCrunch List
End Sub

'Create a gap in the array before the specified index and insert
'the item into it.
Public Sub ArrayInsert(List, ByVal Index As Integer, Item)
    Dim i As Integer
    If Index > UBound(List) Then
        ReDim Preserve List(Index)
    Else
        ReDim Preserve List(UBound(List) + 1)
        For i = UBound(List) To Index + 1 Step -1
            ArraySetItem List, i, List(i - 1)
        Next
    End If
    ArraySetItem List, Index, Item
End Sub

'Create a string representation of the array.  If you don't specify a
'separator, <CR><LF>, the Windows standard new-line combination, is
'assumed.  The following values have special representations:
' - Null        -> {Null}
' - Empty       -> {Empty}
' - Object Ref  -> {Object:<type>} (e.g., {Object:Collection})
'This routine is great for debugging apps that use arrays.
Public Function ArrayToString(List, Optional ByVal Separator As String = vbCrLf) As String
    Dim i As Integer
    For i = 0 To UBound(List)
        If i > 0 Then ArrayToString = ArrayToString & Separator
        If IsObject(List(i)) Then
            ArrayToString = ArrayToString & "{Object:" & TypeName(List(i)) & "}"
        ElseIf IsNull(List(i)) Then
            ArrayToString = ArrayToString & "{Null}"
        ElseIf IsEmpty(List(i)) Then
            ArrayToString = ArrayToString & "{Empty}"
        Else
            ArrayToString = ArrayToString & List(i)
        End If
    Next
End Function

'Concatenate (copy) one list's contents to the end of another.
Public Sub ArrayConcatenate(List, ListToCopyFrom)
    Dim i As Integer, OriginalSize As Integer
    If UBound(ListToCopyFrom) < 0 Then Exit Sub
    OriginalSize = UBound(List) + 1
    ReDim Preserve List(UBound(List) + UBound(ListToCopyFrom) + 1)
    For i = 0 To UBound(ListToCopyFrom)
        ArraySetItem List, i + OriginalSize, ListToCopyFrom(i)
    Next
End Sub

'Trim items from the beginning and/or end of the array.
'For example:
'    ArrayTrim [ A | B | C | D | E ], 2, 1
'yeilds: [ C | D ]
Public Sub ArrayTrim(List, Optional ByVal TrimFromBeginning As Integer = 0, Optional ByVal TrimFromEnd As Integer = 0)
    Dim i As Integer, LastToCopy As Integer
    
    'Resolve bad values
    If TrimFromBeginning < 0 Then TrimFromBeginning = 0
    If TrimFromEnd < 0 Then TrimFromEnd = 0
    
    'Which trimming technique will be optimal?
    If TrimFromBeginning + TrimFromEnd > UBound(List) Then  'Nothing left
        ReDim List(0)
        ArrayCrunch List
        
    ElseIf TrimFromBeginning = 0 Then  'No need to shift stuff
        ReDim Preserve List(UBound(List) - TrimFromEnd)
        
    Else  'Need to shift stuff to the left
        LastToCopy = UBound(List) - TrimFromEnd - TrimFromBeginning
        For i = 0 To LastToCopy
            ArraySetItem List, i, List(i + TrimFromBeginning)
        Next
        ReDim Preserve List(LastToCopy)
    End If
End Sub

'Cut a section out of the list and (optionally) insert the contents
'of another list.  For example:
'    ArraySplice [ A | B | C | D | E ], 2, 2, [ x | y | z ]
'yeilds: [ A | B | x | y | z | E ]
Public Sub ArraySplice(List, ByVal StartAt As Long, ByVal Length As Long, Optional ListToInsert)
    Dim i As Integer, LengthOfTop As Integer, LengthOfListToCopy As Integer
    
    'Prepare for splice operation
    If IsMissing(ListToInsert) Then ListToInsert = Array()
    If StartAt < 0 Then
        Length = Length + StartAt
        StartAt = 0
    End If
    If StartAt > UBound(List) Then StartAt = UBound(List) + 1
    If Length + StartAt > UBound(List) + 1 Then Length = UBound(List) - StartAt + 1
    If Length < 0 Then Length = 0
    LengthOfTop = UBound(List) - Length - StartAt + 1
    LengthOfListToCopy = UBound(ListToInsert) + 1
    
    'Short-cut this if there's nothing to do or the result will be empty
    If Length = 0 And LengthOfListToCopy = 0 Then Exit Sub
    If UBound(List) - Length + UBound(ListToInsert) + 1 < 0 Then
        ReDim List(0)
        ArrayCrunch List
        Exit Sub
    End If
    
    If LengthOfListToCopy > Length Then  'Shove upper part upward
        ReDim Preserve List(UBound(List) - Length + LengthOfListToCopy)
        For i = LengthOfTop - 1 To 0 Step -1
            ArraySetItem List, i + StartAt + LengthOfListToCopy, List(i + StartAt + Length)
        Next
    ElseIf LengthOfListToCopy < Length Then  'Shove upper part downward
        For i = 0 To LengthOfTop - 1
            ArraySetItem List, i + StartAt + LengthOfListToCopy, List(i + StartAt + Length)
        Next
        ReDim Preserve List(UBound(List) - Length + LengthOfListToCopy)
    Else  'Don't need to move anything; we're wiping out the same number
          'of items as we're inserting
    End If
    
    'Insert stuff to copy in
    For i = 0 To UBound(ListToInsert)
        ArraySetItem List, i + StartAt, ListToInsert(i)
    Next
End Sub


'################### Private Support Routines ###################

'Copies the value or attaches the object to the array at the index
'specified.  This is a little more sophisticated than
'MyArray(i) = X, because object references must be copied using
'Set, not Let ("A = 1" is the same as "Let A = 1").
Private Sub ArraySetItem(List, ByVal Index As Integer, Item)
    If IsObject(Item) Then
        Set List(Index) = Item
    Else
        List(Index) = Item
    End If
End Sub

Private Sub ArrayCrunch(List)
    If UBound(List) = 0 Then
        On Error Resume Next
        List = Array()
    Else
        ReDim Preserve List(UBound(List) - 1)
    End If
End Sub
