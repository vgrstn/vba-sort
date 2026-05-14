Attribute VB_Name = "Sort"
'@IgnoreModule ProcedureNotUsed, UnassignedVariableUsage, UseMeaningfulName, MultipleDeclarations, ParameterCanBeByVal, HungarianNotation, VariableNotAssigned
'@Folder("Module")
'@ModuleDescription "Sorting and searching."

'------------------------------------------------------------------------------
' MIT License
'
' Copyright (c) 2025 Vincent van Geerestein
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.
'------------------------------------------------------------------------------

Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Comments
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Author: Vincent van Geerestein
' E-mail: vincent@vangeerestein.com
' Description: Sorting and Searching Module
' Add-in: RubberDuck (https://rubberduckvba.com/)
' Version: 2025.10.12
'
' This module implements sorting and searching using comparison by operator.
'
' Methods
' SortArray arr [, idx, asc]            Sorts an array
' SearchArray(arr, val [, idx, start])  Searches for a value in a sorted array
' IsArraySorted(arr [, idx])            Returns True if an array is sorted
' ShuffleArray arr                      Shuffles an array
'
' The actual sort order can be reversed by providing an optional parameter.
' The optional idx parameter determines whether a sort or search is done in
' place or by index. The index array is automatically created by the sort
' routine and is returned sorted to the caller.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private declarations
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Threshold for applying InsertionSort (taken from Java).
Private Const INSERTION_SORT_THRESHOLD As Long = 47

' Selected VB errors.
Private Enum VBERROR
    vbErrorInvalidProcedureCall = 5
    vbErrorSubscriptOutOfRange = 9
    vbErrorTypeMismatch = 13
End Enum


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Public methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description "Sorts an array either in place or by index."
Public Sub SortArray( _
    ByRef arr As Variant, _
    Optional ByRef idx As Variant, _
    Optional ByVal asc As Boolean = True _
)

    ' The array must not be empty and must be one dimensional.
    If IsVector(arr) = False Then Err.Raise vbErrorTypeMismatch

    ' Seed the random-number generator.
    Static Seed As Boolean
    If Seed = False Then Randomize: Seed = True

    ' Perform the sort either in place or by index.
    If VBA.IsMissing(idx) Then
        QuickSortInPlace arr
        If asc = False Then ReverseArray arr
    Else
        idx = CreateIndexArray(arr)
        QuickSortByIndex arr, idx
        If asc = False Then ReverseArray idx
    End If

End Sub


'@Description "Returns the position of an element with a specified value in an ordered array."
Public Function SearchArray( _
    ByRef arr As Variant, _
    ByVal value As Variant, _
    Optional ByRef idx As Variant, _
    Optional ByVal Start As Variant _
) As Variant
' Returns Null if the search value is not found.

' Looks for next value at start+1 if start is provided. If the other parameters
' have changed since the original search unpredictable results will be obtained.

    ' The array must not be empty and must be one dimensional.
    If IsVector(arr) = False Then Err.Raise vbErrorTypeMismatch

    ' Perform the search either in place or by index.
    If VBA.IsMissing(idx) Then
        If VBA.IsMissing(Start) Then
            SearchArray = BinarySearchInPlace(arr, value)
        Else
            SearchArray = NextValueInPlace(arr, value, Start)
        End If
    ElseIf IsIndexArray(idx, arr) Then
        If VBA.IsMissing(Start) Then
            SearchArray = BinarySearchByIndex(arr, value, idx)
        Else
            SearchArray = NextValueByIndex(arr, value, idx, Start)
        End If
    Else
        Err.Raise vbErrorInvalidProcedureCall, , "Index is invalid"
    End If

End Function


'@Description "Returns True if an array is sorted or False if not."
Public Function IsArraySorted( _
    ByVal arr As Variant, _
    Optional ByRef idx As Variant _
) As Boolean

    ' The array must not be empty and must be one dimensional.
    If IsVector(arr) = False Then Err.Raise vbErrorTypeMismatch

    ' Perform the check either in place or by index.
    If VBA.IsMissing(idx) Then
        IsArraySorted = IsArraySortedInPlace(arr)
    ElseIf IsIndexArray(idx, arr) Then
        IsArraySorted = IsArraySortedByIndex(arr, idx)
    Else
        Err.Raise vbErrorInvalidProcedureCall, , "Index is invalid"
    End If

End Function


'@Description "Randomizes the order of the elements in an array."
Public Sub ShuffleArray(ByRef arr As Variant)

    ' The array must not be empty and must be one dimensional.
    If IsVector(arr) = False Then Err.Raise vbErrorTypeMismatch

    ' Seed the random-number generator.
    Static Seed As Boolean
    If Seed = False Then Randomize: Seed = True

    ' Random swap the ith array element.
    Dim Lower As Long: Lower = LBound(arr)
    Dim i As Long, index As Long, x As Variant
    For i = UBound(arr) To Lower + 1 Step -1
        index = Lower + Int((i - Lower + 1) * Rnd)
        x = arr(i): arr(i) = arr(index): arr(index) = x
    Next

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description "Sorts an array."
Private Sub QuickSortInPlace( _
    ByRef arr As Variant _
)
' Iterative algorithm after Hardcore Visual Basic version 5.0 (Bruce McKinney).

    ' An empty stack has one item with prevents a extra test for push.
    Dim Stack As Collection: Set Stack = New Collection: Stack.Add Empty

    Dim Lower As Long: Lower = LBound(arr)
    Dim upper As Long: upper = UBound(arr)
    Dim i As Long, j As Long, x As Variant, Pivot As Variant, index As Long
    Do
      Do
        If Lower < upper Then
            ' Swap from ends until lower and upper meet in the middle
            If upper - Lower < INSERTION_SORT_THRESHOLD Then
                InsertionSortInPlace arr, Lower, upper
            Else
                ' Start with random partitioning.
                index = Int((upper - Lower + 1) * Rnd + Lower)
                x = arr(upper): arr(upper) = arr(index): arr(index) = x
                ' Swap high values below the split for low values above.
                i = Lower
                j = upper
                Pivot = arr(upper)

                Do
                    ' Find any low value > split.
                    For i = i To j - 1
                        If Pivot < arr(i) Then Exit For
                    Next
                    ' Find any high value < split.
                    For j = j To i + 1 Step -1
                        If Pivot > arr(j) Then Exit For
                    Next
                    ' Swap too high low value for too low high value.
                    If i < j Then
                        x = arr(i): arr(i) = arr(j): arr(j) = x
                    End If
                Loop While i < j

                If i <> upper Then
                    x = arr(i): arr(i) = arr(upper): arr(upper) = x
                End If

                ' Push range markers of larger part for later sorting.
                If i - Lower < upper - i Then
                    Stack.Add i + 1, before:=1
                    Stack.Add upper, before:=1
                    upper = i - 1
                Else
                    Stack.Add Lower, before:=1
                    Stack.Add i - 1, before:=1
                    Lower = i + 1
                End If

                ' Exit from inner loop to process smaller part.
                Exit Do
            End If
        End If

        ' One element left on the stack which means the stack is empty.
        If Stack.count = 1 Then Exit Sub

        ' Pop range markers for next partition.
        upper = Stack.Item(1): Stack.Remove 1
        Lower = Stack.Item(1): Stack.Remove 1
      Loop
    Loop

End Sub


'@Description "Sorts an array by index."
Private Sub QuickSortByIndex( _
    ByRef arr As Variant, _
    ByRef idx As Variant _
)
' Iterative algorithm after Hardcore Visual Basic version 5.0 (Bruce McKinney).

    ' An empty stack has one item with prevents a extra test for push.
    Dim Stack As Collection: Set Stack = New Collection: Stack.Add Empty

    Dim Lower As Long: Lower = LBound(idx)
    Dim upper As Long: upper = UBound(idx)
    Dim i As Long, j As Long, Pivot As Variant, x As Long, index As Long
    Do
      Do
        If Lower < upper Then
            ' Swap from ends until first and last meet in the middle
            If upper - Lower < INSERTION_SORT_THRESHOLD Then
                InsertionSortByIndex arr, Lower, upper, idx
            Else
                ' Start with random partitioning.
                index = Int((upper - Lower + 1) * Rnd + Lower)
                x = idx(upper): idx(upper) = idx(index): idx(index) = x
                ' Swap high values below the split for low values above.
                i = Lower
                j = upper
                ' Save current pivot value for efficiency.
                Pivot = arr(idx(upper))

                Do
                    ' Find any low value > split.
                    For i = i To j - 1
                        If Pivot < arr(idx(i)) Then Exit For
                    Next
                    ' Find any high value < split.
                    For j = j To i + 1 Step -1
                        If Pivot > arr(idx(j)) Then Exit For
                    Next
                    ' Swap too high low value for too low high value.
                    If i < j Then
                        x = idx(i): idx(i) = idx(j): idx(j) = x
                    End If
                Loop While i < j

                If i <> upper Then
                    x = idx(i): idx(i) = idx(upper): idx(upper) = x
                End If

                ' Push range markers of larger part for later sorting.
                If i - Lower < upper - i Then
                    Stack.Add i + 1, before:=1
                    Stack.Add upper, before:=1
                    upper = i - 1
                Else
                    Stack.Add Lower, before:=1
                    Stack.Add i - 1, before:=1
                    Lower = i + 1
                End If

                ' Exit from inner loop to process smaller part.
                Exit Do
            End If
        End If

        ' One element left on the stack which means the stack is empty.
        If Stack.count = 1 Then Exit Sub

        ' Pop range markers for next partition.
        upper = Stack.Item(1): Stack.Remove 1
        Lower = Stack.Item(1): Stack.Remove 1
      Loop
    Loop

End Sub


'@Description "Sorts a subarray."
Private Sub InsertionSortInPlace( _
    ByRef arr As Variant, _
    ByVal Lower As Long, _
    ByVal upper As Long _
)

    Dim Item As Variant
    Dim i As Long, j As Long
    For i = Lower + 1 To upper
        Item = arr(i)
        For j = i To Lower + 1 Step -1
            If Item > arr(j - 1) Then Exit For
            arr(j) = arr(j - 1)
        Next
        arr(j) = Item
    Next

End Sub


'@Description "Sorts a subarray by index."
Private Sub InsertionSortByIndex( _
    ByRef arr As Variant, _
    ByVal Lower As Long, _
    ByVal upper As Long, _
    ByRef idx As Variant _
)

    Dim Item As Variant, index As Long
    Dim i As Long, j As Long
    For i = Lower + 1 To upper
        index = idx(i)
        Item = arr(index)
        For j = i To Lower + 1 Step -1
            If Item > arr(idx(j - 1)) Then Exit For
            idx(j) = idx(j - 1)
        Next
        idx(j) = index
    Next

End Sub


'@Description "Searches for a value in an ordered array."
Private Function BinarySearchInPlace( _
    ByRef arr As Variant, _
    ByVal value As Variant _
) As Variant
' Returns Null if the value is not found.

    Dim Lower As Long: Lower = LBound(arr)
    Dim upper As Long: upper = UBound(arr)

    ' Determine the order in the array (ascending or descending).
    Dim Order As Long: Order = Compare(arr(Lower), arr(upper))
    Dim middle As Long
    Do
        middle = (Lower + upper) \ 2
        Select Case Compare(value, arr(middle))
        Case 0
            ' Return the lowest index for which arr(i) = value.
            Dim i As Long
            For i = middle - 1 To Lower Step -1
                If Compare(value, arr(i)) <> 0 Then Exit For
            Next
            BinarySearchInPlace = i + 1
            Exit Function
        Case Order
            upper = middle - 1
        Case Else
            Lower = middle + 1
        End Select
    Loop Until Lower > upper

    BinarySearchInPlace = Null

End Function


'@Description "Searches for a value in an array ordered by an index."
Private Function BinarySearchByIndex( _
    ByRef arr As Variant, _
    ByVal value As Variant, _
    ByVal idx As Variant _
) As Variant
' Returns Null if the value is not found.

    Dim Lower As Long: Lower = LBound(idx)
    Dim upper As Long: upper = UBound(idx)

    ' Determine the order in the array (ascending or descending).
    Dim Order As Long: Order = Compare(arr(idx(Lower)), arr(idx(upper)))
    Dim middle As Long
    Do
        middle = (Lower + upper) \ 2
        Select Case Compare(value, arr(idx(middle)))
        Case 0
            ' Return the lowest index for which arr(i) = value.
            Dim i As Long
            For i = middle - 1 To Lower Step -1
                If Compare(value, arr(idx(i))) <> 0 Then Exit For
            Next
            BinarySearchByIndex = i + 1
            Exit Function
        Case Order
            upper = middle - 1
        Case Else
            Lower = middle + 1
        End Select
    Loop Until Lower > upper

    BinarySearchByIndex = Null

End Function


'@Description "Gives the next matching value in a sorted array."
Private Function NextValueInPlace( _
    ByRef arr As Variant, _
    ByVal value As Variant, _
    ByVal Start As Long _
) As Variant

    If Start + 1 <= UBound(arr) Then
        If Compare(value, arr(Start + 1)) = 0 Then
            NextValueInPlace = Start + 1
            Exit Function
        End If
    End If

    NextValueInPlace = Null

End Function


'@Description "Gives the next matching value in a sorted array by index."
Private Function NextValueByIndex( _
    ByRef arr As Variant, _
    ByVal value As Variant, _
    ByRef idx As Variant, _
    ByVal Start As Long _
) As Variant

    If Start + 1 <= UBound(idx) Then
        If Compare(value, arr(idx(Start + 1))) = 0 Then
            NextValueByIndex = Start + 1
            Exit Function
        End If
    End If

    NextValueByIndex = Null

End Function


'@Description "Determines whether an array is sorted or not."
Private Function IsArraySortedInPlace( _
    ByRef arr As Variant _
) As Boolean

    Dim Lower As Long: Lower = LBound(arr)
    Dim upper As Long: upper = UBound(arr)

    ' Determine the order in the array (ascending or descending).
    Dim Order As Long: Order = Compare(arr(Lower), arr(upper))
    Dim i As Long, Current As Variant, Previous As Variant
    If Order = 0 Then
        ' Check whether all the array elements have identical values.
        Current = arr(Lower)
        For i = Lower + 1 To upper
            If arr(i) <> Current Then Exit Function
        Next
    Else
        ' Check the order for all array elements.
        Previous = arr(Lower)
        For i = Lower + 1 To upper
            Current = arr(i)
            If Compare(Current, Previous) = Order Then Exit Function
            Previous = Current
        Next
    End If

    IsArraySortedInPlace = True

End Function


'@Description "Determines whether an array is sorted by an index or not."
Private Function IsArraySortedByIndex( _
    ByRef arr As Variant, _
    ByVal idx As Variant _
) As Boolean

    Dim Lower As Long: Lower = LBound(idx)
    Dim upper As Long: upper = UBound(idx)

    ' Determine the order in the array (ascending or descending).
    Dim Order As Long: Order = Compare(arr(idx(Lower)), arr(idx(upper)))
    Dim i As Long, Current As Variant, Previous As Variant
    If Order = 0 Then
        ' All array elements should have identical values.
        Current = arr(idx(Lower))
        For i = Lower + 1 To upper
            If arr(idx(i)) <> Current Then Exit Function
        Next
    Else
        ' Check the order by index for all array elements.
        Previous = arr(idx(Lower))
        For i = Lower + 1 To upper
            Current = arr(idx(i))
            If Compare(Current, Previous) = Order Then Exit Function
            Previous = Current
        Next
    End If

    IsArraySortedByIndex = True

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Private utility methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'@Description "Reverses the order of the elements in an array."
Private Sub ReverseArray(ByRef arr As Variant)

    Dim offset As Long: offset = UBound(arr) + LBound(arr)
    Dim i As Long, x As Variant
    For i = LBound(arr) To LBound(arr) + (UBound(arr) - LBound(arr)) \ 2
        x = arr(i): arr(i) = arr(offset - i): arr(offset - i) = x
    Next

End Sub


'@Description "Creates an index array."
Private Function CreateIndexArray( _
    ByRef arr As Variant, _
    Optional ByVal base As Long _
) As Long()

    Dim idx() As Long: ReDim idx(base To base + UBound(arr) - LBound(arr))
    Dim offset As Long: offset = LBound(arr) - base
    Dim i As Long
    For i = LBound(idx) To UBound(idx)
        idx(i) = offset + i
    Next

    CreateIndexArray = idx

End Function


'@Description "Returns True if an array is a valid index array or False otherwise."
Private Function IsIndexArray( _
    ByRef idx As Variant, _
    ByRef arr As Variant _
) As Boolean
' Check the integrity of an index array and check for consistency with the indexed array.

    ' The array must not be empty and be one dimensional.
    If IsVector(idx) = False Then Exit Function

    Dim Lower As Long: Lower = LBound(arr)
    Dim upper As Long: upper = UBound(arr)

    ' Check whether the array lengths match.
    If UBound(idx) - LBound(idx) <> upper - Lower Then Exit Function

    ' Check whether all indices are in range and are all used once.
    Dim Used() As Boolean: ReDim Used(Lower To upper)
    Dim i As Long, index As Long
    For i = LBound(idx) To UBound(idx)
        index = idx(i)
        If index < Lower Or index > upper Then Exit Function
        If Used(index) Then Exit Function
        Used(index) = True
    Next

    IsIndexArray = True

End Function


'@Description "Returns True if a variable is an empty array or False otherwise."
Private Function IsArrayEmpty(ByRef var As Variant) As Boolean

    If VBA.IsArray(var) = False Then Exit Function

    On Error Resume Next
    Err.Clear
    Dim Lo As Long: Lo = LBound(var)
    Dim Hi As Long: Hi = UBound(var)
    If Err.Number = 0 Then
        IsArrayEmpty = (Hi < Lo)
    Else
        IsArrayEmpty = True
    End If
    Err.Clear
    On Error GoTo 0

End Function


'@Description "Returns the number of array dimensions of a variable."
Private Function ArrayNDims(ByRef var As Variant) As Long
' An allocated but empty first dimension returns 0.

    If VBA.IsArray(var) = False Then Exit Function

    Dim DimIndex As Long
    Dim Lo As Long, Hi As Long
    On Error GoTo ForcedError
    Do
        DimIndex = DimIndex + 1
        Lo = LBound(var, DimIndex)
        Hi = UBound(var, DimIndex)
    Loop While Lo <= Hi

ForcedError:
    ArrayNDims = DimIndex - 1
    On Error GoTo 0

End Function


'@Description "Returns True if a variable is a vector or False otherwise."
Private Function IsVector(ByRef var As Variant) As Boolean
' A vector is an one-dimensional non-empty array of any type.

    If VBA.IsArray(var) = False Then Exit Function

    Dim Lo As Long, Hi As Long
    On Error Resume Next
    Lo = LBound(var, 2)
    If Err.Number = 0 Then Exit Function
    Err.Clear
    Lo = LBound(var, 1)
    Hi = UBound(var, 1)
    If Err.Number = 0 Then IsVector = (Lo <= Hi)
    On Error GoTo 0

End Function


'@Description "Compares two values."
Private Function Compare(ByVal value1 As Variant, ByVal value2 As Variant) As Long

    Select Case True
    Case value1 < value2
        Compare = -1
    Case value1 = value2
        Compare = 0
    Case Else
        Compare = 1
    End Select

End Function
