VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsQuickSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Comparisonval As Long
Private Timeval As Long
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private pivot As Long

Public Sub DISCLAIMER()
'-------------------------------------------------------------
'QuickSort Module by Adrian "JimCamel" Clark 2002
'email:jimcamel@jimcamel.8m.com
'icq:25282667
'msn:adrian_clark@hotmail.com
'-------------------------------------------------------------

'So I had to write a whole bunch of sorting algorithms in java for
'an assignment, so I thought I'd convert them to vb.
'Use these as you wish, just so long as you leave this text intact
'If you use this for anything interesting, feel free to tell me.

End Sub

Public Sub Description()
'QuickSort is a sorting algorithm which falls into the category of
'“Divide and conquer” algorithms. This means that instead of
'solving the problem as one large part, the algorithm instead breaks
'the data into smaller chunks and solves them. By solving all these
'smaller parts, QuickSort can solve the whole problem more
'efficiently.

'QuickSort is known as a “Hard Split, Easy Join” algorithm. This is
'because, as the algorithm is broken into smaller parts, each
'smaller part is partially sorted. Once we reach pieces of size 1, we
'know these are already sorted, and it’s just a matter of putting the
'data back together.

'To spilt the array, we need a partition, which is simply the middle
'number. One half of the array will be less than that number, the
'other half will be more than that number. In my code, we take the
'partition as the leftmost value in the array.

'The processing time for this algorithm varies on it’s choice of the
'pivot. With the implementation I have used, it has a best case of
'O(n log_2 n), and a worse case (when the numbers are in reverse
'order) of O(n2). This can be remedied by making choosing the pivot
'from the median of the first value in the array, the last value in the
'array, and the middle value in the array.
End Sub

Private Function Quicksort(arr() As Long, left As Long, right As Long) As Long()
    Dim mid As Long
        If (left < right) Then
            pivot = arr(left)
            mid = partition(arr, left, right)
            Call Quicksort(arr, left, mid - 1)
            Call Quicksort(arr, mid + 1, right)
        End If
    Quicksort = arr
End Function

Private Function partition(arr() As Long, left As Long, right As Long) As Long
    Dim i As Long, j As Long
    i = left: j = right + 1
        While (i < j)
            Do
                 Comparisonval = Comparisonval + 1
                    j = j - 1
                    If (i = j) Then arr(i) = pivot: partition = i: Exit Function
             Loop While (arr(j) > pivot)
            arr(i) = arr(j)
            Do
                Comparisonval = Comparisonval + 1
                    i = i + 1
                    If (i = j) Then arr(i) = pivot: partition = i: Exit Function
                Loop While (arr(i) < pivot)
            arr(j) = arr(i)
        Wend
        arr(i) = pivot
        partition = i
End Function

Public Function Sort(arr() As Long, arraySize As Long) As Long()
Dim startTime As Long, finishTime As Long

Comparisonval = 0

startTime = GetTickCount
Sort = Quicksort(arr, 0, arraySize)
finishTime = GetTickCount
Timeval = finishTime - startTime

End Function

Public Function TimeTaken() As Long
    TimeTaken = Timeval
End Function

Public Function Comparisons() As Long
    Comparisons = Comparisonval
End Function

