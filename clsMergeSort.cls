VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMergeSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Comparisonval As Long
Private Timeval As Long
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Sub DISCLAIMER()
'-------------------------------------------------------------
'MergeSort Module by Adrian "JimCamel" Clark 2002
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
'MergeSort, like QuickSort, is a divide and conquer algorithm. But
'instead of sorting as it splits, MergeSort splits the array into single
'elements, and then sorts them as it joins them back together. This
'is known as “Easy Split, Hard Join”.

'As the algorithm joins the numbers together it
'sorts them. The final steps are often the hardest, as joining
'together 2 partially sorted arrays can be quite time consuming. To
'counter this however, people have made modifications to
'MergeSort, such as natural MergeSort, which, instead of the
'merging being in pairs, it takes as many increasing numbers as it
'can and makes an array out of them. This technique often speeds
'up the time slightly.

'Because of the nature of it’s “Divide and conquer” basis, MergeSort,
'like QuickSort, runs on average at O(n log_2 n) time.

End Sub

Private Function MergeSort(arr() As Long, left As Long, right As Long) As Long()
    Dim mid As Long
    If (left < right) Then
            mid = Int((left + right) / 2)
            Call MergeSort(arr, left, mid)
            Call MergeSort(arr, mid + 1, right)
            Call merge(arr, left, mid + 1, right + 1)
    End If
    MergeSort = arr
End Function
           
Private Sub merge(arr() As Long, left As Long, mid As Long, right As Long)
    Dim i As Long, j As Long, k As Long, tempArr() As Long
       ReDim tempArr(right)
       i = left: j = mid: k = left:
       Do While (i < mid And j < right)
           Comparisonval = Comparisonval + 1
           If (arr(i) <= arr(j)) Then
                tempArr(k) = arr(i): i = i + 1
           Else
                tempArr(k) = arr(j): j = j + 1
           End If
           k = k + 1
       Loop
       Do While (i < mid)
           tempArr(k) = arr(i)
           i = i + 1: k = k + 1
       Loop
        Do While (j < right)
            tempArr(k) = arr(j)
            j = j + 1: k = k + 1
        Loop
        For k = left To right - 1
            arr(k) = tempArr(k)
        Next k
End Sub

Public Function Sort(arr() As Long, arraySize As Long) As Long()
Dim startTime As Long, finishTime As Long

Comparisonval = 0

startTime = GetTickCount
Sort = MergeSort(arr, 0, arraySize)
finishTime = GetTickCount
Timeval = finishTime - startTime

End Function

Public Function TimeTaken() As Long
    TimeTaken = Timeval
End Function

Public Function Comparisons() As Long
    Comparisons = Comparisonval
End Function

