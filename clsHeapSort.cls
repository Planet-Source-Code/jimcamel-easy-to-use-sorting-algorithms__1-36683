VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHeapSort"
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
'HeapSort Module by Adrian "JimCamel" Clark 2002
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
'HeapSort is one of the more difficult algorithms to understand (at
'least it is for me). It involves using a heap to sort the data, and
'because of this is different to all the algorithms we’ve encountered
'so far. A heap is like a tree, but each child of a node must be larger
'than the parent

'To use a heap to sort data, we need a method to organise the data
'in and ordered way. This method is called “siftup”, and it does just
'that. It sifts smaller numbers from the bottom to the top of the
'tree.

'When we have our unsorted array, we first form a heap out of it.
'This is done by calling siftup on all the nodes from the parent of
'the last child to the root. Then we simply continuously get the
'value of the root and replace it’s value with that of the last child,
'and then call sift up. Doing this will always leave the heap with the
'smallest value at the top. Once the heap is size zero, we will have
'got all the numbers from smallest to largest out.

'We can see that due to it’s binary nature, the computing time for
'HeapSort is O(n log_2 n)
End Sub

Private Function siftup(ByRef arr() As Long, begin As Long, size As Long)
    Dim y As Long, j As Long, k As Long, z As Long
        y = arr(begin): j = begin: k = 2 * begin
        Do While (k < size)
            z = arr(k)
            If (k < size) Then
                Comparisonval = Comparisonval + 1
                If z > arr(k + 1) Then k = k + 1: z = arr(k)
            End If
            If y <= z Then GoTo 1
            arr(j) = z: j = k: k = 2 * j
        Loop
1         arr(j) = y
End Function
   
Private Function deleteMin(ByRef arr() As Long, ByRef treeSize As Long) As Long
    treeSize = treeSize - 1
        deleteMin = arr(1)
        Comparisonval = Comparisonval + 1
            If arr(1) <= arr(treeSize + 1) Then
                arr(1) = arr(treeSize + 1)
                Call siftup(arr, 1, treeSize)
            End If
End Function

Private Function HeapSort(arr() As Long, arraySize As Long) As Long()
    Dim tempArr() As Long, treeSize As Long, j As Long, i As Long
    ReDim tempArr(arraySize + 1)
    treeSize = arraySize
        
        For j = 0 To arraySize
            tempArr(j + 1) = arr(j)
        Next j
        
        arraySize = arraySize + 1
        
        For j = arraySize / 2 To 0 Step -1
        Call siftup(tempArr, j, arraySize)
        Next j
       
        For j = 1 To arraySize
            arr(j - 1) = deleteMin(tempArr, treeSize)
        Next j
        HeapSort = arr
End Function

Public Function Sort(arr() As Long, arraySize As Long) As Long()
Dim startTime As Long, finishTime As Long

Comparisonval = 0

startTime = GetTickCount
Sort = HeapSort(arr, arraySize)
finishTime = GetTickCount
Timeval = finishTime - startTime

End Function

Public Function TimeTaken() As Long
    TimeTaken = Timeval
End Function

Public Function Comparisons() As Long
    Comparisons = Comparisonval
End Function

