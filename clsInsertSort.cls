VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInsertSort"
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
'InsertionSort Module by Adrian "JimCamel" Clark 2002
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
'As I am too lazy to comment this source code, I thought I would
'instead include snippets from my assignment to give a brief
'explaination of the algorithm.

'The Insertion sort method works by sorting the array of data. It
'works by having two loops which run through the array. The first
'loop cycles through the length of the array. The second loop then
'runs through the sorted part of the data, finding the appropriate
'space to put it in.

'This sorting algorithm has been likened to that of someone sorting
'a pack of playing cards. They take the outer most card and place it
'in the sorted area until all the cards are sorted. This is probably the
'most natural type of sorting, and is often used. Unfortunately, it is
'also very slow, running at an average time of O(n2) time, due to
'it’s two loops at one time
End Sub

Private Function InsertSort(arr() As Long, arraySize As Long) As Long()
Dim i As Long, j As Long, temp As Long, k As Long
    For i = 1 To arraySize
        For j = 0 To i - 1
            Comparisonval = Comparisonval + 1
                    If arr(i) < arr(j) Then
                        temp = arr(i)
                        For k = i To j + 1 Step -1
                            arr(k) = arr(k - 1)
                        Next k
                        arr(j) = temp
                    End If
        Next j
    Next i
InsertSort = arr
End Function

Public Function Sort(arr() As Long, arraySize As Long) As Long()
Dim startTime As Long, finishTime As Long

Comparisonval = 0

startTime = GetTickCount
Sort = InsertSort(arr, arraySize)
finishTime = GetTickCount
Timeval = finishTime - startTime

End Function

Public Function TimeTaken() As Long
    TimeTaken = Timeval
End Function

Public Function Comparisons() As Long
    Comparisons = Comparisonval
End Function
