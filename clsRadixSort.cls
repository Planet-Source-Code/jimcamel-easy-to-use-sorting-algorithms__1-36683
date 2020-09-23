VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRadixSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Timeval As Long
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Sub DISCLAIMER()
'-------------------------------------------------------------
'RadixSort Module by Adrian "JimCamel" Clark 2002
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
'Radix sort is the odd algorithm out in this collection. Radixsort does
'not compare numbers to one another. Instead it collects the
'numbers in queues. We must first decide a radix, that is, the base
'the numbers will be in

'We look at the least significant digit first, ie, the ones column. We
'add each number to the corresponding queue, or “bucket” as the
'are called. The we move on to the next digit, and taking the
'numbers from the left of the array.
'The we put them back into the array. We keep doing this
'as many times as there are digits in the number. In this case there
'are only 2 digits, so we are finished sorting. The numbers are
'sorted from top to bottom and left to right.

'For different radixes, we have different sized arrays of “buckets”.
'For example, radix 4 would have buckets 0-3, while radix 8 would
'have buckets 0-7.

'The time of Radixsort can vary, but it is generally accepted to be
'O(n), which is the fastest. However, due to the data structures it is
'kept in, it may be slower than other sorting algorithms depending
'on the construction of it.
End Sub

Private Function radixSort(keys() As Long, arraySize As Long, maxSize As Long, radix As Long) As Long()
    Dim i As Long, j As Long, k As Long, x As Long, length As Long
    Dim count() As Long, bucket() As Long
    ReDim count(radix), bucket(arraySize)
    
    x = 1
            
    For i = 1 To digitLen(maxSize, radix)
            
            For j = 0 To radix - 1
                count(j) = 0
            Next j

            For j = 0 To arraySize
                count((Int(keys(j) / x) Mod radix)) = count((Int(keys(j) / x) Mod radix)) + 1
            Next j
            
            
            For j = radix - 1 To 0 Step -1
                count(j) = 0
                For k = j - 1 To 0 Step -1
                    count(j) = count(j) + count(k)
                Next k
            Next j

            For j = 0 To arraySize
                bucket(count(Int(keys(j) / x) Mod radix)) = keys(j)
                count((Int(keys(j) / x) Mod radix)) = count((Int(keys(j) / x) Mod radix)) + 1
            Next j
            
            For j = 0 To arraySize
                keys(j) = bucket(j)
            Next j
            
            x = x * radix
        Next i
    
    radixSort = keys
End Function

Public Function Sort(arr() As Long, arraySize As Long, maxSize As Long, Optional RadixNumber As Long = 10) As Long()
Dim startTime As Long, finishTime As Long

startTime = GetTickCount
Sort = radixSort(arr, arraySize, maxSize, RadixNumber)
finishTime = GetTickCount
Timeval = finishTime - startTime

End Function

Public Function TimeTaken() As Long
    TimeTaken = Timeval
End Function

Private Function digitLen(strNum As Long, radix As Long) As Long
    On Error GoTo errHand
    Dim iCount As Long
    iCount = 0
    If radix < 2 Then GoTo errHand
    Do
    If Int(strNum / radix ^ iCount) = 0 Then
        digitLen = iCount
        Exit Function
    End If
    iCount = iCount + 1
    Loop
errHand:
    digitLen = 0
End Function
