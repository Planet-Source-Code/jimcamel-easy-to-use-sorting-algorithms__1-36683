VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sorting Algorithm Comparison"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRadixSort 
      Caption         =   "RadixSort"
      Height          =   255
      Left            =   480
      TabIndex        =   18
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton cmdMergeSort 
      Caption         =   "MergeSort"
      Height          =   255
      Left            =   480
      TabIndex        =   17
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuickSort 
      Caption         =   "QuickSort"
      Height          =   255
      Left            =   480
      TabIndex        =   16
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdHeapSort 
      Caption         =   "HeapSort"
      Height          =   255
      Left            =   480
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox time 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   4
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.TextBox time 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   3
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1680
      Width           =   975
   End
   Begin VB.TextBox time 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   2
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin VB.TextBox time 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      Height          =   285
      Index           =   1
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.CommandButton cmdInsertSort 
         Caption         =   "InsertSort"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Comparisons 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   4
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox Comparisons 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   3
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox Comparisons 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   2
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox Comparisons 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   1
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Comparisons 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox time 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Index           =   0
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Comparisons:"
         Height          =   195
         Left            =   2880
         TabIndex        =   13
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Time Taken:"
         Height          =   195
         Left            =   1800
         TabIndex        =   12
         Top             =   240
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Algorithm"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is just the testing/comparing form to give an idea
'Of the speed and comparisons of each algorithm
'We use 1000 random numbers
Const numberRandom = 1000
Dim random(numberRandom) As Long
'Declare all the algorithms
Dim InsertSort As New clsInsertSort
Dim HeapSort As New clsHeapSort
Dim Quicksort As New clsQuickSort
Dim MergeSort As New clsMergeSort
Dim radixSort As New clsRadixSort

'Note: to retrieve the sorted data from any of these algorithms
'Simply replace "Call" with the array you want the data to be
'Copied into, ie:
'Dim sortedArray() as long
'sortedArray = Heapsort.sort(random, numberRandom

Private Sub cmdHeapSort_Click()
        'Randomise the array
        RandomizeValues random, numberRandom
        'Sort the numbers
        Call HeapSort.Sort(random, numberRandom)
        'Find the values
        time(1).Text = HeapSort.TimeTaken
        Comparisons(1).Text = HeapSort.Comparisons
End Sub

Private Sub cmdInsertSort_Click()
        'Randomise the array
        RandomizeValues random, numberRandom
        'Sort the numbers
        Call InsertSort.Sort(random, numberRandom)
        'Find the values
        time(0).Text = InsertSort.TimeTaken
        Comparisons(0).Text = InsertSort.Comparisons
End Sub

Private Sub cmdMergeSort_Click()
        'Randomise the array
        RandomizeValues random, numberRandom
        'Sort the numbers
        Call MergeSort.Sort(random, numberRandom)
        'Find the values
        time(3).Text = MergeSort.TimeTaken
        Comparisons(3).Text = MergeSort.Comparisons
End Sub

Private Sub cmdQuickSort_Click()
        'Randomise the array
        RandomizeValues random, numberRandom
        'Sort the numbers
        Call Quicksort.Sort(random, numberRandom)
        'Find the values
        time(2).Text = Quicksort.TimeTaken
        Comparisons(2).Text = Quicksort.Comparisons
End Sub

Private Sub cmdRadixSort_Click()
        'Randomise the array
        RandomizeValues random, numberRandom
        'Sort the numbers
        Call radixSort.Sort(random, numberRandom, 1000000)
        'Find the values
        time(4).Text = radixSort.TimeTaken
        Comparisons(4).Text = "N\A"
End Sub

Private Function RandomizeValues(ByRef arr() As Long, size As Long)
'Get a new random seed
Randomize Timer
Dim i As Long
'Fill the array with random numbers
For i = 0 To size
    arr(i) = Rnd * (size * 1000)
Next i
End Function

Private Sub Form_Load()
Dim temp As Long
End Sub
