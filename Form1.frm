VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Removing duplicates from an unsorted string array"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   4470
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label1 
      Caption         =   "Copyright © 2006 by Olof Larsson (kalebeck@hotmail.com)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long 'used to measure the speed, not essential for the code to work

Private Sub Form_Load()

    '------------------------------------
    '
    '   This project is made by Olof Larsson
    '   © 2006, kalebeck@hotmail.com
    '
    '   This code contains two different functions which both do the same thing,
    '   but in two very different ways. It more or less demonstrates the power
    '   of using a collection to detect and remove duplicates from an unsorted
    '   string array. But regardless of that fact I decided to include both of the
    '   algorithms, the first one which uses a collection to remove duplicates
    '   from the array. And the second one which uses only the original array and
    '   not any additional memory. The major difference between these two algorithms
    '   doesn't become obvious until you process an array with tens of thousand of
    '   entries. You can edit these functions to support any type of array that
    '   you see fit.
    '
    '   Enjoy!
    '
    '------------------------------------
    
    
    '-----------------------------
    ' The following code demonstrates on how to create your array and populate it using a file
    ' You can of course use any other normal way to populate the array with strings
    '
    ' This code will load a file with 182,193 entries and remove the duplicates, it will
    ' also measure how fast this process is completed on your computer
    '-----------------------------
    
    
    '-----------------------------
    ' Opens the file muff.txt in the application directory and uses it
    ' to populate the array to demonstrate how the code works

    Dim ho() As String, g As Long, tim As Long
    ReDim ho(0) As String
    
    Open App.Path & "\muff.txt" For Input As #1
    Dim a As String, total As Long
    Do Until EOF(1)
        Line Input #1, a
        If a <> vbNullString Then
            If g >= UBound(ho) Then
                ReDim Preserve ho(UBound(ho) + 20000) As String
            End If
            total = total + 1
            ho(g) = a
            g = g + 1
        End If
    Loop
    Close #1
    ReDim Preserve ho(total) As String
    
    g = GetTickCount 'measures the speed of the process
    
    '-------------------------
    ' This is where the function is called that does the trick
    '-------------------------
    ndiusCOL ho 'Uses a collection to remove duplicates
    '-------------------------
    '-------------------------
    'ndius ho 'Uses only the original array to remove duplicates, inferior in speed
    '-------------------------
   
    Text1.Text = total - UBound(ho) - 1 & " duplicates removed in " & Round((GetTickCount - g) / 1000, 10) & " seconds" & vbCrLf & vbCrLf & "Items left in the array: " & (UBound(ho) + 1) & vbCrLf & "Original size: " & total
    
    '-----------------------------
    ' Prints the contents of your array to the file output.txt in the
    ' application directory, after the duplicates have been removed
    
    Open App.Path & "\output.txt" For Output As #1
    For g = 0 To UBound(ho)
        Print #1, ho(g)
    Next g
    Close #1
   

End Sub

Private Function ndiusCOL(ByRef arr() As String) As String

    'pros:    very high processing speed regardless of arraysize
    '
    'cons:    may use a lot of memory to check large arrays with hundreds of thousands of entries
    '         or more where there are few duplicates.
    '
    'comment: i recommend you to use this function if you do not have very specific needs in
    '         memory management.
    
    Dim h As Collection
    Set h = New Collection
    
    Dim g As Long
    Dim arru As Long
    Dim remcount As Long
    Dim colcount As Long
    
    arru = UBound(arr)
    
    For g = 0 To arru
        On Error Resume Next
        h.Add arr(g), arr(g)
        arr(g) = vbNullString
        If Err.Number <> 0 Then remcount = remcount + 1
    Next g
    
    g = 0
    colcount = h.Count
    Do Until colcount = 0
        arr(g) = h.Item(1)
        h.Remove (1)
        colcount = colcount - 1
        g = g + 1
    Loop
    
    ReDim Preserve arr(arru - remcount) As String
    
End Function

Private Function ndius(ByRef arr() As String) As String

    'pros:    works very good on small arrays with a few thousand items or less, doesn't take any
    '         additional memory to process and remove duplicates from an array
    '
    'cons:    will take forever if you have a large array to check
    '
    'comment: you may find other functions that seemingly do the same thing with a shorter code, but
    '         when you examine them closely you will find that these functions 99% of the time moves
    '         the entries to a new array thereby using up unnecessary memory. this functions only
    '         uses the original array. it assumes that you have a string array and that you do not
    '         have or want to keep any pure vbNullString entries in your array. you can change this
    '         to a different character or a string if you so wish.
    
    Dim arru As Long
    Dim arrl As Long
    Dim g As Long
    Dim g2 As Long
    Dim remcount As Long
    Dim stepback As Boolean
    
    arru = UBound(arr)
    arrl = LBound(arr)
    
    For g = arrl To arru
        For g2 = (g + 1) To arru
            If arr(g) = arr(g2) Then
                arr(g2) = vbNullString
                remcount = remcount + 1
            End If
        Next g2
    Next g
    remcount = 0
    
    For g = arrl To arru
        If g + remcount > arru Then Exit For
        If stepback = True Then g = g - 1: stepback = False
        If arr(g) = vbNullString Then
            remcount = remcount + 1
            For g2 = g To arru - 1
                arr(g2) = arr(g2 + 1)
                If arr(g2 + 1) = vbNullString Then stepback = True
                arr(g2 + 1) = vbNullString
            Next g2
        End If
    Next g
    
    ReDim Preserve arr(arru - remcount) As String
    
End Function
