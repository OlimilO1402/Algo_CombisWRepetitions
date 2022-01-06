VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   6990
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      Height          =   6255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   600
      Width           =   7095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Text1.Text = GetCombinationsWReputations("Andre")
End Sub

Private Sub Form_Load()
    Text1.Font = "Courier New"
'    Dim n As Byte: n = 5
'    Dim k As Byte: k = 5
'    ReDim values(0 To k - 1) As Byte
'    Dim getNext As Boolean
'    ReDim chars(0 To k - 1) As String
'    Dim i As Long
'    chars(i) = "A": i = i + 1
'    chars(i) = "n": i = i + 1
'    chars(i) = "d": i = i + 1
'    chars(i) = "r": i = i + 1
'    chars(i) = "e"
'    getNext = True
'    While getNext
'        Debug.Print getString(chars, values)
'        getNext = CombiWReps_getNext(values, n, k)
'    Wend
End Sub
Function getString(chars() As String, v() As Long) As String
    Dim s As String
    Dim i As Long
    For i = 0 To UBound(chars)
        s = s & chars(v(i))
    Next
    getString = s
End Function
Function CombiWReps_getNext(values() As Byte, ByVal n As Byte, ByVal k As Byte) As Boolean
    ' das rechteste Element finden ...
    Dim j As Long: j = k - 1
    Do
        If (j < 0) Then Exit Do
        If values(j) <> n - 1 Then Exit Do
        j = j - 1
    Loop
    If j < 0 Then Exit Function
    ' ... und vergrößern
    values(j) = values(j) + 1
    ' alle Elemente rechts daneben entsprechend setzen set right-hand elements
    j = j + 1
    For j = j To k - 1
        values(j) = values(j - 1)
    Next
    CombiWReps_getNext = True
End Function

Function GetCombinationsWReputations(ByVal word As String) As String
    Dim s As String
    Dim n As Long: n = Len(word)
    ReDim values(0 To n - 1) As Long
    ReDim chars(0 To n - 1) As String
    Dim i As Long
    For i = 1 To n
        chars(i - 1) = Mid$(word, i, 1)
    Next
    Do
        s = s & getString(chars, values) & vbCrLf
        Dim j As Long: j = n - 1
        Do
            If (j < 0) Then Exit Do
            If values(j) <> n - 1 Then Exit Do
            j = j - 1
        Loop
        If j < 0 Then Exit Do
        values(j) = values(j) + 1
        j = j + 1
        For j = j To n - 1
            values(j) = values(j - 1)
        Next
    Loop
    GetCombinationsWReputations = s
End Function
