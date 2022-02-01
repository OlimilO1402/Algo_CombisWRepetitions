VERSION 5.00
Begin VB.Form FMain 
   Caption         =   "FMain"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   6990
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   6990
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   15
      TabIndex        =   2
      Top             =   30
      Width           =   3585
   End
   Begin VB.TextBox Text2 
      Height          =   6375
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   1
      Top             =   480
      Width           =   6975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create Combinations with Repetitions"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Text1.Font = "Courier New"
    Text2.Font = "Courier New"
    Text1.Text = "Andre"
End Sub

Private Sub Form_Resize()
    Dim L As Single: L = 0
    Dim T As Single: T = Text2.Top
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then Text2.Move L, T, W, H
End Sub

Private Sub Command1_Click()
    Text2.Text = GetCombinationsWRepetitions(Text1.Text)
End Sub

Function GetCombinationsWRepetitions(ByVal word As String) As String
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
    GetCombinationsWRepetitions = s
End Function

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
