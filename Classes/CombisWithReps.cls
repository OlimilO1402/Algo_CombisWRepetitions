VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CombisWithReps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_values() As Long
Private m_chars() As String
Private m_N As Long
Private m_k As Long

Public Sub init(ByVal word As String, ByVal k As Long)
    m_N = Len(word)
    m_k = k
    ReDim m_values(0 To k - 1)
    ReDim m_chars(0 To k - 1)
End Sub

Public Function HasNext() As Boolean
    '
End Function

Public Function getNext() As Long()
    '
End Function
