VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStringBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const STRING_BUFFER_SIZE As Long = 4096

Private m1BufPos As Long
Private m1BufSize As Long
Private msBuf As String


Private Sub Class_Initialize()
    Clear
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    msBuf = vbNullString
End Sub
Public Sub Append(ByVal TheText As Variant)
    On Error GoTo ErrorHandler
    
    Dim BytesToAppend As Long
    Dim TargetSize As Long
    
    If VarType(TheText) = vbNull Then
        TheText = ""
    End If
    
    BytesToAppend = Len(TheText)
    TargetSize = m1BufPos + BytesToAppend
    
    If TargetSize < m1BufSize Then
        Mid$(msBuf, m1BufPos, BytesToAppend) = TheText
        m1BufPos = TargetSize
    Else
        msBuf = msBuf + Space$(STRING_BUFFER_SIZE)
        m1BufSize = m1BufSize + STRING_BUFFER_SIZE
        Append TheText
    End If
    
    Exit Sub
ErrorHandler:
    'HandleError Err.Number, Err.Description, Err.Source, "cStringBuffer_Append"
End Sub
Public Property Get Value() As String
    Value = Left$(msBuf, m1BufPos - 1)
    
End Property
Public Property Let Value(NewValue As String)
    m1BufPos = 1
    Append NewValue
End Property
Public Sub Clear()
   On Error GoTo ErrorHandler
    
   m1BufSize = m1BufSize + STRING_BUFFER_SIZE
   msBuf = Space$(m1BufSize)
   m1BufPos = 1
    
    Exit Sub
ErrorHandler:
    'HandleError Err.Number, Err.Description, Err.Source, "cStringBuffer_Clear"
End Sub
Private Sub HandleError(ErrNumber As Long, sErrDesc As String, _
    sErrSource As String, sOccurredin As String)
    'App.LogEvent "Error in  " + sOccurredin + " Number: " + serrnumber + _
     '   vbCr + "Description: " + sErrDesc + vbCr + " Source:" + serrsourcesOccurredin
    
    
End Sub





