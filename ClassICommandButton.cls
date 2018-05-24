VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassICommandButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Host As Object
Public WithEvents Instance As CommandButton
Attribute Instance.VB_VarHelpID = -1


Private Type tHandle
    EventName As String
    HandleName As String
End Type

Private Handles(255) As tHandle
Private CountHandles As Long

Public Sub AddHandles(ByVal EventName As String, ByVal HandleName As String)
    With Handles(CountHandles)
        .EventName = EventName
        .HandleName = HandleName
    End With
    CountHandles = CountHandles + 1
End Sub

Private Sub DoHandle(ByVal EventName As String, ParamArray Params())
    If Host Is Nothing Then Exit Sub
    Dim i As Long
    For i = 0 To CountHandles
        With Handles(i)
            If .EventName = EventName Then
                
                
                If IsArray(Params) Then
                    Dim CodeTemplate As String
                    CodeTemplate = "Call {0}({1})"
                    CodeTemplate = Replace(CodeTemplate, "{0}", .HandleName)
                    Dim strParamExpress As String
                    strParamExpress = Join(Params, ",")
                    CodeTemplate = Replace(CodeTemplate, "{1}", strParamExpress)
                    Call Host.ExecuteStatement(CodeTemplate)
                Else
                    Call Host.ExecuteStatement(.HandleName)
                End If
                
                Exit Sub
            End If
        End With
    Next i
End Sub


'½Ó¿Ú
Private Sub Instance_Click()
    Call DoHandle("Click")
End Sub

Private Sub Instance_KeyDown(KeyCode As Integer, Shift As Integer)
    Call DoHandle("KeyDown", KeyCode, Shift)
End Sub

Private Sub Instance_KeyPress(KeyAscii As Integer)
    Call DoHandle("KeyPress", KeyAscii)
End Sub

Private Sub Instance_KeyUp(KeyCode As Integer, Shift As Integer)
    Call DoHandle("KeyUp", KeyCode, Shift)
End Sub

Private Sub Instance_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoHandle("MouseDown", Button, Shift, X, Y)
End Sub

Private Sub Instance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoHandle("MouseMove", Button, Shift, X, Y)
End Sub

Private Sub Instance_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call DoHandle("MouseUp", Button, Shift, X, Y)
End Sub
