Attribute VB_Name = "ModuleCommon"
Option Explicit

Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As guid) As Long
Private Type guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
 
Public Function GenGuid() As String
        Dim udtGUID As guid
        If (CoCreateGuid(udtGUID) = 0) Then
                With udtGUID
                        GenGuid = String(8 - Len(Hex$(.Data1)), "0") & Hex$(.Data1) & _
                                  String(4 - Len(Hex$(.Data2)), "0") & Hex$(.Data2) & _
                                  String(4 - Len(Hex$(.Data3)), "0") & Hex$(.Data3) & _
                                IIf((.Data4(0) < &H10), "0", "") & Hex$(.Data4(0)) & _
                                IIf((.Data4(1) < &H10), "0", "") & Hex$(.Data4(1)) & _
                                IIf((.Data4(2) < &H10), "0", "") & Hex$(.Data4(2)) & _
                                IIf((.Data4(3) < &H10), "0", "") & Hex$(.Data4(3)) & _
                                IIf((.Data4(4) < &H10), "0", "") & Hex$(.Data4(4)) & _
                                IIf((.Data4(5) < &H10), "0", "") & Hex$(.Data4(5)) & _
                                IIf((.Data4(6) < &H10), "0", "") & Hex$(.Data4(6)) & _
                                IIf((.Data4(7) < &H10), "0", "") & Hex$(.Data4(7))
                End With
        End If
End Function

