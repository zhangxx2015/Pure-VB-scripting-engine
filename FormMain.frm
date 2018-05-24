VERSION 5.00
Begin VB.Form FormMain 
   Caption         =   "xScript"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "FormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public IScript As Object, IDebug As New ClassDebug

'������,��������
Private ContainerHeap() As ClassEvent, Count_Of_Container As Long


Public Sub SetWindowStyle(ByVal WinStyle As Long)
    Dim lStyle As Long
    Const GWL_STYLE As Long = (-16)
    
    Const WS_MAXIMIZEBOX As Long = &H10000
    Const WS_MAXIMIZE As Long = &H1000000

    'lStyle = GetWindowLong(Me.hwnd, GWL_STYLE)
    
    lStyle = Choose(WinStyle + 1, &H6000000, &H6C80000, &H6CF0000, &H6C80080, &H6C80000, &H6CC0000)
    Call SetWindowLong(Me.hwnd, GWL_STYLE, lStyle)
    
    Const SWP_NOSIZE As Integer = &H1
    Const SWP_NOZORDER As Integer = &H4
    Const SWP_NOMOVE As Integer = &H2
    Const SWP_FRAMECHANGED As Integer = &H20

    Call SetWindowPos(Me.hwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOMOVE Or SWP_FRAMECHANGED)
End Sub


'"VB.PictureBox","VB.Label","VB.TextBox","VB.Frame","VB.CommandButton","VB.CheckBox","VB.OptionButton","VB.ComboBox","VB.ListBox","VB.HScrollBar","VB.VScrollBar","VB.Timer","VB.DriveListBox","VB.DirListBox","VB.FileListBox","VB.Shape","VB.Line","VB.Image","VB.Data","VB.OLE"
Public Function CreateControl(ByVal TypeDef As String) As Object
        
        Dim ClassObject
        '����ʵ��
        Select Case TypeDef
        Case "VB.CommandButton"
            Set ClassObject = New ClassICommandButton
        Case "VB.HScrollBar"
            Set ClassObject = New ClassIHScrollBar
        End Select
        
        '��������
        Set ClassObject.Host = IScript
        '��ӵ�����
        Set ClassObject.Instance = Me.Controls.Add(TypeDef, "hInstance" & GenGuid, Me)
        '����
        Set CreateControl = ClassObject
End Function



Private Function ValueVariable(V) As String
'
End Function



'�¼���
Public Sub EventManage(ByRef Source As ClassEvent, Info As EventInfo)
        '�¼�����ԭ�� �ؼ�����,�¼�����,��������,��������,��������,����ֵ
        'Sub EventHandle(index,name,pCount,pNames(),pType(),pValues())
        


        Dim Codes As String
        Codes = "Call EventHandle({0},{1},{2},{3},{4},{5})"
        Codes = Replace(Codes, "{0}", Source.Index)
        Codes = Replace(Codes, "{1}", Chr(34) & Info.Name & Chr(34))
        
        With Info.EventParameters
                If .Count > 0 Then
                        Dim Names() As Variant, Types() As Variant, Values() As Variant
                        ReDim Names(.Count)
                        ReDim Types(.Count)
                        ReDim Values(.Count)
                        
'On Error Resume Next
                        Dim i As Long
                        For i = 0 To .Count - 1
                                If UCase(.Item(i).Name) <> "POSTDATA" Then
                                        Names(i) = Chr(34) & .Item(i).Name & Chr(34)
                                        Types(i) = Chr(34) & TypeName(.Item(i).Value) & Chr(34)
                                        Values(i) = Chr(34) & .Item(i).Value & Chr(34)
                                End If
                        Next i
                        
                        Codes = Replace(Codes, "{2}", .Count)
                        
                        Dim StringParam As String
                        StringParam = Join(Names, ","): StringParam = Left(StringParam, Len(StringParam) - 1)
                        
                        Codes = Replace(Codes, "{3}", "Array(" & StringParam & ")")
                        
                        StringParam = Join(Types, ","): StringParam = Left(StringParam, Len(StringParam) - 1)
                        Codes = Replace(Codes, "{4}", "Array(" & StringParam & ")")
                        
                        StringParam = Join(Values, ","): StringParam = Left(StringParam, Len(StringParam) - 1)
                        Codes = Replace(Codes, "{5}", "Array(" & StringParam & ")")
                Else
                        Codes = Replace(Codes, "{2}", 0)
                        Codes = Replace(Codes, "{3}", "Array(vbnull)")
                        Codes = Replace(Codes, "{4}", "Array(vbnull)")
                        Codes = Replace(Codes, "{5}", "Array(vbnull)")
                End If
        End With
        
        '�㲥��Ϣ
        Call PostEvent(Codes)
End Sub

'�㲥��Ϣ
Private Sub PostEvent(ByVal ScriptCode As String)
On Error Resume Next
        Call IScript.ExecuteStatement(ScriptCode)
End Sub



    
'�ж��Ƿ�������IDE��
Private Function IsInIDE() As Boolean
        IsInIDE = Not App.LogMode
End Function


Private Sub Form_Load()
        If IsInIDE = False Then MsgBox "Create By QQ:20437023"
        
        '��ʼ��������
        ReDim ContainerHeap(Count_Of_Container)
        '�����ű�ִ����
        Set IScript = CreateObject("MSScriptControl.ScriptControl.1")
        
        '���ýű�����
        Set IDebug.IScript = IScript
        
        With IScript
                .Language = "VBScript"
                
                '���Ϊ�ű�����Ĺؼ���(�ű���������)
                .AddObject "xVB", Me, True
                
                '���Ϊ�ű�����Ĺؼ���(Ӧ�ó������)
                .AddObject "Host", VB.Global.App, True
                
                '���Ϊ�ű�����Ĺؼ���(������Զ���)
                .AddObject "xDebug", IDebug, True
                
                
                '��ȡ�ű��ļ���
                Dim ScriptFile As String
                'If (Command$ = vbNullString Or Dir(Command$) = vbNullString) Then
                        Dim Dialog As Object
                        Set Dialog = CreateObject("MSComDlg.CommonDialog")
                        With Dialog
                                .InitDir = App.Path
                                .Filter = "�ı��ļ�|*.txt"
                                .DialogTitle = "��ѡ��ű��ļ�"
                                .ShowOpen
                                ScriptFile = .FileName
                        End With
                        Set Dialog = Nothing
                        If (ScriptFile = vbNullString) Then Exit Sub
                'Else
                '        ScriptFile = Replace(Command$, Chr(34), vbNullString)
                'End If
                
                '��ȡ�ű��ļ�
                Dim Bytes() As Byte, StringCodes As String
                Open ScriptFile For Binary As #1
                        ReDim Bytes(LOF(1) - 1)
                        Get #1, , Bytes
                Close #1
                StringCodes = StrConv(Bytes, vbUnicode)
                Erase Bytes
                'ȥ������
'                Dim Codes() As String
'                Codes = Split(StringCodes, vbCrLf)
                
On Error GoTo OnErrorHandle
                Dim i As Long
'                For i = LBound(Codes) To UBound(Codes)
'                        '��ӽű�����
'                        .AddCode Codes(i) & vbCrLf
'                Next i

                '���Ƹ���
                Dim StringCodes2 As String
                StringCodes2 = StringCodes
                StringCodes2 = Replace(StringCodes2, "'", vbNullString)
                StringCodes2 = LCase(StringCodes2)
                StringCodes2 = " " & StringCodes2
                
                '���Ĭ����Ϣ�ӿ�
                If (InStr(1, StringCodes2, "sub eventhandle(index,name,pcount,pnames(),ptype(),pvalues())")) = 0 Then
                        StringCodes = "Sub EventHandle(index,name,pCount,pNames(),pType(),pValues())" & vbCrLf & _
                                      "End Sub" & vbCrLf & _
                                      StringCodes
                End If
                
                '��ӵ��ű�����
                .AddCode StringCodes
                
                
                'ִ��������
                Dim Result As Variant
                Result = .ExecuteStatement(Replace("Call Main({0})", "{0}", Chr(34) & Command$ & Chr(34)))
        End With
        Exit Sub
OnErrorHandle:
'        MsgBox "��:" & i & "��     " & Codes(i) & "     " & Err.Description
        MsgBox Err.Description
        Err.Clear
End Sub


''�������ƻ�ȡ����
'Public Function CallObjectByName(ByVal ObjName As String) As Object
'        Dim i As Long
'        For i = UBound(ContainerHeap) - 1 To LBound(ContainerHeap) Step -1
'                Set CallObjectByName = ContainerHeap(i)
'                Exit For
'        Next i
'End Function

'Public Function CallObjectByName(Object As Object, ProcName As String, CallType As VbCallType, Args() As Variant)
'        CallObjectByName = CallByName(Object, ProcName, CallType, Args)
'End Function




'"MSComDlg.CommonDialog","MSComctlLib.TreeCtrl.2","MSComctlLib.ListViewCtrl.2","MSComctlLib.TabStrip.2","MSComctlLib.ImageListCtrl.2","MSComctlLib.ProgCtrl.2","MSComctlLib.Toolbar.2","MSComctlLib.SBarCtrl.2","MSComctlLib.ImageComboCtl.2","MSComctlLib.Slider.2","MSComCtl2.DTPicker.2","MSComCtl2.MonthView.2","MSComCtl2.UpDown.2","MSComCtl2.Animation.2","MSComCtl2.FlatScrollBar.2","ComCtl3.CoolBar","MSWinsock.Winsock.1","Shell.Explorer.2","RICHTEXT.RichtextCtrl.1"
Public Sub CreateActiveX(ByVal Alias As String, ByVal ClassID As String, Optional ByVal Visible As Boolean = True, Optional ByVal Left As Long = 0, Optional ByVal Top As Long = 0, Optional ByVal Width As Long = 100, Optional ByVal Height As Long = 100) 'As Object
On Error Resume Next
        '������֤
        Licenses.Add ClassID
        
        'ʵ������
        Set ContainerHeap(Count_Of_Container) = New ClassEvent
        With ContainerHeap(Count_Of_Container)
                '������Ϣ�йܶ���
                Set .Master = Me
                '����������
                .Index = Count_Of_Container
                
                '����ActiveX�ؼ�
                Set .Container = Me.Controls.Add(ClassID, "xContainer" & Count_Of_Container, Me)
                
                '��ӵ�ȫ�ֱ����ռ�
                IScript.AddObject Alias, .Container, True
                
                With .Container
                        '���ÿؼ�����
                        .Visible = Visible
                        .Left = Left
                        .Top = Top
                        .Width = Width
                        .Height = Height
                        
'                        '���ض���ʵ��
'                        Set CreateActiveX = .Object
                End With
                '�����ۼ�
                Count_Of_Container = Count_Of_Container + 1
        End With
        
        '��̬�����ڴ�
        ReDim Preserve ContainerHeap(Count_Of_Container)
End Sub

'�ṩ��VBS�Ľӿ�,���ڶ�̬���ɶ���
Public Function CreateObjectEx(Class, Optional ServerName As String) As Object
        Set CreateObjectEx = CreateObject(Class, ServerName)
End Function






Private Sub Form_Resize()
        If IScript Is Nothing Then Exit Sub
        '�¼�����ԭ�� �ؼ�����,�¼�����,��������,��������,��������,����ֵ
        'Sub EventHandle(index,name,pCount,pNames(),pType(),pValues())

        Dim Codes As String
        Codes = "Call EventHandle({0},{1},{2},{3},{4},{5})"
        Codes = Replace(Codes, "{0}", "-1")
        Codes = Replace(Codes, "{1}", Chr(34) & "xVB_Resize" & Chr(34))
        Codes = Replace(Codes, "{2}", "0")
        Codes = Replace(Codes, "{3}", "Array(vbnull)")
        Codes = Replace(Codes, "{4}", "Array(vbnull)")
        Codes = Replace(Codes, "{5}", "Array(vbnull)")
        '�㲥��Ϣ
        Call PostEvent(Codes)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'�ͷ���Դ
On Error Resume Next
        Dim i As Long
        For i = UBound(ContainerHeap) - 1 To LBound(ContainerHeap) Step -1
                Set ContainerHeap(i) = Nothing
        Next i
        Set IScript = Nothing
        Set FormMain = Nothing
End Sub






