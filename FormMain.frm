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
   StartUpPosition =   1  '所有者中心
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

'容器堆,容器索引
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
        '创建实例
        Select Case TypeDef
        Case "VB.CommandButton"
            Set ClassObject = New ClassICommandButton
        Case "VB.HScrollBar"
            Set ClassObject = New ClassIHScrollBar
        End Select
        
        '设置宿主
        Set ClassObject.Host = IScript
        '添加到窗体
        Set ClassObject.Instance = Me.Controls.Add(TypeDef, "hInstance" & GenGuid, Me)
        '返回
        Set CreateControl = ClassObject
End Function



Private Function ValueVariable(V) As String
'
End Function



'事件堆
Public Sub EventManage(ByRef Source As ClassEvent, Info As EventInfo)
        '事件函数原型 控件索引,事件名称,参数个数,参数名称,参数类型,参数值
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
        
        '广播消息
        Call PostEvent(Codes)
End Sub

'广播消息
Private Sub PostEvent(ByVal ScriptCode As String)
On Error Resume Next
        Call IScript.ExecuteStatement(ScriptCode)
End Sub



    
'判断是否运行在IDE中
Private Function IsInIDE() As Boolean
        IsInIDE = Not App.LogMode
End Function


Private Sub Form_Load()
        If IsInIDE = False Then MsgBox "Create By QQ:20437023"
        
        '初始化容器堆
        ReDim ContainerHeap(Count_Of_Container)
        '创建脚本执行器
        Set IScript = CreateObject("MSScriptControl.ScriptControl.1")
        
        '设置脚本宿主
        Set IDebug.IScript = IScript
        
        With IScript
                .Language = "VBScript"
                
                '添加为脚本引擎的关键字(脚本宿主对象)
                .AddObject "xVB", Me, True
                
                '添加为脚本引擎的关键字(应用程序对象)
                .AddObject "Host", VB.Global.App, True
                
                '添加为脚本引擎的关键字(程序调试对象)
                .AddObject "xDebug", IDebug, True
                
                
                '获取脚本文件名
                Dim ScriptFile As String
                'If (Command$ = vbNullString Or Dir(Command$) = vbNullString) Then
                        Dim Dialog As Object
                        Set Dialog = CreateObject("MSComDlg.CommonDialog")
                        With Dialog
                                .InitDir = App.Path
                                .Filter = "文本文件|*.txt"
                                .DialogTitle = "请选择脚本文件"
                                .ShowOpen
                                ScriptFile = .FileName
                        End With
                        Set Dialog = Nothing
                        If (ScriptFile = vbNullString) Then Exit Sub
                'Else
                '        ScriptFile = Replace(Command$, Chr(34), vbNullString)
                'End If
                
                '读取脚本文件
                Dim Bytes() As Byte, StringCodes As String
                Open ScriptFile For Binary As #1
                        ReDim Bytes(LOF(1) - 1)
                        Get #1, , Bytes
                Close #1
                StringCodes = StrConv(Bytes, vbUnicode)
                Erase Bytes
                '去除换行
'                Dim Codes() As String
'                Codes = Split(StringCodes, vbCrLf)
                
On Error GoTo OnErrorHandle
                Dim i As Long
'                For i = LBound(Codes) To UBound(Codes)
'                        '添加脚本代码
'                        .AddCode Codes(i) & vbCrLf
'                Next i

                '复制副本
                Dim StringCodes2 As String
                StringCodes2 = StringCodes
                StringCodes2 = Replace(StringCodes2, "'", vbNullString)
                StringCodes2 = LCase(StringCodes2)
                StringCodes2 = " " & StringCodes2
                
                '添加默认消息接口
                If (InStr(1, StringCodes2, "sub eventhandle(index,name,pcount,pnames(),ptype(),pvalues())")) = 0 Then
                        StringCodes = "Sub EventHandle(index,name,pCount,pNames(),pType(),pValues())" & vbCrLf & _
                                      "End Sub" & vbCrLf & _
                                      StringCodes
                End If
                
                '添加到脚本引擎
                .AddCode StringCodes
                
                
                '执行主函数
                Dim Result As Variant
                Result = .ExecuteStatement(Replace("Call Main({0})", "{0}", Chr(34) & Command$ & Chr(34)))
        End With
        Exit Sub
OnErrorHandle:
'        MsgBox "第:" & i & "行     " & Codes(i) & "     " & Err.Description
        MsgBox Err.Description
        Err.Clear
End Sub


''根据名称获取对象
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
        '添加许可证
        Licenses.Add ClassID
        
        '实例化类
        Set ContainerHeap(Count_Of_Container) = New ClassEvent
        With ContainerHeap(Count_Of_Container)
                '设置消息托管对象
                Set .Master = Me
                '设置类索引
                .Index = Count_Of_Container
                
                '创建ActiveX控件
                Set .Container = Me.Controls.Add(ClassID, "xContainer" & Count_Of_Container, Me)
                
                '添加到全局变量空间
                IScript.AddObject Alias, .Container, True
                
                With .Container
                        '设置控件属性
                        .Visible = Visible
                        .Left = Left
                        .Top = Top
                        .Width = Width
                        .Height = Height
                        
'                        '返回对象实例
'                        Set CreateActiveX = .Object
                End With
                '索引累加
                Count_Of_Container = Count_Of_Container + 1
        End With
        
        '动态分配内存
        ReDim Preserve ContainerHeap(Count_Of_Container)
End Sub

'提供给VBS的接口,用于动态生成对象
Public Function CreateObjectEx(Class, Optional ServerName As String) As Object
        Set CreateObjectEx = CreateObject(Class, ServerName)
End Function






Private Sub Form_Resize()
        If IScript Is Nothing Then Exit Sub
        '事件函数原型 控件索引,事件名称,参数个数,参数名称,参数类型,参数值
        'Sub EventHandle(index,name,pCount,pNames(),pType(),pValues())

        Dim Codes As String
        Codes = "Call EventHandle({0},{1},{2},{3},{4},{5})"
        Codes = Replace(Codes, "{0}", "-1")
        Codes = Replace(Codes, "{1}", Chr(34) & "xVB_Resize" & Chr(34))
        Codes = Replace(Codes, "{2}", "0")
        Codes = Replace(Codes, "{3}", "Array(vbnull)")
        Codes = Replace(Codes, "{4}", "Array(vbnull)")
        Codes = Replace(Codes, "{5}", "Array(vbnull)")
        '广播消息
        Call PostEvent(Codes)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'释放资源
On Error Resume Next
        Dim i As Long
        For i = UBound(ContainerHeap) - 1 To LBound(ContainerHeap) Step -1
                Set ContainerHeap(i) = Nothing
        Next i
        Set IScript = Nothing
        Set FormMain = Nothing
End Sub






