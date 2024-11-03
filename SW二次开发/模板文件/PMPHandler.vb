Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports SolidWorks.Interop.swpublished

' 定义一个名为PMPageHandler的公共类，它实现了PropertyManagerPage2Handler9接口
Public Class PMPageHandler
    Implements PropertyManagerPage2Handler9

    ' 定义一个SldWorks类型的变量，用于存储SolidWorks应用程序对象
    Dim iSwApp As SldWorks
    ' 定义一个SwAddin类型的变量，用于存储插件对象
    Dim userAddin As SwAddin
    ' 定义一个UserPMPage类型的变量，用于存储用户自定义的属性管理器页面对象
    Dim ppage As UserPMPage

    ' 初始化函数，用于接收SolidWorks应用程序对象、插件对象和用户自定义的属性管理器页面对象，并进行赋值
    Function Init(ByVal sw As SldWorks, ByVal addin As SwAddin, page As UserPMPage) As Integer
        iSwApp = sw
        userAddin = addin
        ppage = page
    End Function

    ' 实现PropertyManagerPage2Handler9接口中的AfterClose方法
    ' 即使此方法不执行任何实际操作，也必须包含代码，以防止.NET运行时环境在错误的时间进行垃圾回收
    Sub AfterClose() Implements PropertyManagerPage2Handler9.AfterClose
        ''This function must contain code, even if it does nothing, to prevent the
        ''.NET runtime environment from doing garbage collection at the wrong time.

        ' 定义一个整数变量IndentSize，用于存储调试缩进大小
        Dim IndentSize As Integer
        ' 获取当前调试缩进大小并赋值给IndentSize变量
        IndentSize = System.Diagnostics.Debug.IndentSize
        ' 将IndentSize的值输出到调试控制台
        System.Diagnostics.Debug.WriteLine(IndentSize)

    End Sub

    ' 实现PropertyManagerPage2Handler9接口中的OnCheckboxCheck方法
    ' 当复选框的状态发生改变时会调用此方法
    Sub OnCheckboxCheck(ByVal id As Integer, ByVal status As Boolean) Implements PropertyManagerPage2Handler9.OnCheckboxCheck
        ' 此处可添加针对复选框状态改变的具体处理逻辑，当前为空实现
    End Sub

    ' 实现PropertyManagerPage2Handler9接口中的OnClose方法
    ' 即使此方法不执行任何实际操作，也必须包含代码，以防止.NET运行时环境在错误的时间进行垃圾回收
    Sub OnClose(ByVal reason As Integer) Implements PropertyManagerPage2Handler9.OnClose
        ''This function must contain code, even if it does nothing, to prevent the
        ''.NET runtime environment from doing garbage collection at the wrong time.

        ' 定义一个整数变量IndentSize，用于存储调试缩进大小
        Dim IndentSize As Integer
        ' 获取当前调试缩进大小并赋值给IndentSize变量
        IndentSize = System.Diagnostics.Debug.IndentSize
        ' 将IndentSize的值输出到调试控制台
        System.Diagnostics.Debug.WriteLine(IndentSize)
    End Sub

    ' 实现PropertyManagerPage2Handler9接口中的OnComboboxEditChanged方法
    ' 当组合框的编辑内容发生改变时会调用此方法
    Sub OnComboboxEditChanged(ByVal id As Integer, ByVal text As String) Implements PropertyManagerPage2Handler9.OnComboboxEditChanged
        ' 此处可添加针对组合框编辑内容改变的具体处理逻辑，当前为空实现
    End Sub

    ' 实现PropertyManagerPage2Handler9接口中的OnActiveXControlCreated方法
    ' 当ActiveX控件被创建时会调用此方法，并返回一个整数结果
    Function OnActiveXControlCreated(ByVal id As Integer, ByVal status As Boolean) As Integer Implements PropertyManagerPage2Handler9.OnActiveXControlCreated
        ' 默认返回 -1，可根据实际需求修改返回值以表示不同的创建状态或结果
        OnActiveXControlCreated = -1
    End Function

    ' 实现PropertyManagerPage2Handler9接口中的OnButtonPress方法
    ' 当按钮被按下时会调用此方法
    Sub OnButtonPress(ByVal id As Integer) Implements PropertyManagerPage2Handler9.OnButtonPress
        ' 判断按下的按钮是否是ppage中的buttonID1按钮
        If id = ppage.buttonID1 Then
            ' 如果是，切换text1文本框的可见性状态
            If ppage.text1.Visible = True Then
                ppage.text1.Visible = False
            Else
                ppage.text1.Visible = True
            End If

            ' 判断按下的按钮是否是ppage中的buttonID2按钮
        ElseIf id = ppage.buttonID2 Then
            ' 如果是，切换text2文本框的启用/禁用状态
            If ppage.text2.Enabled = True Then
                ppage.text2.Enabled = False
            Else
                ppage.text2.Enabled = True
            End If
        End If
    End Sub

    ' 实现PropertyManagerPage2Handler9接口中的OnComboboxSelectionChanged方法
    ' 当组合框的选择项发生改变时会调用此方法
    Sub OnComboboxSelectionChanged(ByVal id As Integer, ByVal item As Integer) Implements PropertyManagerPage2Handler9.OnComboboxSelectionChanged
        ' 此处可添加针对组合框选择项改变的具体处理逻辑，当前为空实现
    End Sub

    ' 实现PropertyManagerPage2Handler9接口中的OnGroupCheck方法
    ' 当组的选中状态发生改变时会调用此方法
    Sub OnGroupCheck(ByVal id As Integer, ByVal status As Boolean) Implements PropertyManagerPage2Handler9.OnGroupCheck
        ' 此处可添加针对组选中状态改变的具体处理逻辑，当前为空实现
    End Sub

    ' 实现PropertyManagerPage2Handler9接口中的OnGroupExpand方法
    ' 当组展开或折叠状态发生改变时会调用此方法
    Sub OnGroupExpand(ByVal id As Integer, ByVal status As Boolean) Implements PropertyManagerPage2Handler9.OnGroupExpand
        ' 此处可添加针对组展开/折叠状态改变的具体处理逻辑，当前为空实现
    End Sub

    ' 实现PropertyManagerPage2Handler9接口中的OnHelp方法
    ' 当用户请求帮助时会调用此方法，返回一个布尔值表示是否成功提供帮助
    Function OnHelp() As Boolean Implements PropertyManagerPage2Handler9.OnHelp
        ' 定义一个字符串变量helppath，用于存储帮助文件的路径
        Dim helppath As String

        ' 指定一个URL路径或一个CHM文件的路径作为帮助文件的路径
        helppath = "http://help.solidworks.com/2016/English/api/sldworksapiprogguide/Welcome.htm"
        'helppath = "C:\Program Files\SolidWorks Corp\SOLIDWORKS\api\apihelp.chm"

        ' 创建一个Windows Forms表单对象，用于显示帮助内容
        Dim helpForm As System.Windows.Forms.Form
        helpForm = New System.Windows.Forms.Form

        ' 使用System.Windows.Forms.Help类的ShowHelp方法在创建的表单中显示指定路径的帮助内容
        System.Windows.Forms.Help.ShowHelp(helpForm, helppath)

        ' 返回True表示成功提供帮助（可根据实际情况修改返回值逻辑）
        OnHelp = True
    End Function

    ' 实现PropertyManagerPage2Handler9接口中的OnListboxSelectionChanged方法
    ' 当列表框的选择项发生改变时会调用此方法
    Sub OnListboxSelectionChanged(ByVal id As Integer, ByVal item As Integer) Implements PropertyManagerPage2Handler9.OnListboxSelectionChanged
        ' 此处可添加针对列表框选择项改变的具体处理逻辑，当前为空实现
    End Sub

    ' 实现PropertyManagerPage2Handler9接口中的OnNextPage方法
    ' 当用户请求切换到下一页时会调用此方法，返回一个布尔值表示是否允许切换
    Function OnNextPage() As Boolean Implements PropertyManagerPage2Handler9.OnNextPage
        ' 默认返回True，表示允许切换到下一页（可根据实际需求修改返回值逻辑）
        OnNextPage = True
    End Function

    ' 实现PropertyManagerPage2Handler9接口中的OnNumberboxChanged方法
    ' 当数字框的值发生改变时会调用此方法
    Sub OnNumberboxChanged(ByVal id As Integer, ByVal val As Double) Implements PropertyManagerPage2Handler9.OnNumberboxChanged
        ' 此处可添加针对数字框值改变的具体处理逻辑，当前为空实现
    End Sub

    ' 实现PropertyManagerPage2Handler9接口中的OnOptionCheck方法
    ' 当选项的选中状态发生改变时会调用此方法
    Sub OnOptionCheck(ByVal id As Integer) Implements PropertyManagerPage2Handler9.OnOptionCheck
        ' 此处可添加针对选项选中状态改变的具体处理逻辑，当前为空实现
    End Sub

    ' 实现PropertyManagerPage2Handler9接口中的OnPreviousPage方法
    ' 当用户请求切换到上一页时会调用此方法，返回一个布尔值表示是否允许切换
    Function OnPreviousPage() As Boolean Implements PropertyManagerPage2Handler9.OnPreviousPage
        ' 默认返回True，表示允许切换到上一页（可根据实际需求修改返回值逻辑）
        OnPreviousPage = True
    End Function

    ' 实现PropertyManagerPage2Handler9接口中的OnSelectionboxCalloutCreated方法
    ' 当选择框的标注被创建时会调用此方法
    Sub OnSelectionboxCalloutCreated(ByVal id As Integer) Implements PropertyManagerPage2Handler9.OnSelectionboxCalloutCreated
        ' 此处可添加针对选择框标注创建的具体处理逻辑，当前为空实现
    End Sub

    ' 实现PropertyManagerPage2Handler9接口中的OnSelectionboxCalloutDestroyed方法
    ' 当选择框的标注被销毁时会调用此方法
    Sub OnSelectionboxCalloutDestroyed(ByVal id As Integer) Implements PropertyManagerPage2Handler9.OnSelectionboxCalloutDestroyed
        ' 此处可添加针对选择框标注销毁的具体处理逻辑，当前为空实现
    End Sub

    ' 实现PropertyManagerPage2Handler9接口中的OnSelectionboxFocusChanged方法
    ' 当选择框的焦点发生改变时会调用此方法
    Sub OnSelectionboxFocusChanged(ByVal Id As Integer) Implements PropertyManagerPage2Handler9.OnSelectionboxFocusChanged
        ' 此处可添加针对选择框焦点改变的具体处理逻辑，当前为空实现
    End Sub

    ' 实现PropertyManagerPage2Handler9接口中的OnSelectionboxListChanged方法
    ' 当选择框的列表内容发生改变时会调用此方法
    Sub OnSelectionboxListChanged(ByVal id As Integer, ByVal item As Integer) Implements PropertyManagerPage2Handler9.OnSelectionboxListChanged
        ' 当用户选择实体来填充选择框时，设置光标为前进光标样式
        ppage.PropMgrPage.SetCursor(swPropertyManagerPageCursors_e.swPropertyManagerPageCursors_Advance)
    End Sub

    ' 实现PropertyManagerPage2Handler9接口中的OnTextboxChanged方法
    ' 当文本框的内容发生改变时会调用此方法
    Sub OnTextboxChanged(ByVal id As Integer, ByVal text As String) Implements PropertyManagerPage2Handler9.OnTextboxChanged
        ' 此处可添加针对文本框内容改变的具体处理逻辑，当前为空实现
    End Sub

    ' 实现SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9接口中的AfterActivation方法
    Public Sub AfterActivation() Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.AfterActivation
        ' 此处可添加在属性管理器页面激活后的具体处理逻辑，当前为空实现
    End Sub

    ' 实现SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9接口中的OnKeystroke方法
    Public Function OnKeystroke(ByVal Wparam As Integer, ByVal Message As Integer, ByVal Lparam As Integer, ByVal Id As Integer) As Boolean Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnKeystroke
        ' 此处可添加针对按键事件的具体处理逻辑，当前为空实现
        OnKeystroke = True
    End Function

    ' 实现SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9接口中的OnPopupMenuItem方法
    Public Sub OnPopupMenuItem(ByVal Id As Integer) Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnPopupMenuItem
        ' 此处可添加针对弹出菜单项点击事件的具体处理逻辑，当前为空实现
    End Sub

    ' 实现SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9接口中的OnPopupMenuItemUpdate方法
    Public Sub OnPopupMenuItemUpdate(ByVal Id As Integer, ByRef retval As Integer) Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnPopupMenuItemUpdate
        ' 此处可添加针对弹出菜单项更新事件的具体处理逻辑，当前为空实现
    End Sub

    ' 实现SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9接口中的OnPreview方法
    Public Function OnPreview() As Boolean Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnPreview
        ' 默认返回True，表示允许预览（可根据实际需求修改返回值逻辑）
        OnPreview = True
    End Function

    ' 实现SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9接口中的OnSliderPositionChanged方法
    Public Sub OnSliderPositionChanged(ByVal Id As Integer, ByVal Value As Double) Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnSliderPositionChanged
        ' 此处可添加针对滑块位置改变事件的具体处理逻辑，当前为空实现
    End Sub

    ' 实现SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9接口中的OnSliderTrackingCompleted方法
    Public Sub OnSliderTrackingCompleted(ByVal Id As Integer, ByVal Value As Double) Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnSliderTrackingCompleted
        ' 此处可添加针对滑块跟踪完成事件的具体处理逻辑，当前为空实现
    End Sub

    ' 实现SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9接口中的OnSubmitSelection方法
    Public Function OnSubmitSelection(ByVal Id As Integer, ByVal Selection As Object, ByVal SelType As Integer, ByRef ItemText As String) As Boolean Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnSubmitSelection
        ' 默认返回True，表示允许提交选择（可根据实际需求修改返回值逻辑）
        OnSubmitSelection = True
    End Function

    ' 实现SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9接口中的OnTabClicked方法
    Public Function OnTabClicked(ByVal Id As Integer) As Boolean Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnTabClicked
        ' 默认返回True，表示允许点击标签（可根据实际需求修改返回值逻辑）
        OnTabClicked = True
    End Function

    ' 实现SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9接口中的OnUndo方法
    Public Sub OnUndo() Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnUndo
        ' 此处可添加针对撤销操作的具体处理逻辑，当前为空实现
    End Sub

    ' 实现SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9接口中的OnWhatsNew方法
    Public Sub OnWhatsNew() Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnWhatsNew
        ' 此处可添加针对查看新特性操作的具体处理逻辑，当前为空实现
    End Sub

    ' 实现SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9接口中的OnWindowFromHandleControlCreated方法
    Function OnWindowFromHandleControlCreated(ByVal Id As Integer, ByVal Status As Boolean) As Integer Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnWindowFromHandleControlCreated
        ' 此处可添加针对从句柄创建窗口控件事件的具体处理逻辑，当前为空实现
        OnWindowFromHandleControlCreated = -1
    End Function

    ' 实现SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9接口中的OnGainedFocus方法
    Sub OnGainedFocus(ByVal Id As Integer) Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnGainedFocus
        ' 此处可添加针对获得焦点事件的具体处理逻辑，当前为空实现
    End Sub

    ' 实现SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9接口中的OnListboxRMBUp方法
    Sub OnListboxRMBUp(ByVal Id As Integer, ByVal PosX As Integer, ByVal PosY As Integer) Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnListboxRMBUp
        ' 此处可添加针对列表框右键抬起事件的具体处理逻辑，当前为空实现
    End Sub

    ' 实现SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9接口中的OnLostFocus方法
    Sub OnLostFocus(ByVal Id As Integer) Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnLostFocus
        ' 此处可添加针对失去焦点事件的具体处理逻辑，当前为空实现
    End Sub

    ' 实现SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9接口中的OnRedo方法
    Sub OnRedo() Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnRedo
        ' 此处可添加针对重做操作的具体处理逻辑，当前为空实现
    End Sub


    ' 实现SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9接口中的OnNumberBoxTrackingCompleted方法
    Sub OnNumberBoxTrackingCompleted(ByVal id As Integer, ByVal val As Double) Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnNumberBoxTrackingCompleted
        ' 此处可添加针对数字框跟踪完成事件的具体处理逻辑，当前为空实现
    End Sub
End Class