Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports SolidWorks.Interop.swpublished

' ����һ����ΪPMPageHandler�Ĺ����࣬��ʵ����PropertyManagerPage2Handler9�ӿ�
Public Class PMPageHandler
    Implements PropertyManagerPage2Handler9

    ' ����һ��SldWorks���͵ı��������ڴ洢SolidWorksӦ�ó������
    Dim iSwApp As SldWorks
    ' ����һ��SwAddin���͵ı��������ڴ洢�������
    Dim userAddin As SwAddin
    ' ����һ��UserPMPage���͵ı��������ڴ洢�û��Զ�������Թ�����ҳ�����
    Dim ppage As UserPMPage

    ' ��ʼ�����������ڽ���SolidWorksӦ�ó�����󡢲��������û��Զ�������Թ�����ҳ����󣬲����и�ֵ
    Function Init(ByVal sw As SldWorks, ByVal addin As SwAddin, page As UserPMPage) As Integer
        iSwApp = sw
        userAddin = addin
        ppage = page
    End Function

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�AfterClose����
    ' ��ʹ�˷�����ִ���κ�ʵ�ʲ�����Ҳ����������룬�Է�ֹ.NET����ʱ�����ڴ����ʱ�������������
    Sub AfterClose() Implements PropertyManagerPage2Handler9.AfterClose
        ''This function must contain code, even if it does nothing, to prevent the
        ''.NET runtime environment from doing garbage collection at the wrong time.

        ' ����һ����������IndentSize�����ڴ洢����������С
        Dim IndentSize As Integer
        ' ��ȡ��ǰ����������С����ֵ��IndentSize����
        IndentSize = System.Diagnostics.Debug.IndentSize
        ' ��IndentSize��ֵ��������Կ���̨
        System.Diagnostics.Debug.WriteLine(IndentSize)

    End Sub

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnCheckboxCheck����
    ' ����ѡ���״̬�����ı�ʱ����ô˷���
    Sub OnCheckboxCheck(ByVal id As Integer, ByVal status As Boolean) Implements PropertyManagerPage2Handler9.OnCheckboxCheck
        ' �˴��������Ը�ѡ��״̬�ı�ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnClose����
    ' ��ʹ�˷�����ִ���κ�ʵ�ʲ�����Ҳ����������룬�Է�ֹ.NET����ʱ�����ڴ����ʱ�������������
    Sub OnClose(ByVal reason As Integer) Implements PropertyManagerPage2Handler9.OnClose
        ''This function must contain code, even if it does nothing, to prevent the
        ''.NET runtime environment from doing garbage collection at the wrong time.

        ' ����һ����������IndentSize�����ڴ洢����������С
        Dim IndentSize As Integer
        ' ��ȡ��ǰ����������С����ֵ��IndentSize����
        IndentSize = System.Diagnostics.Debug.IndentSize
        ' ��IndentSize��ֵ��������Կ���̨
        System.Diagnostics.Debug.WriteLine(IndentSize)
    End Sub

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnComboboxEditChanged����
    ' ����Ͽ�ı༭���ݷ����ı�ʱ����ô˷���
    Sub OnComboboxEditChanged(ByVal id As Integer, ByVal text As String) Implements PropertyManagerPage2Handler9.OnComboboxEditChanged
        ' �˴�����������Ͽ�༭���ݸı�ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnActiveXControlCreated����
    ' ��ActiveX�ؼ�������ʱ����ô˷�����������һ���������
    Function OnActiveXControlCreated(ByVal id As Integer, ByVal status As Boolean) As Integer Implements PropertyManagerPage2Handler9.OnActiveXControlCreated
        ' Ĭ�Ϸ��� -1���ɸ���ʵ�������޸ķ���ֵ�Ա�ʾ��ͬ�Ĵ���״̬����
        OnActiveXControlCreated = -1
    End Function

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnButtonPress����
    ' ����ť������ʱ����ô˷���
    Sub OnButtonPress(ByVal id As Integer) Implements PropertyManagerPage2Handler9.OnButtonPress
        ' �жϰ��µİ�ť�Ƿ���ppage�е�buttonID1��ť
        If id = ppage.buttonID1 Then
            ' ����ǣ��л�text1�ı���Ŀɼ���״̬
            If ppage.text1.Visible = True Then
                ppage.text1.Visible = False
            Else
                ppage.text1.Visible = True
            End If

            ' �жϰ��µİ�ť�Ƿ���ppage�е�buttonID2��ť
        ElseIf id = ppage.buttonID2 Then
            ' ����ǣ��л�text2�ı��������/����״̬
            If ppage.text2.Enabled = True Then
                ppage.text2.Enabled = False
            Else
                ppage.text2.Enabled = True
            End If
        End If
    End Sub

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnComboboxSelectionChanged����
    ' ����Ͽ��ѡ������ı�ʱ����ô˷���
    Sub OnComboboxSelectionChanged(ByVal id As Integer, ByVal item As Integer) Implements PropertyManagerPage2Handler9.OnComboboxSelectionChanged
        ' �˴�����������Ͽ�ѡ����ı�ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnGroupCheck����
    ' �����ѡ��״̬�����ı�ʱ����ô˷���
    Sub OnGroupCheck(ByVal id As Integer, ByVal status As Boolean) Implements PropertyManagerPage2Handler9.OnGroupCheck
        ' �˴�����������ѡ��״̬�ı�ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnGroupExpand����
    ' ����չ�����۵�״̬�����ı�ʱ����ô˷���
    Sub OnGroupExpand(ByVal id As Integer, ByVal status As Boolean) Implements PropertyManagerPage2Handler9.OnGroupExpand
        ' �˴�����������չ��/�۵�״̬�ı�ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnHelp����
    ' ���û��������ʱ����ô˷���������һ������ֵ��ʾ�Ƿ�ɹ��ṩ����
    Function OnHelp() As Boolean Implements PropertyManagerPage2Handler9.OnHelp
        ' ����һ���ַ�������helppath�����ڴ洢�����ļ���·��
        Dim helppath As String

        ' ָ��һ��URL·����һ��CHM�ļ���·����Ϊ�����ļ���·��
        helppath = "http://help.solidworks.com/2016/English/api/sldworksapiprogguide/Welcome.htm"
        'helppath = "C:\Program Files\SolidWorks Corp\SOLIDWORKS\api\apihelp.chm"

        ' ����һ��Windows Forms������������ʾ��������
        Dim helpForm As System.Windows.Forms.Form
        helpForm = New System.Windows.Forms.Form

        ' ʹ��System.Windows.Forms.Help���ShowHelp�����ڴ����ı�����ʾָ��·���İ�������
        System.Windows.Forms.Help.ShowHelp(helpForm, helppath)

        ' ����True��ʾ�ɹ��ṩ�������ɸ���ʵ������޸ķ���ֵ�߼���
        OnHelp = True
    End Function

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnListboxSelectionChanged����
    ' ���б���ѡ������ı�ʱ����ô˷���
    Sub OnListboxSelectionChanged(ByVal id As Integer, ByVal item As Integer) Implements PropertyManagerPage2Handler9.OnListboxSelectionChanged
        ' �˴����������б��ѡ����ı�ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnNextPage����
    ' ���û������л�����һҳʱ����ô˷���������һ������ֵ��ʾ�Ƿ������л�
    Function OnNextPage() As Boolean Implements PropertyManagerPage2Handler9.OnNextPage
        ' Ĭ�Ϸ���True����ʾ�����л�����һҳ���ɸ���ʵ�������޸ķ���ֵ�߼���
        OnNextPage = True
    End Function

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnNumberboxChanged����
    ' �����ֿ��ֵ�����ı�ʱ����ô˷���
    Sub OnNumberboxChanged(ByVal id As Integer, ByVal val As Double) Implements PropertyManagerPage2Handler9.OnNumberboxChanged
        ' �˴������������ֿ�ֵ�ı�ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnOptionCheck����
    ' ��ѡ���ѡ��״̬�����ı�ʱ����ô˷���
    Sub OnOptionCheck(ByVal id As Integer) Implements PropertyManagerPage2Handler9.OnOptionCheck
        ' �˴���������ѡ��ѡ��״̬�ı�ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnPreviousPage����
    ' ���û������л�����һҳʱ����ô˷���������һ������ֵ��ʾ�Ƿ������л�
    Function OnPreviousPage() As Boolean Implements PropertyManagerPage2Handler9.OnPreviousPage
        ' Ĭ�Ϸ���True����ʾ�����л�����һҳ���ɸ���ʵ�������޸ķ���ֵ�߼���
        OnPreviousPage = True
    End Function

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnSelectionboxCalloutCreated����
    ' ��ѡ���ı�ע������ʱ����ô˷���
    Sub OnSelectionboxCalloutCreated(ByVal id As Integer) Implements PropertyManagerPage2Handler9.OnSelectionboxCalloutCreated
        ' �˴���������ѡ����ע�����ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnSelectionboxCalloutDestroyed����
    ' ��ѡ���ı�ע������ʱ����ô˷���
    Sub OnSelectionboxCalloutDestroyed(ByVal id As Integer) Implements PropertyManagerPage2Handler9.OnSelectionboxCalloutDestroyed
        ' �˴���������ѡ����ע���ٵľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnSelectionboxFocusChanged����
    ' ��ѡ���Ľ��㷢���ı�ʱ����ô˷���
    Sub OnSelectionboxFocusChanged(ByVal Id As Integer) Implements PropertyManagerPage2Handler9.OnSelectionboxFocusChanged
        ' �˴���������ѡ��򽹵�ı�ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnSelectionboxListChanged����
    ' ��ѡ�����б����ݷ����ı�ʱ����ô˷���
    Sub OnSelectionboxListChanged(ByVal id As Integer, ByVal item As Integer) Implements PropertyManagerPage2Handler9.OnSelectionboxListChanged
        ' ���û�ѡ��ʵ�������ѡ���ʱ�����ù��Ϊǰ�������ʽ
        ppage.PropMgrPage.SetCursor(swPropertyManagerPageCursors_e.swPropertyManagerPageCursors_Advance)
    End Sub

    ' ʵ��PropertyManagerPage2Handler9�ӿ��е�OnTextboxChanged����
    ' ���ı�������ݷ����ı�ʱ����ô˷���
    Sub OnTextboxChanged(ByVal id As Integer, ByVal text As String) Implements PropertyManagerPage2Handler9.OnTextboxChanged
        ' �˴����������ı������ݸı�ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9�ӿ��е�AfterActivation����
    Public Sub AfterActivation() Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.AfterActivation
        ' �˴�����������Թ�����ҳ�漤���ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9�ӿ��е�OnKeystroke����
    Public Function OnKeystroke(ByVal Wparam As Integer, ByVal Message As Integer, ByVal Lparam As Integer, ByVal Id As Integer) As Boolean Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnKeystroke
        ' �˴��������԰����¼��ľ��崦���߼�����ǰΪ��ʵ��
        OnKeystroke = True
    End Function

    ' ʵ��SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9�ӿ��е�OnPopupMenuItem����
    Public Sub OnPopupMenuItem(ByVal Id As Integer) Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnPopupMenuItem
        ' �˴��������Ե����˵������¼��ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9�ӿ��е�OnPopupMenuItemUpdate����
    Public Sub OnPopupMenuItemUpdate(ByVal Id As Integer, ByRef retval As Integer) Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnPopupMenuItemUpdate
        ' �˴��������Ե����˵�������¼��ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9�ӿ��е�OnPreview����
    Public Function OnPreview() As Boolean Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnPreview
        ' Ĭ�Ϸ���True����ʾ����Ԥ�����ɸ���ʵ�������޸ķ���ֵ�߼���
        OnPreview = True
    End Function

    ' ʵ��SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9�ӿ��е�OnSliderPositionChanged����
    Public Sub OnSliderPositionChanged(ByVal Id As Integer, ByVal Value As Double) Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnSliderPositionChanged
        ' �˴��������Ի���λ�øı��¼��ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9�ӿ��е�OnSliderTrackingCompleted����
    Public Sub OnSliderTrackingCompleted(ByVal Id As Integer, ByVal Value As Double) Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnSliderTrackingCompleted
        ' �˴��������Ի����������¼��ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9�ӿ��е�OnSubmitSelection����
    Public Function OnSubmitSelection(ByVal Id As Integer, ByVal Selection As Object, ByVal SelType As Integer, ByRef ItemText As String) As Boolean Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnSubmitSelection
        ' Ĭ�Ϸ���True����ʾ�����ύѡ�񣨿ɸ���ʵ�������޸ķ���ֵ�߼���
        OnSubmitSelection = True
    End Function

    ' ʵ��SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9�ӿ��е�OnTabClicked����
    Public Function OnTabClicked(ByVal Id As Integer) As Boolean Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnTabClicked
        ' Ĭ�Ϸ���True����ʾ��������ǩ���ɸ���ʵ�������޸ķ���ֵ�߼���
        OnTabClicked = True
    End Function

    ' ʵ��SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9�ӿ��е�OnUndo����
    Public Sub OnUndo() Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnUndo
        ' �˴��������Գ��������ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9�ӿ��е�OnWhatsNew����
    Public Sub OnWhatsNew() Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnWhatsNew
        ' �˴��������Բ鿴�����Բ����ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9�ӿ��е�OnWindowFromHandleControlCreated����
    Function OnWindowFromHandleControlCreated(ByVal Id As Integer, ByVal Status As Boolean) As Integer Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnWindowFromHandleControlCreated
        ' �˴��������ԴӾ���������ڿؼ��¼��ľ��崦���߼�����ǰΪ��ʵ��
        OnWindowFromHandleControlCreated = -1
    End Function

    ' ʵ��SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9�ӿ��е�OnGainedFocus����
    Sub OnGainedFocus(ByVal Id As Integer) Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnGainedFocus
        ' �˴��������Ի�ý����¼��ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9�ӿ��е�OnListboxRMBUp����
    Sub OnListboxRMBUp(ByVal Id As Integer, ByVal PosX As Integer, ByVal PosY As Integer) Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnListboxRMBUp
        ' �˴����������б���Ҽ�̧���¼��ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9�ӿ��е�OnLostFocus����
    Sub OnLostFocus(ByVal Id As Integer) Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnLostFocus
        ' �˴���������ʧȥ�����¼��ľ��崦���߼�����ǰΪ��ʵ��
    End Sub

    ' ʵ��SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9�ӿ��е�OnRedo����
    Sub OnRedo() Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnRedo
        ' �˴������������������ľ��崦���߼�����ǰΪ��ʵ��
    End Sub


    ' ʵ��SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9�ӿ��е�OnNumberBoxTrackingCompleted����
    Sub OnNumberBoxTrackingCompleted(ByVal id As Integer, ByVal val As Double) Implements SolidWorks.Interop.swpublished.IPropertyManagerPage2Handler9.OnNumberBoxTrackingCompleted
        ' �˴������������ֿ��������¼��ľ��崦���߼�����ǰΪ��ʵ��
    End Sub
End Class