Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Diagnostics
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports SolidWorks.Interop.sldworks
Imports SolidWorks.Interop.swconst
Imports SolidWorks.Interop.swpublished
Imports SolidWorksTools
Imports SolidWorksTools.File

<Guid("d0e5cde0-03f9-4b4b-92a0-346d5ff5ac31")> _
<ComVisible(True)> _
<SwAddin( _
        Description:="SW二次开发 description", _
        Title:="SW二次开发", _
        LoadAtStartup:=True _
        )> _
Public Class SwAddin
    Implements SolidWorks.Interop.swpublished.SwAddin

#Region "局部变量"
    Dim WithEvents iSwApp As SldWorks
    Dim iCmdMgr As ICommandManager
    Dim addinID As Integer
    Dim openDocs As Hashtable
    Dim SwEventPtr As SldWorks
    Dim ppage As UserPMPage
    Dim iBmp As BitmapHandler
    Dim frame As IFrame
    Dim bRet As Boolean
    Dim registerID As Integer

    Public Const mainCmdGroupID As Integer = 0
    Public Const mainItemID1 As Integer = 0
    Public Const mainItemID2 As Integer = 1
    Public Const flyoutGroupID As Integer = 91

    ' Public Properties
    ReadOnly Property SwApp() As SldWorks
        Get
            Return iSwApp
        End Get
    End Property

    ReadOnly Property CmdMgr() As ICommandManager
        Get
            Return iCmdMgr
        End Get
    End Property

    ReadOnly Property OpenDocumentsTable() As Hashtable
        Get
            Return openDocs
        End Get
    End Property
#End Region

#Region "SolidWorks 注册"

    <ComRegisterFunction()> Public Shared Sub RegisterFunction(ByVal t As Type)

        ' Get Custom Attribute: SwAddinAttribute
        Dim attributes() As Object
        Dim SWattr As SwAddinAttribute = Nothing

        attributes = System.Attribute.GetCustomAttributes(GetType(SwAddin), GetType(SwAddinAttribute))

        If attributes.Length > 0 Then
            SWattr = DirectCast(attributes(0), SwAddinAttribute)
        End If
        Try
            Dim hklm As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine
            Dim hkcu As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser

            Dim keyname As String = "SOFTWARE\SolidWorks\Addins\{" + t.GUID.ToString() + "}"
            Dim addinkey As Microsoft.Win32.RegistryKey = hklm.CreateSubKey(keyname)
            addinkey.SetValue(Nothing, 0)
            addinkey.SetValue("Description", SWattr.Description)
            addinkey.SetValue("Title", SWattr.Title)

            keyname = "Software\SolidWorks\AddInsStartup\{" + t.GUID.ToString() + "}"
            addinkey = hkcu.CreateSubKey(keyname)
            addinkey.SetValue(Nothing, SWattr.LoadAtStartup, Microsoft.Win32.RegistryValueKind.DWord)
        Catch nl As System.NullReferenceException
            Console.WriteLine("There was a problem registering this dll: SWattr is null.\n " & nl.Message)
            System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n" & nl.Message)
        Catch e As System.Exception
            Console.WriteLine("There was a problem registering this dll: " & e.Message)
            System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: " & e.Message)
        End Try
    End Sub

    <ComUnregisterFunction()> Public Shared Sub UnregisterFunction(ByVal t As Type)
        Try
            Dim hklm As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine
            Dim hkcu As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser

            Dim keyname As String = "SOFTWARE\SolidWorks\Addins\{" + t.GUID.ToString() + "}"
            hklm.DeleteSubKey(keyname)

            keyname = "Software\SolidWorks\AddInsStartup\{" + t.GUID.ToString() + "}"
            hkcu.DeleteSubKey(keyname)
        Catch nl As System.NullReferenceException
            Console.WriteLine("There was a problem unregistering this dll: SWattr is null.\n " & nl.Message)
            System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: SWattr is null.\n" & nl.Message)
        Catch e As System.Exception
            Console.WriteLine("There was a problem unregistering this dll: " & e.Message)
            System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: " & e.Message)
        End Try

    End Sub

#End Region

#Region "ISw 插件实现"

    Function ConnectToSW(ByVal ThisSW As Object, ByVal Cookie As Integer) As Boolean Implements SolidWorks.Interop.swpublished.SwAddin.ConnectToSW
        iSwApp = ThisSW
        addinID = Cookie

        ' Setup callbacks
        iSwApp.SetAddinCallbackInfo(0, Me, addinID)

        ' Setup the Command Manager
        iCmdMgr = iSwApp.GetCommandManager(Cookie)
        AddCommandMgr()

        'Setup the Event Handlers
        SwEventPtr = iSwApp
        openDocs = New Hashtable
        AttachEventHandlers()

        'Setup Sample Property Manager
        AddPMP()

        ConnectToSW = True
    End Function

    Function DisconnectFromSW() As Boolean Implements SolidWorks.Interop.swpublished.SwAddin.DisconnectFromSW

        RemoveCommandMgr()
        RemovePMP()
        DetachEventHandlers()

        System.Runtime.InteropServices.Marshal.ReleaseComObject(iCmdMgr)
        iCmdMgr = Nothing
        System.Runtime.InteropServices.Marshal.ReleaseComObject(iSwApp)
        iSwApp = Nothing
        'The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
        GC.Collect()
        GC.WaitForPendingFinalizers()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        DisconnectFromSW = True
    End Function
#End Region

#Region "UI 方法"
    Public Sub AddCommandMgr()

        Dim cmdGroup As ICommandGroup

        If iBmp Is Nothing Then
            iBmp = New BitmapHandler()
        End If

        Dim thisAssembly As Assembly

        Dim cmdIndex0 As Integer, cmdIndex1 As Integer
        Dim Title As String = "VB Addin-插件"
        Dim ToolTip As String = "VB Addin"


        Dim docTypes() As Integer = {swDocumentTypes_e.swDocASSEMBLY, _
                                       swDocumentTypes_e.swDocDRAWING, _
                                       swDocumentTypes_e.swDocPART}

        thisAssembly = System.Reflection.Assembly.GetAssembly(Me.GetType())

        Dim cmdGroupErr As Integer = 0
        Dim ignorePrevious As Boolean = False

        Dim registryIDs As Object = Nothing
        Dim getDataResult As Boolean = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, registryIDs)

        Dim knownIDs As Integer() = New Integer(1) {mainItemID1, mainItemID2}

        If getDataResult Then
            If Not CompareIDs(registryIDs, knownIDs) Then 'if the IDs don't match, reset the commandGroup
                ignorePrevious = True
            End If
        End If

        cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "", -1, ignorePrevious, cmdGroupErr)
        If cmdGroup Is Nothing Or thisAssembly Is Nothing Then
            Throw New NullReferenceException()
        End If

        ' Add bitmaps to your project and set them as embedded resources or provide a direct path to the bitmaps
        Dim mainIcons(6) As String
        Dim icons(6) As String
        icons(0) = iBmp.CreateFileFromResourceBitmap("SW二次开发.toolbar20x.png", thisAssembly)
        icons(1) = iBmp.CreateFileFromResourceBitmap("SW二次开发.toolbar32x.png", thisAssembly)
        icons(2) = iBmp.CreateFileFromResourceBitmap("SW二次开发.toolbar40x.png", thisAssembly)
        icons(3) = iBmp.CreateFileFromResourceBitmap("SW二次开发.toolbar64x.png", thisAssembly)
        icons(4) = iBmp.CreateFileFromResourceBitmap("SW二次开发.toolbar96x.png", thisAssembly)
        icons(5) = iBmp.CreateFileFromResourceBitmap("SW二次开发.toolbar128x.png", thisAssembly)

        mainIcons(0) = iBmp.CreateFileFromResourceBitmap("SW二次开发.mainicon_20.png", thisAssembly)
        mainIcons(1) = iBmp.CreateFileFromResourceBitmap("SW二次开发.mainicon_32.png", thisAssembly)
        mainIcons(2) = iBmp.CreateFileFromResourceBitmap("SW二次开发.mainicon_40.png", thisAssembly)
        mainIcons(3) = iBmp.CreateFileFromResourceBitmap("SW二次开发.mainicon_64.png", thisAssembly)
        mainIcons(4) = iBmp.CreateFileFromResourceBitmap("SW二次开发.mainicon_96.png", thisAssembly)
        mainIcons(5) = iBmp.CreateFileFromResourceBitmap("SW二次开发.mainicon_128.png", thisAssembly)

        cmdGroup.IconList = icons
        cmdGroup.MainIconList = mainIcons

        Dim menuToolbarOption As Integer = swCommandItemType_e.swMenuItem Or swCommandItemType_e.swToolbarItem

        cmdIndex0 = cmdGroup.AddCommandItem2("CreateCube", -1, "Create a cube", "Create cube", 0, "CreateCube", "", mainItemID1, menuToolbarOption)
        cmdIndex1 = cmdGroup.AddCommandItem2("Show PMP", -1, "Display sample property manager", "Show PMP", 2, "ShowPMP", "PMPEnable", mainItemID2, menuToolbarOption)

        cmdGroup.HasToolbar = True
        cmdGroup.HasMenu = True
        cmdGroup.Activate()

        Dim flyGroup As FlyoutGroup
        flyGroup = iCmdMgr.CreateFlyoutGroup2(flyoutGroupID, "Dynamic Flyout", "Flyout Tooltip", "Flyout Hint",
              cmdGroup.MainIconList, cmdGroup.IconList, "FlyoutCallback", "FlyoutEnable")

        flyGroup.AddCommandItem("FlyoutCommand 1", "test", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1")

        flyGroup.FlyoutType = swCommandFlyoutStyle_e.swCommandFlyoutStyle_Simple


        For Each docType As Integer In docTypes
            Dim cmdTab As ICommandTab = iCmdMgr.GetCommandTab(docType, Title)
            Dim bResult As Boolean

            If Not cmdTab Is Nothing And Not getDataResult Or ignorePrevious Then 'if tab exists, but we have ignored the registry info, re-create the tab.  Otherwise the ids won't matchup and the tab will be blank
                Dim res As Boolean = iCmdMgr.RemoveCommandTab(cmdTab)
                cmdTab = Nothing
            End If

            If cmdTab Is Nothing Then
                cmdTab = iCmdMgr.AddCommandTab(docType, Title)

                Dim cmdBox As CommandTabBox = cmdTab.AddCommandTabBox

                Dim cmdIDs(3) As Integer
                Dim TextType(3) As Integer

                cmdIDs(0) = cmdGroup.CommandID(cmdIndex0)
                TextType(0) = swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal

                cmdIDs(1) = cmdGroup.CommandID(cmdIndex1)
                TextType(1) = swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal

                cmdIDs(2) = cmdGroup.ToolbarId
                TextType(2) = swCommandTabButtonTextDisplay_e.swCommandTabButton_TextHorizontal


                bResult = cmdBox.AddCommands(cmdIDs, TextType)

                Dim cmdBox1 As CommandTabBox = cmdTab.AddCommandTabBox()
                ReDim cmdIDs(1)
                ReDim TextType(1)

                cmdIDs(0) = flyGroup.CmdID
                TextType(0) = swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow

                bResult = cmdBox1.AddCommands(cmdIDs, TextType)

                cmdTab.AddSeparator(cmdBox1, cmdIDs(0))

            End If
        Next

        ' create third-party icon in the context sensitive menus of faces in parts
        ' To see this menu, right click on any face in the part

        frame = iSwApp.Frame()
        bRet = frame.AddMenuPopupIcon3(swDocumentTypes_e.swDocPART, swSelectType_e.swSelFACES, "third-party context-sensitive", addinID, "PopupCallbackFunction", "PopupEnable", "", cmdGroup.MainIconList)

        ' create and register the shortcut menu
        registerID = iSwApp.RegisterThirdPartyPopupMenu()
        ' add a menu break at the top of the shortcut menu
        bRet = iSwApp.AddItemToThirdPartyPopupMenu2(registerID, CInt(swDocumentTypes_e.swDocPART), "Menu Break", addinID, "", "", "", "", "", CInt(swMenuItemType_e.swMenuItemType_Break))
        ' add a couple of items to the shortcut menu
        bRet = iSwApp.AddItemToThirdPartyPopupMenu2(registerID, CInt(swDocumentTypes_e.swDocPART), "Test1", addinID, "TestCallback", "EnableTest", "", "Test1", mainIcons(0), CInt(swMenuItemType_e.swMenuItemType_Default))
        bRet = iSwApp.AddItemToThirdPartyPopupMenu2(registerID, CInt(swDocumentTypes_e.swDocPART), "Test4", addinID, "TestCallback", "EnableTest", "", "Test4", mainIcons(0), CInt(swMenuItemType_e.swMenuItemType_Default))
        ' add a separator bar to the shortcut menu 
        bRet = iSwApp.AddItemToThirdPartyPopupMenu2(registerID, CInt(swDocumentTypes_e.swDocPART), "separator", addinID, "", "", "", "", "", CInt(swMenuItemType_e.swMenuItemType_Separator))
        ' add another item to the shortcut menu
        bRet = iSwApp.AddItemToThirdPartyPopupMenu2(registerID, CInt(swDocumentTypes_e.swDocPART), "Test5", addinID, "TestCallback", "EnableTest", "", "Test5", mainIcons(0), CInt(swMenuItemType_e.swMenuItemType_Default))

        ' add an icon to a menu bar of the shortcut menu
        bRet = iSwApp.AddItemToThirdPartyPopupMenu2(registerID, CInt(swDocumentTypes_e.swDocPART), "", addinID, "TestCallback", "EnableTest", "", "NoOp", mainIcons(0), CInt(swMenuItemType_e.swMenuItemType_Default))

        thisAssembly = Nothing

    End Sub


    Public Sub RemoveCommandMgr()
        Try
            iBmp.Dispose()
            iCmdMgr.RemoveCommandGroup(mainCmdGroupID)
            iCmdMgr.RemoveFlyoutGroup(flyoutGroupID)
        Catch e As Exception
        End Try
    End Sub


    Function AddPMP() As Boolean
        ppage = New UserPMPage
        ppage.Init(iSwApp, Me)
    End Function

    Function RemovePMP() As Boolean
        ppage = Nothing
    End Function

    Function CompareIDs(ByVal storedIDs() As Integer, ByVal addinIDs() As Integer) As Boolean

        Dim storeList As New List(Of Integer)(storedIDs)
        Dim addinList As New List(Of Integer)(addinIDs)

        addinList.Sort()
        storeList.Sort()

        If Not addinList.Count = storeList.Count Then

            Return False
        Else

            For i As Integer = 0 To addinList.Count - 1
                If Not addinList(i) = storeList(i) Then

                    Return False
                End If
            Next
        End If

        Return True
    End Function
#End Region

#Region "事件方法"
    Sub AttachEventHandlers()
        AttachSWEvents()

        'Listen for events on all currently open docs
        AttachEventsToAllDocuments()
    End Sub

    Sub DetachEventHandlers()
        DetachSWEvents()

        'Close events on all currently open docs
        Dim docHandler As DocumentEventHandler
        Dim key As ModelDoc2
        Dim numKeys As Integer
        numKeys = openDocs.Count
        If numKeys > 0 Then
            Dim keys() As Object = New Object(numKeys - 1) {}

            'Remove all document event handlers
            openDocs.Keys.CopyTo(keys, 0)
            For Each key In keys
                docHandler = openDocs.Item(key)
                docHandler.DetachEventHandlers() 'This also removes the pair from the hash
                docHandler = Nothing
                key = Nothing
            Next
        End If
    End Sub

    Sub AttachSWEvents()
        Try
            AddHandler iSwApp.ActiveDocChangeNotify, AddressOf Me.SldWorks_ActiveDocChangeNotify
            AddHandler iSwApp.DocumentLoadNotify2, AddressOf Me.SldWorks_DocumentLoadNotify2
            AddHandler iSwApp.FileNewNotify2, AddressOf Me.SldWorks_FileNewNotify2
            AddHandler iSwApp.ActiveModelDocChangeNotify, AddressOf Me.SldWorks_ActiveModelDocChangeNotify
            AddHandler iSwApp.FileOpenPostNotify, AddressOf Me.SldWorks_FileOpenPostNotify
        Catch e As Exception
            Console.WriteLine(e.Message)
        End Try
    End Sub

    Sub DetachSWEvents()
        Try
            RemoveHandler iSwApp.ActiveDocChangeNotify, AddressOf Me.SldWorks_ActiveDocChangeNotify
            RemoveHandler iSwApp.DocumentLoadNotify2, AddressOf Me.SldWorks_DocumentLoadNotify2
            RemoveHandler iSwApp.FileNewNotify2, AddressOf Me.SldWorks_FileNewNotify2
            RemoveHandler iSwApp.ActiveModelDocChangeNotify, AddressOf Me.SldWorks_ActiveModelDocChangeNotify
            RemoveHandler iSwApp.FileOpenPostNotify, AddressOf Me.SldWorks_FileOpenPostNotify
        Catch e As Exception
            Console.WriteLine(e.Message)
        End Try
    End Sub

    Sub AttachEventsToAllDocuments()
        Dim modDoc As ModelDoc2
        modDoc = iSwApp.GetFirstDocument()
        While Not modDoc Is Nothing
            If Not openDocs.Contains(modDoc) Then
                AttachModelDocEventHandler(modDoc)
            Else
                Dim docHandler As DocumentEventHandler = openDocs(modDoc)
                If Not docHandler Is Nothing Then
                    docHandler.ConnectModelViews()
                End If
            End If
            modDoc = modDoc.GetNext()
        End While
    End Sub

    Function AttachModelDocEventHandler(ByVal modDoc As ModelDoc2) As Boolean
        If modDoc Is Nothing Then
            Return False
        End If
        Dim docHandler As DocumentEventHandler = Nothing

        If Not openDocs.Contains(modDoc) Then
            Select Case modDoc.GetType
                Case swDocumentTypes_e.swDocPART
                    docHandler = New PartEventHandler()
                Case swDocumentTypes_e.swDocASSEMBLY
                    docHandler = New AssemblyEventHandler()
                Case swDocumentTypes_e.swDocDRAWING
                    docHandler = New DrawingEventHandler()
            End Select

            docHandler.Init(iSwApp, Me, modDoc)
            docHandler.AttachEventHandlers()
            openDocs.Add(modDoc, docHandler)
        End If
    End Function

    Sub DetachModelEventHandler(ByVal modDoc As ModelDoc2)
        Dim docHandler As DocumentEventHandler
        docHandler = openDocs.Item(modDoc)
        openDocs.Remove(modDoc)
        modDoc = Nothing
        docHandler = Nothing
    End Sub
#End Region

#Region "事件处理程序"
    Function SldWorks_ActiveDocChangeNotify() As Integer
        'TODO: Add your implementation here
    End Function

    Function SldWorks_DocumentLoadNotify2(ByVal docTitle As String, ByVal docPath As String) As Integer

    End Function

    Function SldWorks_FileNewNotify2(ByVal newDoc As Object, ByVal doctype As Integer, ByVal templateName As String) As Integer
        AttachEventsToAllDocuments()
    End Function

    Function SldWorks_ActiveModelDocChangeNotify() As Integer
        'TODO: Add your implementation here
    End Function

    Function SldWorks_FileOpenPostNotify(ByVal FileName As String) As Integer
        AttachEventsToAllDocuments()
    End Function
#End Region

#Region "UI 回调"
    Sub CreateCube()

        'make sure we have a part open
        Dim partTemplate As String
        Dim model As ModelDoc2
        Dim featMan As FeatureManager

        partTemplate = iSwApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)
        If Not partTemplate = "" Then
            model = iSwApp.NewDocument(partTemplate, swDwgPaperSizes_e.swDwgPaperA2size, 0.0, 0.0)

            model.InsertSketch2(True)
            model.SketchRectangle(0, 0, 0, 0.1, 0.1, 0.1, False)

            'Extrude the sketch
            featMan = model.FeatureManager
            featMan.FeatureExtrusion(True,
                                      False, False,
                                      swEndConditions_e.swEndCondBlind, swEndConditions_e.swEndCondBlind,
                                      0.1, 0.0,
                                      False, False,
                                      False, False,
                                      0.0, 0.0,
                                      False, False,
                                      False, False,
                                      True,
                                      False, False)
        Else
            System.Windows.Forms.MessageBox.Show("There is no part template available. Please check your options and make sure there is a part template selected, or select a new part template.")
        End If
    End Sub

    Sub PopupCallbackFunction()
        bRet = iSwApp.ShowThirdPartyPopupMenu(registerID, 500, 500)
    End Sub

    Function PopupEnable() As Integer
        If iSwApp.ActiveDoc Is Nothing Then
            PopupEnable = 0
        Else
            PopupEnable = 1
        End If
    End Function

    Sub TestCallback()
        Debug.Print("Test callback")
    End Sub
    Function EnableTest() As Integer
        If iSwApp.ActiveDoc Is Nothing Then
            EnableTest = 0
        Else
            EnableTest = 1
        End If
    End Function

    Sub ShowPMP()
        If Not ppage Is Nothing Then
            ppage.Show()
        End If
    End Sub

    Function PMPEnable() As Integer
        If iSwApp.ActiveDoc Is Nothing Then
            PMPEnable = 0
        Else
            PMPEnable = 1
        End If
    End Function

    Sub FlyoutCallback()

        Dim flyGroup As FlyoutGroup = iCmdMgr.GetFlyoutGroup(flyoutGroupID)
        flyGroup.RemoveAllCommandItems()

        flyGroup.AddCommandItem(System.DateTime.Now.ToLongTimeString(), "test", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1")

    End Sub

    Function FlyoutEnable() As Integer
        Return 1
    End Function

    Sub FlyoutCommandItem1()
        iSwApp.SendMsgToUser("Flyout command 1")
    End Sub

    Function FlyoutEnableCommandItem1() As Integer
        Return 1
    End Function


#End Region

End Class

