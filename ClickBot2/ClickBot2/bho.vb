
Imports System.Runtime.InteropServices
Imports Microsoft.Win32
Imports SHDocVw
Imports mshtml

''' <summary>
''' Set the GUID of this class and specify that this class is ComVisible.
''' A BHO must implement the interface IObjectWithSite. 
''' </summary>
<ComVisible(True), ClassInterface(ClassInterfaceType.None),
Guid("C42D40F0-BEBF-418D-8EA1-18D99AC2AB17")>
Public Class bho
    Implements IObjectWithSite


    ' To register a BHO, a new key should be created under this key.
    Private Const BHORegistryKey As String ="Software\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects"

    Private WebBrowser As WebBrowser
    Private Document As HTMLDocument


    Public Sub GetSite(ByRef riid As Guid, ByRef ppvSite As Object) Implements IObjectWithSite.GetSite

        Dim pUnk As IntPtr = Marshal.GetIUnknownForObject(WebBrowser)
        Dim hr As Integer = Marshal.QueryInterface(punk, riid, ppvSite)
        Marshal.Release(punk)

    End Sub

    Public Sub SetSite(pUnkSite As Object) Implements IObjectWithSite.SetSite

        If pUnkSite IsNot Nothing Then
            WebBrowser = DirectCast(pUnkSite, WebBrowser)
            AddHandler WebBrowser.DocumentComplete, AddressOf OnDocumentComplete
        Else
            RemoveHandler WebBrowser.DocumentComplete, AddressOf OnDocumentComplete
            WebBrowser = Nothing
        End If

    End Sub




#Region "Com Register/UnRegister Methods"
    ''' <summary>
    ''' When this class is registered to COM, add a new key to the BHORegistryKey 
    ''' to make IE use this BHO.
    ''' On 64bit machine, if the platform of this assembly and the installer is x86,
    ''' 32 bit IE can use this BHO. If the platform of this assembly and the installer
    ''' is x64, 64 bit IE can use this BHO.
    ''' </summary>
    <ComRegisterFunction()>
    Public Shared Sub RegisterBHO(ByVal t As Type)
        Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey(BHORegistryKey, True)
        If key Is Nothing Then
            key = Registry.LocalMachine.CreateSubKey(BHORegistryKey)
        End If

        ' 32 digits separated by hyphens, enclosed in braces: 
        ' {00000000-0000-0000-0000-000000000000}
        Dim bhoKeyStr As String = t.GUID.ToString("B")

        Dim bhoKey As RegistryKey = key.OpenSubKey(bhoKeyStr, True)

        ' Create a new key.
        If bhoKey Is Nothing Then
            bhoKey = key.CreateSubKey(bhoKeyStr)
        End If

        ' NoExplorer:dword = 1 prevents the BHO to be loaded by Explorer
        Dim name As String = "NoExplorer"
        Dim value As Object = CObj(1)
        bhoKey.SetValue(name, value)
        key.Close()
        bhoKey.Close()
    End Sub

    ''' <summary>
    ''' When this class is unregistered from COM, delete the key.
    ''' </summary>
    <ComUnregisterFunction()>
    Public Shared Sub UnregisterBHO(ByVal t As Type)
        Dim key As RegistryKey = Registry.LocalMachine.OpenSubKey(BHORegistryKey, True)
        Dim guidString As String = t.GUID.ToString("B")
        If key IsNot Nothing Then
            key.DeleteSubKey(guidString, False)
        End If
    End Sub
#End Region



    ''' <summary>
    ''' Handle the DocumentComplete event.
    ''' </summary>
    ''' <param name="pDisp">
    ''' The pDisp is an an object implemented the interface InternetExplorer.
    ''' By default, this object is the same as the ieInstance, but if the page 
    ''' contains many frames, each frame has its own document.
    ''' </param>
    Public Sub OnDocumentComplete(pDisp As Object, ByRef URL As Object)


        Select Case URL.ToString
            Case "about:blank"
                Exit Sub
        End Select

        Document = DirectCast(WebBrowser.Document, HTMLDocument)

        System.Windows.Forms.MessageBox.Show("Completed: " & URL.ToString, "Lobo's BHO")
        ' For Each tempElement As IHTMLInputElement In Document.getElementsByTagName("INPUT")
        'System.Windows.Forms.MessageBox.Show(If(tempElement.name IsNot Nothing, tempElement.name, "it sucks, no name, try id" + DirectCast(tempElement, IHTMLElement).id))
        'Next


    End Sub

End Class
