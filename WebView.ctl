VERSION 5.00
Begin VB.UserControl WebView 
   BackColor       =   &H00F3F3F3&
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3330
   ForeColor       =   &H00333333&
   ScaleHeight     =   185
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   222
End
Attribute VB_Name = "WebView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' O    O    O
'  \__/ \__/
'  /=||=||=\   oouuuunnnnnnnnmmmmmmmmmmmmmm\
' // ||_||               WebView Control    \
' \\ /\ #\     oouuuunnnnnnnnmmmmmmmmmmmmmmmm\
' /=(  \  )==>       Coded by EBArtSoft@      \
'//  \O_\/         Copyright © 2006 - 2025     \
'\\  || ||        https://www.ebartsoft.com     \
' \==||=||==/  oouuuunnnnnnnnmmmmmmmmmmmmmmmmmmmm\
' ===========
'==== E.B ====
'
' ALL RIGHTS RESERVED ::..
' Permission  to  use,  copy,  modify,  and  distribute this software for
' any  purpose and  without  fee  is  hereby  granted,  provided that the
' above copyright notice appear in all copies and that both the copyright
' notice  and  this permission notice appear in supporting documentation.
'
' THE  MATERIAL  EMBODIED  ON  THIS  SOFTWARE IS PROVIDED TO YOU "AS-IS"
' AND  WITHOUT  WARRANTY  OF  ANY  KIND,  EXPRESS, IMPLIED OR OTHERWISE,
' INCLUDING  WITHOUT  LIMITATION,  ANY  WARRANTY  OF  MERCHANTABILITY OR
' FITNESS  FOR  A  PARTICULAR  PURPOSE.  IN  NO EVENT SHALL WE BE LIABLE
' TO  YOU  OR  ANYONE ELSE FOR ANY DIRECT, SPECIAL, INCIDENTAL, INDIRECT
' OR  CONSEQUENTIAL  DAMAGES  OF  ANY  KIND,  OR ANY DAMAGES WHATSOEVER,
' INCLUDING  WITHOUT  LIMITATION,  LOSS  OF PROFIT, LOSS OF USE, SAVINGS
' OR REVENUE, OR THE CLAIMS OF THIRD PARTIES, WHETHER OR NOT WE HAS BEEN
' ADVISED  OF  THE  POSSIBILITY  OF  SUCH  LOSS,  HOWEVER  CAUSED AND ON
' ANY  THEORY  OF  LIABILITY,  ARISING  OUT OF OR IN CONNECTION WITH THE
' POSSESSION, USE OR PERFORMANCE OF THIS SOFTWARE.
'
' Addendum: Use, copy, modify and distribute as you will
'
Option Explicit

'================================================
' PROP DEFAULT VALUES
'================================================
Private Const PROP_USELOADER        As Boolean = True
Private Const PROP_USERLIBPATH      As String = "%app.path%\microsoft.web.webview2\EBWebView\x86\EmbeddedBrowserWebView.dll"
Private Const PROP_USERDATAFOLDER   As String = "%app.path%\userdata"

'================================================
' API
'================================================
Private Declare Function CreateWebViewEnvironmentWithOptionsInternal Lib "EmbeddedBrowserWebView.dll" (ByVal pCheckRunningInstance As Long, ByVal pRuntimeType As Long, ByVal pUserDataFolder As Long, ByVal pEnvironmentOptions As Long, ByVal pEnvCompletedHandler As ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler) As Long
Private Declare Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" (ByVal pBrowserExecutableFolder As Long, ByVal pUserDataFolder As Long, ByVal pEnvironmentOptions As Long, ByVal pEnvCompletedHandler As ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler) As Long

'================================================
' EVENTS
'================================================
Public Event OnNavigationStarting(ByVal pIsRedirected As Boolean, ByVal pIsUserInitiated As Boolean, ByVal pNavigationId As Currency, ByVal p As ICoreWebView2HttpRequestHeaders, ByVal pUri As String, ByRef pCancel As Boolean)
Public Event OnPermissionRequested(ByRef pHandled As Boolean, ByVal pIsUserInitiated As Boolean, ByVal pPermissionKind As COREWEBVIEW2_PERMISSION_KIND, ByRef pUri As String, ByRef pState As COREWEBVIEW2_PERMISSION_STATE)
Public Event OnScriptDialogOpening(ByRef pAccept As Boolean, ByVal pKind As COREWEBVIEW2_SCRIPT_DIALOG_KIND, ByRef pUri As String, ByRef pDefaultValue As String, ByRef pMessage As String, ByRef pResultTest As String)
Public Event OnAcceleratorKeyPressed(ByRef pHandled As Boolean, ByVal pKeyEventKind As COREWEBVIEW2_KEY_EVENT_KIND, ByVal pKeyEventLParam As Long, ByVal pVirtualKey As Long, ByVal pPhysicalKeyStatus As Long)
Public Event OnNewWindowRequested(ByRef pHandled As Boolean, ByVal pIsUserInitiated As Boolean, ByRef pNewWindow As ICoreWebView2, ByVal pUri As String, ByVal pWindowFeatures As ICoreWebView2WindowFeatures)
Public Event OnWebResourceRequested(ByVal pResourceContext As COREWEBVIEW2_WEB_RESOURCE_CONTEXT, ByVal pRequest As ICoreWebView2WebResourceRequest, ByRef pResponse As ICoreWebView2WebResourceResponse)
Public Event OnNavigationCompleted(ByVal pIsSuccess As Boolean, ByVal pNavigationId As Currency, ByVal pWebErrorStatus As COREWEBVIEW2_WEB_ERROR_STATUS)
Public Event OnCallDevToolsProtocolMethodCompleted(ByVal pErrorCode As Long, ByVal pReturnObjectAsJson As String)
Public Event OnMoveFocusRequested(ByRef pHandled As Boolean, ByVal pReason As COREWEBVIEW2_MOVE_FOCUS_REASON)
Public Event OnAddScriptToExecuteOnDocumentCreatedCompleted(ByVal pErrorCode As Long, ByVal pId As String)
Public Event OnPrintCompleted(ByVal pErrorCode As Long, ByVal pPrintStatus As COREWEBVIEW2_PRINT_STATUS)
Public Event OnExecuteScriptCompleted(ByVal pErrorCode As Long, ByVal pResultObjectAsJson As String)
Public Event OnContentLoading(ByVal pIsErrorPage As Boolean, ByVal pNavigationId As Currency)
Public Event OnWebMessageReceived(ByRef pSource As String, ByRef pWebMessageAsJson As String)
Public Event OnPrintToPdfCompleted(ByVal pErrorCode As Long, ByVal pIsSuccessful As Boolean)
Public Event OnProcessFailed(ByVal pProcessFailedKind As COREWEBVIEW2_PROCESS_FAILED_KIND)
Public Event OnCapturePreviewCompleted(ByVal pErrorCode As Long, ByRef pData() As Byte)
Public Event OnSourceChanged(ByVal pIsNewDocument As Boolean, ByVal pSource As String)
Public Event OnNewBrowserVersionAvailable(ByVal pBrowserVersion As String)
Public Event OnDocumentTitleChanged(ByVal pDocumentTitle As String)
Public Event OnZoomFactorChanged(ByVal pZoomFactor As Double)
Public Event OnContainsFullScreenElementChanged()
Public Event OnWindowCloseRequested()
Public Event OnFocusChangedEvent()
Public Event OnHistoryChanged()
Public Event OnWebViewReady()

'================================================
' HANDLERS
'================================================
Implements ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler
Implements ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler
Implements ICoreWebView2ContainsFullScreenElementChangedEventHandler
Implements ICoreWebView2CreateCoreWebView2ControllerCompletedHandler
Implements ICoreWebView2CallDevToolsProtocolMethodCompletedHandler
Implements ICoreWebView2NewBrowserVersionAvailableEventHandler
Implements ICoreWebView2AcceleratorKeyPressedEventHandler
Implements ICoreWebView2WebResourceRequestedEventHandler
Implements ICoreWebView2WindowCloseRequestedEventHandler
Implements ICoreWebView2DocumentTitleChangedEventHandler
Implements ICoreWebView2PermissionRequestedEventHandler
Implements ICoreWebView2NavigationCompletedEventHandler
Implements ICoreWebView2ScriptDialogOpeningEventHandler
Implements ICoreWebView2NewWindowRequestedEventHandler
Implements ICoreWebView2MoveFocusRequestedEventHandler
Implements ICoreWebView2CapturePreviewCompletedHandler
Implements ICoreWebView2NavigationStartingEventHandler
Implements ICoreWebView2WebMessageReceivedEventHandler
Implements ICoreWebView2ZoomFactorChangedEventHandler
Implements ICoreWebView2ExecuteScriptCompletedHandler
Implements ICoreWebView2PrintToPdfCompletedHandler
Implements ICoreWebView2ContentLoadingEventHandler
Implements ICoreWebView2HistoryChangedEventHandler
Implements ICoreWebView2ProcessFailedEventHandler
Implements ICoreWebView2SourceChangedEventHandler
Implements ICoreWebView2FocusChangedEventHandler
Implements ICoreWebView2PrintCompletedHandler

'================================================
' PRIVATE VARS
'================================================
Private mEnvironment    As ICoreWebView2Environment
Private mController     As ICoreWebView2Controller
Private mWebView        As ICoreWebView2_17
Private mCaptureStream  As IStream
Private mDesignMode     As Boolean
Private mUseLoader      As Boolean
Private mUserDataFolder As String
Private mUserLibPath    As String
Private mLibHandle      As Long

'================================================
' UserControl_Initialize
'================================================
Private Sub UserControl_Initialize()
    'NTD
End Sub

'================================================
' UserControl_Terminate
'================================================
Private Sub UserControl_Terminate()
    WebView_Terminate
End Sub

'================================================
' UserControl_Resize
'================================================
Private Sub UserControl_Resize()
    WebView_Resize
End Sub

'================================================
' UserControl_Show
'================================================
Private Sub UserControl_Show()
    WebView_Resize
End Sub

'================================================
' UserControl_WriteProperties
'================================================
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "UseLoader", mUseLoader
    PropBag.WriteProperty "UserLibPath", mUserLibPath
    PropBag.WriteProperty "UserDataFolder", mUserDataFolder
End Sub

'================================================
' UserControl_InitProperties
'================================================
Private Sub UserControl_InitProperties()
    mDesignMode = Not UserControl.Ambient.UserMode
    mUseLoader = PROP_USELOADER
    mUserLibPath = PROP_USERLIBPATH
    mUserDataFolder = PROP_USERDATAFOLDER
    WebView_Initialize
End Sub

'================================================
' UserControl_ReadProperties
'================================================
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    mDesignMode = Not UserControl.Ambient.UserMode
    mUseLoader = PropBag.ReadProperty("UseLoader", PROP_USELOADER)
    mUserLibPath = PropBag.ReadProperty("UserLibPath", PROP_USERLIBPATH)
    mUserDataFolder = PropBag.ReadProperty("UserDataFolder", PROP_USERDATAFOLDER)
    WebView_Initialize
End Sub

'/!//!//!//!//!//!//!//!//!//!//!//!//!//!//!//!/
'                 PROPERTIES
'/!//!//!//!//!//!//!//!//!//!//!//!//!//!//!//!/

'================================================
' LibHandle
'================================================
Public Property Get LibHandle() As Long
    LibHandle = GetModuleHandle("EmbeddedBrowserWebView.dll")
End Property

'================================================
' LibFileName
'================================================
Public Property Get LibFileName() As String
    Dim s As String
    Dim n As Long
    s = String$(260, 0)
    n = GetModuleFileName(LibHandle, s, Len(s))
    LibFileName = Left$(s, n)
End Property

'================================================
' UserLibPath
'================================================
Public Property Get UserLibPath() As String
    UserLibPath = mUserLibPath
End Property

'================================================
' UserLibPath
'================================================
Public Property Let UserLibPath(ByRef v As String)
    If mDesignMode = False Then Exit Property
    mUserLibPath = v
End Property

'================================================
' UserDataFolder
'================================================
Public Property Get UserDataFolder() As String
    UserDataFolder = mUserDataFolder
End Property

'================================================
' UserDataFolder
'================================================
Public Property Let UserDataFolder(ByRef v As String)
    If mDesignMode = False Then Exit Property
    mUserDataFolder = v
End Property

'================================================
' UseLoader
'================================================
Public Property Get UseLoader() As Boolean
    UseLoader = mUseLoader
End Property

'================================================
' UseLoader
'================================================
Public Property Let UseLoader(ByVal v As Boolean)
    If mDesignMode = False Then Exit Property
    mUseLoader = v
End Property

'================================================
' WebView
'================================================
Public Property Get WebView() As ICoreWebView2_17
    Set WebView = mWebView
End Property

'================================================
' Controller
'================================================
Public Property Get Controller() As ICoreWebView2Controller
    Set Controller = mController
End Property

'================================================
' Environment
'================================================
Public Property Get Environment() As ICoreWebView2Environment
    Set Environment = mEnvironment
End Property

'================================================
' ZoomFactor
'================================================
Public Property Get ZoomFactor() As Double
    If mController Is Nothing Then Exit Property
    ZoomFactor = mController.ZoomFactor
End Property

'================================================
' ZoomFactor
'================================================
Public Property Let ZoomFactor(ByVal v As Double)
    If mController Is Nothing Then Exit Property
    mController.ZoomFactor = v
End Property

'================================================
' Source
'================================================
Public Property Get Source() As String
    If mWebView Is Nothing Then Exit Property
    Source = LPWSTR(mWebView.Source)
End Property

'================================================
' DocumentTitle
'================================================
Public Property Get DocumentTitle() As String
    If mWebView Is Nothing Then Exit Property
    DocumentTitle = LPWSTR(mWebView.DocumentTitle)
End Property

'/!//!//!//!//!//!//!//!//!//!//!//!//!//!//!//!/
'               PRIVATE METHODS
'/!//!//!//!//!//!//!//!//!//!//!//!//!//!//!//!/

'================================================
' WebView_Initialize
'================================================
Private Sub WebView_Initialize()
    Dim vUserDataFolder As String
    Dim vUserLibPath    As String
    If mDesignMode = False Then
        vUserDataFolder = Replace$(mUserDataFolder, "%app.path%", App.Path, 1, -1, vbTextCompare)
        If mUseLoader Then
            mLibHandle = LoadLibrary("WebView2Loader.dll")
            If mLibHandle Then
                If CreateCoreWebView2EnvironmentWithOptions(StrPtr(vbNullString), StrPtr(vUserDataFolder), 0, Me) = 0 Then
                    Exit Sub
                End If
            End If
        Else
            vUserLibPath = Replace$(mUserLibPath, "%app.path%", App.Path, 1, -1, vbTextCompare)
            mLibHandle = LoadLibrary(vUserLibPath)
            If mLibHandle Then
                If CreateWebViewEnvironmentWithOptionsInternal(True, 0, StrPtr(vUserDataFolder), 0, Me) = 0 Then
                    Exit Sub
                End If
            End If
        End If
    End If
End Sub

'================================================
' WebView_Terminate
'================================================
Private Sub WebView_Terminate()
    Set mEnvironment = Nothing
    Set mController = Nothing
    Set mWebView = Nothing
    If mLibHandle Then
        FreeLibrary mLibHandle
        mLibHandle = 0
    End If
End Sub

'================================================
' WebView_Resize
'================================================
Private Sub WebView_Resize()
    If mDesignMode Then
        WebView_DrawDesignMode
    Else
        WebView_DrawUserMode
    End If
End Sub

'================================================
' WebView_AddEventHandlers
'================================================
Private Sub WebView_AddEventHandlers()
    With mEnvironment
        .add_NewBrowserVersionAvailable Me, 0
    End With
    With mController
        .add_AcceleratorKeyPressed Me, 0
        .add_MoveFocusRequested Me, 0
        .add_ZoomFactorChanged Me, 0
        .add_LostFocus Me, 0
        .add_GotFocus Me, 0
    End With
    With mWebView
        .add_ContainsFullScreenElementChanged Me, 0
        .add_ContentLoading Me, 0
        .add_DocumentTitleChanged Me, 0
        .add_FrameNavigationCompleted Me, 0
        .add_FrameNavigationStarting Me, 0
        .add_HistoryChanged Me, 0
        .add_NavigationCompleted Me, 0
        .add_NavigationStarting Me, 0
        .add_NewWindowRequested Me, 0
        .add_PermissionRequested Me, 0
        .add_ProcessFailed Me, 0
        .add_ScriptDialogOpening Me, 0
        .add_SourceChanged Me, 0
        .add_WebMessageReceived Me, 0
        .add_WebResourceRequested Me, 0
        .add_WindowCloseRequested Me, 0
    End With
End Sub

'================================================
' WebView_DrawDesignMode
'================================================
Private Sub WebView_DrawDesignMode()
    Dim s As String
    s = "Design Mode"
    AutoRedraw = True
    Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), BackColor, BF
    Line (0, 0)-(ScaleWidth - 1, ScaleHeight - 1), ForeColor, B
    CurrentY = Int((ScaleHeight - TextHeight(s)) * 0.5)
    CurrentX = Int((ScaleWidth - TextWidth(s)) * 0.5)
    Print s
    AutoRedraw = False
End Sub

'================================================
' WebView_DrawUserMode
'================================================
Private Sub WebView_DrawUserMode()
    Dim vRect As RECT
    If mController Is Nothing Then Exit Sub
    GetClientRect UserControl.hwnd, vRect
    mController.SetBounds vRect.Left, vRect.Top, vRect.Right, vRect.Bottom
End Sub

'/!//!//!//!//!//!//!//!//!//!//!//!//!//!//!//!/
'               PUBLIC METHODS
'/!//!//!//!//!//!//!//!//!//!//!//!//!//!//!//!/

'================================================
' GoForward
'================================================
Public Sub Reload()
    If mWebView Is Nothing Then Exit Sub
    mWebView.Reload
End Sub

'================================================
' GoBack
'================================================
Public Sub GoBack()
    If mWebView Is Nothing Then Exit Sub
    mWebView.GoBack
End Sub

'================================================
' GoForward
'================================================
Public Sub GoForward()
    If mWebView Is Nothing Then Exit Sub
    mWebView.GoForward
End Sub

'================================================
' OpenDevToolsWindow
'================================================
Public Sub OpenDevToolsWindow()
    If mWebView Is Nothing Then Exit Sub
    mWebView.OpenDevToolsWindow
End Sub

'================================================
' OpenTaskManagerWindow
'================================================
Public Sub OpenTaskManagerWindow()
    If mWebView Is Nothing Then Exit Sub
    mWebView.OpenTaskManagerWindow
End Sub

'================================================
' ShowPrintUI
'================================================
Public Sub ShowPrintUI(ByVal pPrintDialogKind As COREWEBVIEW2_PRINT_DIALOG_KIND)
    If mWebView Is Nothing Then Exit Sub
    mWebView.ShowPrintUI pPrintDialogKind
End Sub

'================================================
' Navigate
'================================================
Public Sub Navigate(ByRef pUri As String)
    If mWebView Is Nothing Then Exit Sub
    mWebView.Navigate pUri
End Sub

'================================================
' NavigateToString
'================================================
Public Sub NavigateToString(ByRef pHtmlContent As String)
    If mWebView Is Nothing Then Exit Sub
    mWebView.NavigateToString pHtmlContent
End Sub

'================================================
' PostWebMessageAsJson
'================================================
Public Sub PostWebMessageAsJson(ByRef webMessageAsJson As String)
    If mWebView Is Nothing Then Exit Sub
    mWebView.PostWebMessageAsJson webMessageAsJson
End Sub

'================================================
' PostWebMessageAsString
'================================================
Public Sub PostWebMessageAsString(ByRef webMessageAsString As String)
    If mWebView Is Nothing Then Exit Sub
    mWebView.PostWebMessageAsString webMessageAsString
End Sub

'================================================
' ExecuteScript
'================================================
Public Sub ExecuteScript(ByRef pJavaScript As String, Optional ByVal pHandler As ICoreWebView2ExecuteScriptCompletedHandler = Nothing)
    If mWebView Is Nothing Then Exit Sub
    If pHandler Is Nothing Then Set pHandler = Me
    mWebView.ExecuteScript pJavaScript, pHandler
End Sub

'================================================
' AddHostObjectToScript
'================================================
Public Sub AddHostObjectToScript(ByRef pName As String, ByVal pObject As Object)
    If mWebView Is Nothing Then Exit Sub
    mWebView.AddHostObjectToScript pName, pObject
End Sub

'================================================
' RemoveHostObjectFromScript
'================================================
Public Sub RemoveHostObjectFromScript(ByRef pName As String)
    If mWebView Is Nothing Then Exit Sub
    mWebView.RemoveHostObjectFromScript pName
End Sub

'================================================
' AddScriptToExecuteOnDocumentCreated
'================================================
Public Sub AddScriptToExecuteOnDocumentCreated(ByRef pJavaScript As String, Optional ByVal pHandler As ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler = Nothing)
    If mWebView Is Nothing Then Exit Sub
    If pHandler Is Nothing Then Set pHandler = Me
    mWebView.AddScriptToExecuteOnDocumentCreated pJavaScript, pHandler
End Sub

'================================================
' RemoveScriptToExecuteOnDocumentCreated
'================================================
Public Sub RemoveScriptToExecuteOnDocumentCreated(ByRef pId As String)
    If mWebView Is Nothing Then Exit Sub
    mWebView.RemoveScriptToExecuteOnDocumentCreated pId
End Sub

'================================================
' AddWebResourceRequestedFilter
'================================================
Public Sub AddWebResourceRequestedFilter(ByRef pUri As String, ByVal pResourceContext As COREWEBVIEW2_WEB_RESOURCE_CONTEXT)
    If mWebView Is Nothing Then Exit Sub
    mWebView.AddWebResourceRequestedFilter pUri, pResourceContext
End Sub

'================================================
' RemoveWebResourceRequestedFilter
'================================================
Public Sub RemoveWebResourceRequestedFilter(ByRef pUri As String, ByVal pResourceContext As COREWEBVIEW2_WEB_RESOURCE_CONTEXT)
    If mWebView Is Nothing Then Exit Sub
    mWebView.RemoveWebResourceRequestedFilter pUri, pResourceContext
End Sub

'================================================
' CallDevToolsProtocolMethod
'================================================
Public Sub CallDevToolsProtocolMethod(ByRef pMethodName As String, ByRef pParametersAsJson As String, Optional ByVal pHandler As ICoreWebView2CallDevToolsProtocolMethodCompletedHandler = Nothing)
    If mWebView Is Nothing Then Exit Sub
    If pHandler Is Nothing Then Set pHandler = Me
    mWebView.CallDevToolsProtocolMethod pMethodName, pParametersAsJson, pHandler
End Sub

'================================================
' CapturePreview
'================================================
Public Sub CapturePreview(ByVal pFormat As COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT, Optional ByVal pHandler As ICoreWebView2CapturePreviewCompletedHandler = Nothing)
    If mWebView Is Nothing Then Exit Sub
    CreateStreamOnHGlobal 0, True, mCaptureStream
    If pHandler Is Nothing Then Set pHandler = Me
    mWebView.CapturePreview pFormat, mCaptureStream, pHandler
End Sub

'================================================
' Print
'================================================
Public Sub Print_(Optional ByVal pPrintSettings As ICoreWebView2PrintSettings2 = Nothing)
    If mWebView Is Nothing Then Exit Sub
    mWebView.Print pPrintSettings, Me
End Sub

'================================================
' PrintToPdf
'================================================
Public Sub PrintToPdf(ByRef pResultFilePath As String, Optional ByVal pPrintSettings As ICoreWebView2PrintSettings2 = Nothing)
    If mWebView Is Nothing Then Exit Sub
    mWebView.PrintToPdf pResultFilePath, pPrintSettings, Me
End Sub

'/!//!//!//!//!//!//!//!//!//!//!//!//!//!//!//!/
'                   UTILS
'/!//!//!//!//!//!//!//!//!//!//!//!//!//!//!//!/

'================================================
' LPWSTR >> https://learn.microsoft.com/en-us/microsoft-edge/webview2/concepts/win32-api-conventions
'================================================
Public Function LPWSTR(ByVal p As Long, Optional ByVal pMemFree As Boolean = True) As String
    Dim n As Long
    n = lstrlenW(p)
    If p <= 0 Then Exit Function
    If IsBadReadPtr(ByVal p, n) = 0 Then
        LPWSTR = String(n, 0)
        MemoryCopy ByVal StrPtr(LPWSTR), ByVal p, n * 2
    End If
    If pMemFree Then
        CoTaskMemFree p
    End If
End Function

'/!//!//!//!//!//!//!//!//!//!//!//!//!//!//!//!/
'                 EVENT HANDLERS
'/!//!//!//!//!//!//!//!//!//!//!//!//!//!//!//!/

'================================================
' ICoreWebView2CreateCoreWebView2ControllerCompletedHandler_Invoke
'================================================
Private Sub ICoreWebView2CreateCoreWebView2ControllerCompletedHandler_Invoke(ByVal errorCode As Long, ByVal createdController As Win32Tlb_Lib.ICoreWebView2Controller)
    'Debug.Print "ICoreWebView2CreateCoreWebView2ControllerCompletedHandler_Invoke"
    Set mController = createdController
    If mController Is Nothing Then Exit Sub
    Set mWebView = mController.CoreWebView2
    If mWebView Is Nothing Then Exit Sub
    WebView_AddEventHandlers
    WebView_Resize
    RaiseEvent OnWebViewReady
End Sub

'================================================
' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler_Invoke
'================================================
Private Sub ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler_Invoke(ByVal errorCode As Long, ByVal createdEnvironment As Win32Tlb_Lib.ICoreWebView2Environment)
    'Debug.Print "ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler_Invoke"
    Set mEnvironment = createdEnvironment
    If mEnvironment Is Nothing Then Exit Sub
    mEnvironment.CreateCoreWebView2Controller UserControl.hwnd, Me
End Sub

'================================================
' ICoreWebView2CapturePreviewCompletedHandler_Invoke
'================================================
Private Sub ICoreWebView2CapturePreviewCompletedHandler_Invoke(ByVal errorCode As Long)
    'Debug.Print "ICoreWebView2CapturePreviewCompletedHandler_Invoke"
    Dim vData() As Byte
    Dim vSize   As Long
    If errorCode = 0 Then
        vSize = mCaptureStream.Seek(0, STREAM_SEEK_END) * 10000
        If vSize > 0 Then
            ReDim vData(vSize - 1)
            mCaptureStream.Seek 0, STREAM_SEEK_SET
            mCaptureStream.Read vData(0), vSize, 0
        End If
    End If
    RaiseEvent OnCapturePreviewCompleted(errorCode, vData)
End Sub

'================================================
' ICoreWebView2DocumentTitleChangedEventHandler_Invoke
'================================================
Private Sub ICoreWebView2DocumentTitleChangedEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2, ByVal args As Long)
    'Debug.Print "ICoreWebView2DocumentTitleChangedEventHandler_Invoke"
    RaiseEvent OnDocumentTitleChanged(DocumentTitle)
End Sub

'================================================
' ICoreWebView2SourceChangedEventHandler_Invoke
'================================================
Private Sub ICoreWebView2SourceChangedEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2, ByVal args As Win32Tlb_Lib.ICoreWebView2SourceChangedEventArgs)
    'Debug.Print "ICoreWebView2SourceChangedEventHandler_Invoke"
    RaiseEvent OnSourceChanged(args.IsNewDocument, Source)
End Sub

'================================================
' ICoreWebView2ContentLoadingEventHandler_Invoke
'================================================
Private Sub ICoreWebView2ContentLoadingEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2, ByVal args As Win32Tlb_Lib.ICoreWebView2ContentLoadingEventArgs)
    'Debug.Print "ICoreWebView2ContentLoadingEventHandler_Invoke"
    RaiseEvent OnContentLoading(args.IsErrorPage, args.NavigationId)
End Sub

'================================================
' ICoreWebView2NavigationCompletedEventHandler_Invoke
'================================================
Private Sub ICoreWebView2NavigationCompletedEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2, ByVal args As Win32Tlb_Lib.ICoreWebView2NavigationCompletedEventArgs)
    'Debug.Print "ICoreWebView2NavigationCompletedEventHandler_Invoke"
    RaiseEvent OnNavigationCompleted(args.IsSuccess, args.NavigationId, args.WebErrorStatus)
End Sub

'================================================
' ICoreWebView2ContainsFullScreenElementChangedEventHandler_Invoke
'================================================
Private Sub ICoreWebView2ContainsFullScreenElementChangedEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2, ByVal args As Long)
    'Debug.Print "ICoreWebView2ContainsFullScreenElementChangedEventHandler_Invoke"
    RaiseEvent OnContainsFullScreenElementChanged
End Sub

'================================================
' ICoreWebView2HistoryChangedEventHandler_Invoke
'================================================
Private Sub ICoreWebView2HistoryChangedEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2, ByVal args As Long)
    'Debug.Print "ICoreWebView2HistoryChangedEventHandler_Invoke"
    RaiseEvent OnHistoryChanged
End Sub

'================================================
' ICoreWebView2NewBrowserVersionAvailableEventHandler_Invoke
'================================================
Private Sub ICoreWebView2NewBrowserVersionAvailableEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2Environment, ByVal args As Long)
    'Debug.Print "ICoreWebView2NewBrowserVersionAvailableEventHandler_Invoke"
    RaiseEvent OnNewBrowserVersionAvailable(LPWSTR(sender.BrowserVersionString))
End Sub

'================================================
' ICoreWebView2ZoomFactorChangedEventHandler_Invoke
'================================================
Private Sub ICoreWebView2ZoomFactorChangedEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2Controller, ByVal args As Long)
    'Debug.Print "ICoreWebView2ZoomFactorChangedEventHandler_Invoke"
    RaiseEvent OnZoomFactorChanged(sender.ZoomFactor)
End Sub

'================================================
' ICoreWebView2FocusChangedEventHandler_Invoke
'================================================
Private Sub ICoreWebView2FocusChangedEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2Controller, ByVal args As Long)
    'Debug.Print "ICoreWebView2FocusChangedEventHandler_Invoke"
    RaiseEvent OnFocusChangedEvent
End Sub

'================================================
' ICoreWebView2WindowCloseRequestedEventHandler_Invoke
'================================================
Private Sub ICoreWebView2WindowCloseRequestedEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2, ByVal args As Long)
    'Debug.Print "ICoreWebView2WindowCloseRequestedEventHandler_Invoke"
    RaiseEvent OnWindowCloseRequested
End Sub

'================================================
' ICoreWebView2ProcessFailedEventHandler_Invoke
'================================================
Private Sub ICoreWebView2ProcessFailedEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2, ByVal args As Win32Tlb_Lib.ICoreWebView2ProcessFailedEventArgs)
    'Debug.Print "ICoreWebView2ProcessFailedEventHandler_Invoke"
    RaiseEvent OnProcessFailed(args.ProcessFailedKind)
End Sub

'================================================
' ICoreWebView2CallDevToolsProtocolMethodCompletedHandler_Invoke
'================================================
Private Sub ICoreWebView2CallDevToolsProtocolMethodCompletedHandler_Invoke(ByVal errorCode As Long, ByVal returnObjectAsJson As Long)
    'Debug.Print "ICoreWebView2CallDevToolsProtocolMethodCompletedHandler_Invoke"
    RaiseEvent OnCallDevToolsProtocolMethodCompleted(errorCode, LPWSTR(returnObjectAsJson, False))
End Sub

'================================================
' ICoreWebView2ExecuteScriptCompletedHandler_Invoke
'================================================
Private Sub ICoreWebView2ExecuteScriptCompletedHandler_Invoke(ByVal errorCode As Long, ByVal resultObjectAsJson As Long)
    'Debug.Print "ICoreWebView2ExecuteScriptCompletedHandler_Invoke"
    RaiseEvent OnExecuteScriptCompleted(errorCode, LPWSTR(resultObjectAsJson))
End Sub

'================================================
' ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler_Invoke
'================================================
Private Sub ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler_Invoke(ByVal errorCode As Long, ByVal id As Long)
    'Debug.Print "ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler_Invoke"
    RaiseEvent OnAddScriptToExecuteOnDocumentCreatedCompleted(errorCode, LPWSTR(id))
End Sub

'================================================
' ICoreWebView2WebMessageReceivedEventHandler_Invoke
'================================================
Private Sub ICoreWebView2WebMessageReceivedEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2, ByVal args As Win32Tlb_Lib.ICoreWebView2WebMessageReceivedEventArgs)
    'Debug.Print "ICoreWebView2WebMessageReceivedEventHandler_Invoke"
    RaiseEvent OnWebMessageReceived(LPWSTR(args.Source), LPWSTR(args.webMessageAsJson))
End Sub

'================================================
' ICoreWebView2PrintCompletedHandler_Invoke
'================================================
Private Sub ICoreWebView2PrintCompletedHandler_Invoke(ByVal errorCode As Long, ByVal printStatus As Win32Tlb_Lib.COREWEBVIEW2_PRINT_STATUS)
    'Debug.Print "ICoreWebView2PrintCompletedHandler_Invoke"
    RaiseEvent OnPrintCompleted(errorCode, printStatus)
End Sub

'================================================
' ICoreWebView2PrintToPdfCompletedHandler_Invoke
'================================================
Private Sub ICoreWebView2PrintToPdfCompletedHandler_Invoke(ByVal errorCode As Long, ByVal isSuccessful As Long)
    'Debug.Print "ICoreWebView2PrintToPdfCompletedHandler_Invoke"
    RaiseEvent OnPrintToPdfCompleted(errorCode, isSuccessful)
End Sub

'================================================
' ICoreWebView2NavigationStartingEventHandler_Invoke
'================================================
Private Sub ICoreWebView2NavigationStartingEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2, ByVal args As Win32Tlb_Lib.ICoreWebView2NavigationStartingEventArgs)
    'Debug.Print "ICoreWebView2NavigationStartingEventHandler_Invoke"
    Dim vCancel As Boolean
    RaiseEvent OnNavigationStarting(args.IsRedirected, args.IsUserInitiated, args.NavigationId, args.RequestHeaders, LPWSTR(args.uri), vCancel)
    args.Cancel = vCancel
End Sub

'================================================
' ICoreWebView2AcceleratorKeyPressedEventHandler_Invoke
'================================================
Private Sub ICoreWebView2AcceleratorKeyPressedEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2Controller, ByVal args As Win32Tlb_Lib.ICoreWebView2AcceleratorKeyPressedEventArgs)
    'Debug.Print "ICoreWebView2AcceleratorKeyPressedEventHandler_Invoke"
    Dim vHandled As Boolean
    RaiseEvent OnAcceleratorKeyPressed(vHandled, args.KeyEventKind, args.KeyEventLParam, args.VirtualKey, VarPtr(args.PhysicalKeyStatus))
    args.Handled = vHandled
End Sub

'================================================
' ICoreWebView2MoveFocusRequestedEventHandler_Invoke
'================================================
Private Sub ICoreWebView2MoveFocusRequestedEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2Controller, ByVal args As Win32Tlb_Lib.ICoreWebView2MoveFocusRequestedEventArgs)
    'Debug.Print "ICoreWebView2MoveFocusRequestedEventHandler_Invoke"
    Dim vHandled As Boolean
    RaiseEvent OnMoveFocusRequested(vHandled, args.reason)
    args.Handled = vHandled
End Sub

'================================================
' ICoreWebView2WebResourceRequestedEventHandler_Invoke
'================================================
Private Sub ICoreWebView2WebResourceRequestedEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2, ByVal args As Win32Tlb_Lib.ICoreWebView2WebResourceRequestedEventArgs)
    'Debug.Print "ICoreWebView2WebResourceRequestedEventHandler_Invoke"
    Dim vResponse As ICoreWebView2WebResourceResponse
    RaiseEvent OnWebResourceRequested(args.ResourceContext, args.Request, vResponse)
    args.Response = vResponse
End Sub

'================================================
' ICoreWebView2PermissionRequestedEventHandler_Invoke
'================================================
Private Sub ICoreWebView2PermissionRequestedEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2, ByVal args As Win32Tlb_Lib.ICoreWebView2PermissionRequestedEventArgs)
    'Debug.Print "ICoreWebView2PermissionRequestedEventHandler_Invoke"
    Dim vArgs    As ICoreWebView2PermissionRequestedEventArgs3
    Dim vState   As COREWEBVIEW2_PERMISSION_STATE
    Dim vHandled As Boolean
    Set vArgs = args
    vState = vArgs.State
    RaiseEvent OnPermissionRequested(vHandled, vArgs.IsUserInitiated, vArgs.PermissionKind, LPWSTR(vArgs.uri), vState)
    vArgs.Handled = vHandled
    vArgs.State = vState
End Sub

'================================================
' ICoreWebView2NewWindowRequestedEventHandler_Invoke
'================================================
Private Sub ICoreWebView2NewWindowRequestedEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2, ByVal args As Win32Tlb_Lib.ICoreWebView2NewWindowRequestedEventArgs)
    'Debug.Print "ICoreWebView2NewWindowRequestedEventHandler_Invoke"
    Dim vNewWindow As ICoreWebView2
    Dim vHandled   As Boolean
    Set vNewWindow = args.NewWindow
    RaiseEvent OnNewWindowRequested(vHandled, args.IsUserInitiated, vNewWindow, LPWSTR(args.uri), args.WindowFeatures)
    args.NewWindow = vNewWindow
    args.Handled = vHandled
End Sub

'================================================
' ICoreWebView2ScriptDialogOpeningEventHandler_Invoke
'================================================
Private Sub ICoreWebView2ScriptDialogOpeningEventHandler_Invoke(ByVal sender As Win32Tlb_Lib.ICoreWebView2, ByVal args As Win32Tlb_Lib.ICoreWebView2ScriptDialogOpeningEventArgs)
    'Debug.Print "ICoreWebView2ScriptDialogOpeningEventHandler_Invoke"
    Dim vAccept     As Boolean
    Dim vResultText As String
    vResultText = args.ResultText_get
    RaiseEvent OnScriptDialogOpening(vAccept, args.Kind, LPWSTR(args.uri), LPWSTR(args.DefaultText), LPWSTR(args.Message), vResultText)
    args.ResultText_set vResultText
    If vAccept Then args.Accept
End Sub
