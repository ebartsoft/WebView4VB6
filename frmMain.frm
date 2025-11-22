VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "WebView Demo"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   630
   ClientWidth     =   9720
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   7.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   397
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   648
   StartUpPosition =   2  'CenterScreen
   Begin WebViewDemo.WebView WebView1 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4895
      UseLoader       =   -1  'True
      UserLibPath     =   "%app.path%\microsoft.web.webview2\EBWebView\x86\EmbeddedBrowserWebView.dll"
      UserDataFolder  =   "%app.path%\userdata"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "Navigation"
      Begin VB.Menu mnuNavigate 
         Caption         =   "Navigate to"
         Begin VB.Menu mnuNavigateToUrl 
            Caption         =   "Uri"
         End
         Begin VB.Menu mnuNavigateToHtml 
            Caption         =   "Html"
         End
         Begin VB.Menu mnuNavigateToFile 
            Caption         =   "File"
         End
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "History"
         Begin VB.Menu mnuGoBack 
            Caption         =   "Go back"
         End
         Begin VB.Menu mnuGoForward 
            Caption         =   "Go forward"
         End
      End
      Begin VB.Menu mnuReload 
         Caption         =   "Reload"
      End
   End
   Begin VB.Menu mnuSamples 
      Caption         =   "Samples"
      Begin VB.Menu mnuWebMessage 
         Caption         =   "Web Message"
         Begin VB.Menu mnuSamplesPostMessage 
            Caption         =   "1/3 - Receive Message"
            Index           =   0
         End
         Begin VB.Menu mnuSamplesPostMessage 
            Caption         =   "2/3 - Post Message (String)"
            Index           =   1
         End
         Begin VB.Menu mnuSamplesPostMessage 
            Caption         =   "3/3 - Post Message (Object)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuSamplesDialogs 
         Caption         =   "Dialogs"
      End
      Begin VB.Menu mnuSamplesPermissions 
         Caption         =   "Permissions"
      End
      Begin VB.Menu mnuSamplesNewWindow 
         Caption         =   "New Window"
      End
      Begin VB.Menu mnuSamplesRequestfilter 
         Caption         =   "Request Filter"
      End
      Begin VB.Menu mnuSamplesCookies 
         Caption         =   "Cookies"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuChangeZoomFactor 
         Caption         =   "Zoom Factor"
         Begin VB.Menu mnuZoomFactor 
            Caption         =   "Half"
            Index           =   0
         End
         Begin VB.Menu mnuZoomFactor 
            Caption         =   "Default"
            Index           =   1
         End
         Begin VB.Menu mnuZoomFactor 
            Caption         =   "Double"
            Index           =   2
         End
      End
      Begin VB.Menu mnuWindows 
         Caption         =   "Windows"
         Begin VB.Menu mnuOpenDevToolWindow 
            Caption         =   "Dev Tool"
         End
         Begin VB.Menu mnuOpenTaskManagerWindow 
            Caption         =   "Task Manager"
         End
         Begin VB.Menu mnuShowPrintUI 
            Caption         =   "Print Dialog"
         End
      End
      Begin VB.Menu mnuCapturePreview 
         Caption         =   "Capture Preview"
      End
      Begin VB.Menu mnuViewPrintToPDF 
         Caption         =   "Print to PDF"
      End
   End
   Begin VB.Menu mnuScript 
      Caption         =   "Script"
      Begin VB.Menu mnuInfoLibFileName 
         Caption         =   "Lib File Name"
      End
      Begin VB.Menu mnuExecuteScript 
         Caption         =   "Call JavaScript Function"
      End
      Begin VB.Menu mnuCallVBFunction 
         Caption         =   "Call Visual Basic Function"
      End
      Begin VB.Menu mnuCallDevToolsProtocolMethod 
         Caption         =   "Call Dev Tools Protocol Method"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuManual 
         Caption         =   "WebView2 Manual"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About WebView Demo"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Implements ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler
Implements ICoreWebView2ExecuteScriptCompletedHandler

'================================================
' Form_Resize
'================================================
Private Sub Form_Resize()
    On Error Resume Next
    WebView1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

'================================================
' mnuManual_Click
'================================================
Private Sub mnuManual_Click()
    ShellExecute 0, "open", "https://learn.microsoft.com/en-us/microsoft-edge/webview2/", vbNullString, vbNullString, vbNormalFocus
End Sub

'================================================
' mnuAbout_Click
'================================================
Private Sub mnuAbout_Click()
    MsgBox "WebView2 for Visual Basic 6" & vbCrLf & "By EBArtSoft@", vbInformation
End Sub

'================================================
' mnuZoomFactor_Click
'================================================
Private Sub mnuZoomFactor_Click(Index As Integer)
    Select Case Index
    Case 0: WebView1.ZoomFactor = 0.5
    Case 1: WebView1.ZoomFactor = 1
    Case 2: WebView1.ZoomFactor = 2
    End Select
End Sub

'================================================
' mnuGoBack_Click
'================================================
Private Sub mnuGoBack_Click()
    WebView1.GoBack
End Sub

'================================================
' mnuGoForward_Click
'================================================
Private Sub mnuGoForward_Click()
    WebView1.GoForward
End Sub

'================================================
' mnuReload_Click
'================================================
Private Sub mnuReload_Click()
    WebView1.Reload
End Sub

'================================================
' mnuShowPrintUI_Click
'================================================
Private Sub mnuShowPrintUI_Click()
    WebView1.ShowPrintUI COREWEBVIEW2_PRINT_DIALOG_KIND_BROWSER
End Sub

'================================================
' mnuViewPrintToPDF_Click
'================================================
Private Sub mnuViewPrintToPDF_Click()
    WebView1.PrintToPdf App.Path & "\print2pdf.pdf"
End Sub

'================================================
' mnuOpenTaskManagerWindow_Click
'================================================
Private Sub mnuOpenTaskManagerWindow_Click()
    WebView1.OpenTaskManagerWindow
End Sub

'================================================
' mnuOpenDevToolWindow_Click
'================================================
Private Sub mnuOpenDevToolWindow_Click()
    WebView1.OpenDevToolsWindow
End Sub

'================================================
' SampleFunction
'================================================
Public Function SampleFunction(ByVal s As String) As String
    SampleFunction = "Param = " & s
End Function

'================================================
' mnuCallVBFunction_Click
'================================================
Private Sub mnuCallVBFunction_Click()
    WebView1.AddHostObjectToScript "Form1", Me
    WebView1.ExecuteScript "alert(window.chrome.webview.hostObjects.sync.Form1.SampleFunction('hello'))", Me
End Sub

'================================================
' mnuExecuteScript_Click
'================================================
Private Sub mnuExecuteScript_Click()
    WebView1.ExecuteScript "alert('Hello from Javascript');'1+2='+(1+2);", Me
End Sub

'================================================
' mnuCapturePreview_Click
'================================================
Private Sub mnuCapturePreview_Click()
    WebView1.CapturePreview COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT_PNG
End Sub

'================================================
' mnuInfoLibFileName_Click
'================================================
Private Sub mnuInfoLibFileName_Click()
    MsgBox "0x" & Hex(WebView1.LibHandle) & vbCrLf & WebView1.LibFileName
End Sub

'================================================
' mnuNavigateToUrl_Click
'================================================
Private Sub mnuNavigateToUrl_Click()
    WebView1.WebView.Settings.AreDefaultScriptDialogsEnabled = True
    WebView1.Navigate "https://github.com/"
End Sub

'================================================
' mnuNavigateToHtml_Click
'================================================
Private Sub mnuNavigateToHtml_Click()
    WebView1.WebView.Settings.AreDefaultScriptDialogsEnabled = True
    WebView1.NavigateToString "<body style='background-color:#fff;text-align:center;'><u>Hello</u> <i>from</i> <b>HTML</b> string<br><a href='https://github.com/'>github.com</a></body>"
End Sub

'================================================
' mnuNavigateToFile_Click
'================================================
Private Sub mnuNavigateToFile_Click()
    WebView1.WebView.Settings.AreDefaultScriptDialogsEnabled = True
    WebView1.Navigate "file:///" & Replace$(App.Path, "\", "/") & "/www/hello.html"
End Sub

'================================================
' mnuSamplesDialogs_Click
'================================================
Private Sub mnuSamplesDialogs_Click()
    WebView1.WebView.Settings.AreDefaultScriptDialogsEnabled = False
    WebView1.Navigate "file:///" & Replace$(App.Path, "\", "/") & "/www/dialog.html"
End Sub

'================================================
' mnuSamplesRequestfilter_Click
'================================================
Private Sub mnuSamplesRequestfilter_Click()
    WebView1.Navigate "file:///" & Replace$(App.Path, "\", "/") & "/www/resource.html"
End Sub

'================================================
' mnuSamplesPostMessage_Click
'================================================
Private Sub mnuSamplesPostMessage_Click(Index As Integer)
    Select Case Index
    Case 0: WebView1.Navigate "file:///" & Replace$(App.Path, "\", "/") & "/www/postmsg.html"
    Case 1: WebView1.PostWebMessageAsString "Hello from Visual Basic"
    Case 2: WebView1.PostWebMessageAsJson "{""msg"":""Hello from Visual Basic""}"
    End Select
End Sub

'================================================
' mnuSamplesPermissions_Click
'================================================
Private Sub mnuSamplesPermissions_Click()
    WebView1.Navigate "file:///" & Replace$(App.Path, "\", "/") & "/www/cam.html"
End Sub

'================================================
' mnuSamplesNewWindow_Click
'================================================
Private Sub mnuSamplesNewWindow_Click()
    WebView1.Navigate "file:///" & Replace$(App.Path, "\", "/") & "/www/wnd.html"
End Sub

'================================================
' mnuSamplesCookies_Click
'================================================
Private Sub mnuSamplesCookies_Click()
    WebView1.Navigate "http://thissiteforcookies.com/"
End Sub

'================================================
' mnuCallDevToolsProtocolMethod_Click
'================================================
Private Sub mnuCallDevToolsProtocolMethod_Click()
    WebView1.CallDevToolsProtocolMethod "Page.captureScreenshot", "{""format"": ""jpeg""}"
End Sub

'================================================
' WebView1_OnPrintToPdfCompleted
'================================================
Private Sub WebView1_OnPrintToPdfCompleted(ByVal pErrorCode As Long, ByVal pIsSuccessful As Boolean)
    MsgBox "Print to PDF " & pErrorCode & " " & pIsSuccessful
End Sub

'================================================
' WebView1_OnWebViewReady
'================================================
Private Sub WebView1_OnWebViewReady()

    '------------------------------------------------------------------ <= WebView2 Settings
'    WebView1.WebView.Settings.AreDefaultScriptDialogsEnabled = False
'    WebView1.WebView.Settings.AreDefaultContextMenusEnabled = False
'    WebView1.WebView.Settings.AreDevToolsEnabled = False
'    WebView1.WebView.Settings.AreHostObjectsAllowed = False
'    WebView1.WebView.Settings.IsBuiltInErrorPageEnabled = False
'    WebView1.WebView.Settings.IsScriptEnabled = False
'    WebView1.WebView.Settings.IsStatusBarEnabled = False
'    WebView1.WebView.Settings.IsWebMessageEnabled = False
'    WebView1.WebView.Settings.IsZoomControlEnabled = False

    '------------------------------------------------------------------ <= Url filter
    WebView1.AddWebResourceRequestedFilter "http://thissiteforcookies.com/", COREWEBVIEW2_WEB_RESOURCE_CONTEXT_DOCUMENT
    WebView1.AddWebResourceRequestedFilter "http://thissitedosenotexist.com/", COREWEBVIEW2_WEB_RESOURCE_CONTEXT_DOCUMENT
    
    '------------------------------------------------------------------ <== Land page
    WebView1.NavigateToString "<head><title>Hello World</title></head><body style='background-color:#fff;text-align:center;'><br><h1>Hello World</h1><br>Select an action from the main menu</body>"
    
    '------------------------------------------------------------------ <== Preload Script
    WebView1.AddScriptToExecuteOnDocumentCreated "alert('Ok');", Me
    
    '------------------------------------------------------------------ <== Cookies
    Dim vCookie As ICoreWebView2Cookie
    Set vCookie = WebView1.WebView.CookieManager.CreateCookie("cookie1", "hello", "thissiteforcookies.com", "/")
    With vCookie
        .IsHttpOnly = False
        .IsSecure = False
    End With
    WebView1.WebView.CookieManager.AddOrUpdateCookie vCookie
    
End Sub

'================================================
' WebView1_OnCapturePreviewCompleted
'================================================
Private Sub WebView1_OnCapturePreviewCompleted(ByVal pErrorCode As Long, pData() As Byte)
    Dim vFileName As String
    vFileName = App.Path & "\capture.png"
    Open vFileName For Output As #1: Close #1
    Open vFileName For Binary As #1
    Put #1, , pData
    Close #1
    ShellExecute 0, "open", vFileName, vbNullString, vbNullString, vbNormalFocus
End Sub

'================================================
' WebView1_OnDocumentTitleChanged
'================================================
Private Sub WebView1_OnDocumentTitleChanged(ByVal pDocumentTitle As String)
    Caption = pDocumentTitle
End Sub

'================================================
' WebView1_OnZoomFactorChanged
'================================================
Private Sub WebView1_OnZoomFactorChanged(ByVal pZoomFactor As Double)
    Caption = "Zoom Factor " & pZoomFactor
End Sub

'================================================
' WebView1_OnNewBrowserVersionAvailable
'================================================
Private Sub WebView1_OnNewBrowserVersionAvailable(ByVal pBrowserVersion As String)
    Caption = "New Browser Version Available " & pBrowserVersion
End Sub

'================================================
' WebView1_OnAcceleratorKeyPressed
'================================================
Private Sub WebView1_OnAcceleratorKeyPressed(pHandled As Boolean, ByVal pKeyEventKind As Win32Tlb_Lib.COREWEBVIEW2_KEY_EVENT_KIND, ByVal pKeyEventLParam As Long, ByVal pVirtualKey As Long, ByVal pPhysicalKeyStatus As Long)
'    If pKeyEventKind = COREWEBVIEW2_KEY_EVENT_KIND_KEY_DOWN Then
'        If pVirtualKey = vbKeyF12 Then
'            If MsgBox("Show Dev Tool Window ?", vbYesNo Or vbQuestion) = vbNo Then
'                pHandled = True
'            End If
'        End If
'    End If
End Sub

'================================================
' WebView1_OnSourceChanged
'================================================
Private Sub WebView1_OnSourceChanged(ByVal pIsNewDocument As Boolean, ByVal pSource As String)
    'Debug.Print "WebView1_OnSourceChanged: " & pIsNewDocument & " " & pSource
End Sub

'================================================
' WebView1_OnCallDevToolsProtocolMethodCompleted
'================================================
Private Sub WebView1_OnCallDevToolsProtocolMethodCompleted(ByVal pErrorCode As Long, ByVal pReturnObjectAsJson As String)
    MsgBox pReturnObjectAsJson
End Sub

'================================================
' WebView1_OnWebMessageReceived
'================================================
Private Sub WebView1_OnWebMessageReceived(pSource As String, pWebMessageAsJson As String)
    MsgBox "Message : " & vbCrLf & pWebMessageAsJson & vbCrLf & vbCrLf & "From : " & vbCrLf & pSource, vbInformation
End Sub

'================================================
' WebView1_OnWebResourceRequested
'================================================
Private Sub WebView1_OnWebResourceRequested(ByVal pResourceContext As Win32Tlb_Lib.COREWEBVIEW2_WEB_RESOURCE_CONTEXT, ByVal pRequest As Win32Tlb_Lib.ICoreWebView2WebResourceRequest, ByRef pResponse As Win32Tlb_Lib.ICoreWebView2WebResourceResponse)
    Dim vUri    As String
    Dim vMethod As String
    Dim b()     As Byte
            
    vUri = WebView1.LPWSTR(pRequest.Uri_get)
    vMethod = WebView1.LPWSTR(pRequest.Method_get)
    If vUri = "http://thissitedosenotexist.com/" Then
        If MsgBox("Would you like to respond to '" & vUri & "'", vbYesNo Or vbQuestion) = vbYes Then
    
            Dim e  As ICoreWebView2HttpHeadersCollectionIterator
            Dim s  As IStream
            Dim e0 As Long
            Dim e1 As Long
            
            Set e = pRequest.Headers.GetIterator()
            While e.HasCurrentHeader
                e.GetCurrentHeader e0, e1
                Debug.Print WebView1.LPWSTR(e0) & " : " & WebView1.LPWSTR(e1)
                e.MoveNext
            Wend
        
            b = StrConv("<body><center>Hello from 'thissitedosenotexist.com'</center></body>", vbFromUnicode)
            CreateStreamOnHGlobal 0, True, s
            s.Write b(0), UBound(b) + 1, 0
            Set pResponse = WebView1.WebView.Environment.CreateWebResourceResponse(s, 200, "OK", "Content-Type: text/html")
    
        End If
        
    ElseIf vUri = "http://thissiteforcookies.com/" Then

        Open "C:\Users\Travail\Desktop\DEV\windows.tool.WebView\www\cookies.html" For Binary As #1
        ReDim b(LOF(1) - 1)
        Get #1, , b
        Close #1
        
        CreateStreamOnHGlobal 0, True, s
        s.Write b(0), UBound(b) + 1, 0
        Set pResponse = WebView1.WebView.Environment.CreateWebResourceResponse(s, 200, "OK", "Content-Type: text/html")
    
    End If
End Sub

'================================================
' WebView1_OnNewWindowRequested
'================================================
Private Sub WebView1_OnNewWindowRequested(pHandled As Boolean, ByVal pIsUserInitiated As Boolean, pNewWindow As Win32Tlb_Lib.ICoreWebView2, ByVal pUri As String, ByVal pWindowFeatures As Win32Tlb_Lib.ICoreWebView2WindowFeatures)
    Dim v As String
    v = "file:///" & Replace$(App.Path, "\", "/") & "/www/wnd.html"
    If StrComp(v, WebView1.Source, vbTextCompare) = 0 Then
        If MsgBox("Open new window ?", vbYesNo Or vbQuestion) = vbNo Then
            Set pNewWindow = WebView1.WebView
            pHandled = True
        End If
    Else
        Set pNewWindow = WebView1.WebView
        pHandled = True
    End If
End Sub

'================================================
' WebView1_OnPermissionRequested
'================================================
Private Sub WebView1_OnPermissionRequested(pHandled As Boolean, ByVal pIsUserInitiated As Boolean, ByVal pPermissionKind As Win32Tlb_Lib.COREWEBVIEW2_PERMISSION_KIND, pUri As String, pState As Win32Tlb_Lib.COREWEBVIEW2_PERMISSION_STATE)
    If pPermissionKind = COREWEBVIEW2_PERMISSION_KIND_CAMERA Then
        If MsgBox("Do you allow the use of the camera?", vbQuestion Or vbYesNo) = vbYes Then
            pState = COREWEBVIEW2_PERMISSION_STATE_ALLOW
        Else
            pState = COREWEBVIEW2_PERMISSION_STATE_DENY
        End If
        pHandled = True
    End If
End Sub

'================================================
' WebView1_OnScriptDialogOpening
'================================================
Private Sub WebView1_OnScriptDialogOpening(pAccept As Boolean, ByVal pKind As Win32Tlb_Lib.COREWEBVIEW2_SCRIPT_DIALOG_KIND, pUri As String, pDefaultValue As String, pMessage As String, pResultTest As String)
    Select Case pKind
    Case COREWEBVIEW2_SCRIPT_DIALOG_KIND_ALERT
        MsgBox pMessage, vbOKOnly Or vbInformation
        pAccept = True
    
    Case COREWEBVIEW2_SCRIPT_DIALOG_KIND_CONFIRM
        If MsgBox(pMessage, vbYesNo Or vbQuestion) = vbYes Then
            pAccept = True
        End If
    
    Case COREWEBVIEW2_SCRIPT_DIALOG_KIND_PROMPT
        pResultTest = InputBox(pMessage, , pDefaultValue)
        pAccept = Len(pResultTest)
        
    Case COREWEBVIEW2_SCRIPT_DIALOG_KIND_BEFOREUNLOAD
        If MsgBox("Changes you made may not be saved.", vbOKCancel Or vbQuestion, "Leave site?") = vbOK Then
            pAccept = True
        End If
        
    End Select
End Sub

'================================================
' ICoreWebView2ExecuteScriptCompletedHandler_Invoke
'================================================
Private Sub ICoreWebView2ExecuteScriptCompletedHandler_Invoke(ByVal errorCode As Long, ByVal resultObjectAsJson As Long)
    MsgBox WebView1.LPWSTR(resultObjectAsJson)
End Sub

'================================================
' ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler_Invoke
'================================================
Private Sub ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler_Invoke(ByVal errorCode As Long, ByVal id As Long)
    WebView1.RemoveScriptToExecuteOnDocumentCreated WebView1.LPWSTR(id)
End Sub

