Attribute VB_Name = "modExMessageBox"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2002 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' You are free to use this code within your own applications,
' but you are expressly forbidden from selling or otherwise
' distributing this source code without prior written consent.
' This includes both posting free demo projects made from this
' code as well as reproducing the code in text or html format.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim ExtraOptions As Integer

'uType parameters
Private Const MB_NOSOUND = &HF0&
Private Const MB_DEFBUTTON1 = &H0&
Private Const MB_DEFBUTTON2 = &H100&
Private Const MB_DEFBUTTON3 = &H200&
Private Const MB_ICONNONE = 0
Private Const MB_ICONCRITICAL = &H10&
Private Const MB_ICONQUESTION = &H20&
Private Const MB_ICONEXCLAMATION = &H30&
Private Const MB_ICONINFORMATION As Long = &H40&
Private Const MB_ABORTRETRYIGNORE As Long = &H2&
Private Const MB_TASKMODAL As Long = &H2000&

'Windows-defined Return values. The return
'values and control IDs are identical.
Public Const IDOK = 1
Public Const IDCANCEL = 2
Public Const IDABORT = 3
Public Const IDRETRY = 4
Public Const IDIGNORE = 5
Public Const IDYES = 6
Public Const IDNO = 7

'VBnet-defined control ID for the message prompt
Private Const IDPROMPT = &HFFFF&

'misc constants
Private Const WH_CBT = 5
Private Const GWL_HINSTANCE = (-6)
Private Const HCBT_ACTIVATE = 5

'UDT for passing data through the hook
Private Type MSGBOX_HOOK_PARAMS
  hwndOwner   As Long
  hHook       As Long
End Type

'need this declared at module level as
'it is used in the call and the hook proc
Private MSGHOOK As MSGBOX_HOOK_PARAMS

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Public Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" _
    (ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Private Declare Function MessageBox Lib "user32" _
    Alias "MessageBoxA" _
    (ByVal hwnd As Long, _
    ByVal lpText As String, _
    ByVal lpCaption As String, _
    ByVal wType As Long) As Long

Private Declare Function SetDlgItemText Lib "user32" _
    Alias "SetDlgItemTextA" _
    (ByVal hDlg As Long, _
    ByVal nIDDlgItem As Long, _
    ByVal lpString As String) As Long

Private Declare Function SetWindowsHookEx Lib "user32" _
    Alias "SetWindowsHookExA" _
    (ByVal idHook As Long, _
    ByVal lpfn As Long, _
    ByVal hmod As Long, _
    ByVal dwThreadId As Long) As Long

Private Declare Function SetWindowText Lib "user32" _
    Alias "SetWindowTextA" _
    (ByVal hwnd As Long, _
    ByVal lpString As String) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" _
    (ByVal hHook As Long) As Long



Public Function MessageBoxH(hwndThreadOwner As Long, hwndOwner As Long) As Long

  'Wrapper function for the MessageBox API

  Dim hInstance As Long
  Dim hThreadId As Long

  'Set up the CBT (computer-based training) hook
  hInstance = GetWindowLong(hwndThreadOwner, GWL_HINSTANCE)
  hThreadId = GetCurrentThreadId()

  'set up the MSGBOX_HOOK_PARAMS values
  'By specifying a Windows hook as one
  'of the params, we can intercept messages
  'sent by Windows and thereby manipulate
  'the dialog
  With MSGHOOK
    .hwndOwner = hwndOwner
    .hHook = SetWindowsHookEx(WH_CBT, _
        AddressOf MsgBoxHookProc, _
        hInstance, hThreadId)
  End With

  'call the MessageBox API and return the
  'value as the result of the function. The
  'Space$(120) statements assures the messagebox
  'is wide enough for the message that will
  'be set in the hook.
  '
  'NOTE: I am setting the text in the hook only
  'for demo purposes, to show how its done. You
  'certainly can pass the title and prompt text
  'right in the API call instead.

  If frmMain.optQuestion.Value = True Then
    'MessageBoxH = MessageBox(hwndOwner, _
     Space$(120), _
     Space$(120), _
     MB_ABORTRETRYIGNORE Or MB_ICONQUESTION)
    ExtraOptions = 0

    If frmMain.chkDontBeep.Value = 1 Then ExtraOptions = ExtraOptions + MB_NOSOUND

    If ExtraOptions <> 0 Then
      MessageBoxH = MessageBox(hwndOwner, _
          frmMain.txtMsgBoxMsg.Text, _
          frmMain.txtMsgBoxTitle.Text, _
          MB_ABORTRETRYIGNORE Or MB_ICONQUESTION Or ExtraOptions)
    Else
      MessageBoxH = MessageBox(hwndOwner, _
          frmMain.txtMsgBoxMsg.Text, _
          frmMain.txtMsgBoxTitle.Text, _
          MB_ABORTRETRYIGNORE Or MB_ICONQUESTION)
    End If
  ElseIf frmMain.optInformation.Value = True Then
    ExtraOptions = 0

    If frmMain.chkDontBeep.Value = 1 Then ExtraOptions = ExtraOptions + MB_NOSOUND

    If ExtraOptions <> 0 Then
      MessageBoxH = MessageBox(hwndOwner, _
          frmMain.txtMsgBoxMsg.Text, _
          frmMain.txtMsgBoxTitle.Text, _
          MB_ABORTRETRYIGNORE Or MB_ICONINFORMATION Or ExtraOptions)
    Else
      MessageBoxH = MessageBox(hwndOwner, _
          frmMain.txtMsgBoxMsg.Text, _
          frmMain.txtMsgBoxTitle.Text, _
          MB_ABORTRETRYIGNORE Or MB_ICONINFORMATION)
    End If

  ElseIf frmMain.optExclamation.Value = True Then
    ExtraOptions = 0

    If frmMain.chkDontBeep.Value = 1 Then ExtraOptions = ExtraOptions + MB_NOSOUND

    If ExtraOptions <> 0 Then
      MessageBoxH = MessageBox(hwndOwner, _
          frmMain.txtMsgBoxMsg.Text, _
          frmMain.txtMsgBoxTitle.Text, _
          MB_ABORTRETRYIGNORE Or MB_ICONEXCLAMATION Or ExtraOptions)
    Else
      MessageBoxH = MessageBox(hwndOwner, _
          frmMain.txtMsgBoxMsg.Text, _
          frmMain.txtMsgBoxTitle.Text, _
          MB_ABORTRETRYIGNORE Or MB_ICONEXCLAMATION)
    End If
  ElseIf frmMain.optCritical.Value = True Then
    ExtraOptions = 0

    If frmMain.chkDontBeep.Value = 1 Then ExtraOptions = ExtraOptions + MB_NOSOUND

    If ExtraOptions <> 0 Then
      MessageBoxH = MessageBox(hwndOwner, _
          frmMain.txtMsgBoxMsg.Text, _
          frmMain.txtMsgBoxTitle.Text, _
          MB_ABORTRETRYIGNORE Or MB_ICONCRITICAL Or ExtraOptions)
    Else
      MessageBoxH = MessageBox(hwndOwner, _
          frmMain.txtMsgBoxMsg.Text, _
          frmMain.txtMsgBoxTitle.Text, _
          MB_ABORTRETRYIGNORE Or MB_ICONCRITICAL)
    End If
  ElseIf frmMain.optIconNone.Value = True Then
    ExtraOptions = 0

    If frmMain.chkDontBeep.Value = 1 Then ExtraOptions = ExtraOptions + MB_NOSOUND

    If ExtraOptions <> 0 Then
      MessageBoxH = MessageBox(hwndOwner, _
          frmMain.txtMsgBoxMsg.Text, _
          frmMain.txtMsgBoxTitle.Text, _
          MB_ABORTRETRYIGNORE Or MB_ICONNONE Or ExtraOptions)
    Else
      MessageBoxH = MessageBox(hwndOwner, _
          frmMain.txtMsgBoxMsg.Text, _
          frmMain.txtMsgBoxTitle.Text, _
          MB_ABORTRETRYIGNORE Or MB_ICONNONE)
    End If
  End If
End Function


Public Function MsgBoxHookProc(ByVal uMsg As Long, _
      ByVal wParam As Long, _
      ByVal lParam As Long) As Long

  'When the message box is about to be shown,
  'we'll change the titlebar text, prompt message
  'and button captions
  If uMsg = HCBT_ACTIVATE Then

    'in a HCBT_ACTIVATE message, wParam holds
    'the handle to the messagebox

    'SetWindowText wParam, "Andy's MessageBox Hook Demo"
    SetWindowText wParam, frmMain.txtMsgBoxTitle

    'the ID's of the buttons on the message box
    'correspond exactly to the values they return,
    'so the same values can be used to identify
    'specific buttons in a SetDlgItemText call.

    'SetDlgItemText wParam, IDABORT, "Search C:\"
    'SetDlgItemText wParam, IDRETRY, "Search D:\"
    'SetDlgItemText wParam, IDIGNORE, "Cancel"

    SetDlgItemText wParam, IDABORT, frmMain.txtFirstButton.Text
    SetDlgItemText wParam, IDRETRY, frmMain.txtSecondButton.Text
    SetDlgItemText wParam, IDIGNORE, frmMain.txtThirdButton.Text

    'Change the dialog prompt text ...
    'SetDlgItemText wParam, IDPROMPT, "MyApp will now locate the application." & _
     "Please select the drive to search."

    SetDlgItemText wParam, IDPROMPT, frmMain.txtMsgBoxMsg

    'we're done with the dialog, so release the hook
    UnhookWindowsHookEx MSGHOOK.hHook

  End If

  'return False to let normal processing continue
  MsgBoxHookProc = False

End Function


