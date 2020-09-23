Attribute VB_Name = "modForm"
Option Explicit

Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)

Private Const MF_BYPOSITION             As Long = &H400&
Private Const MF_REMOVE                 As Long = &H1000&
Private Const HWND_TOPMOST              As Long = -1
Private Const HWND_NOTOPMOST            As Long = -2
Private Const SWP_NOSIZE                As Long = &H1
Private Const SWP_NOMOVE                As Long = &H2
Private Const SWP_NOACTIVATE            As Long = &H10
Private Const SWP_SHOWWINDOW            As Long = &H40

'=============================================================================
'                     Form-related Routines
'=============================================================================

'-----------------------------------------------------------------------------
Public Sub CenterForm(pobjForm As Form)
'-----------------------------------------------------------------------------

    With pobjForm
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With

  End Sub

'-----------------------------------------------------------------------------
Public Function FormIsLoaded(pstrFormName As String) As Boolean
'-----------------------------------------------------------------------------
    
    Dim objForm As Form
    
    For Each objForm In Forms
        If objForm.Name = pstrFormName Then
            FormIsLoaded = True
            Exit Function
        End If
    Next
    
    FormIsLoaded = False

End Function

'-----------------------------------------------------------------------------
Public Sub DisableFormXButton(pobjForm As Form)
'-----------------------------------------------------------------------------

    Dim hSysMenu As Long
    Dim nCnt     As Long
    
    ' Get handle To our form's system menu
    ' (Restore, Maximize, Move, Close etc.)
    
    hSysMenu = GetSystemMenu(pobjForm.hwnd, False)

    If hSysMenu Then
        ' Get System menu's menu count
        nCnt = GetMenuItemCount(hSysMenu)
        If nCnt Then
            ' Menu count is based On 0 (0, 1, 2, 3...)
            RemoveMenu hSysMenu, nCnt - 1, _
                       MF_BYPOSITION Or MF_REMOVE
            RemoveMenu hSysMenu, nCnt - 2, _
                       MF_BYPOSITION Or MF_REMOVE ' Remove the seperator
            DrawMenuBar pobjForm.hwnd
        End If
    End If

End Sub

'-----------------------------------------------------------------------------
Public Sub LockWindow(hwnd As Long)
'-----------------------------------------------------------------------------
    LockWindowUpdate hwnd
End Sub

'-----------------------------------------------------------------------------
Public Sub UnlockWindow()
'-----------------------------------------------------------------------------
    LockWindowUpdate 0
End Sub

'-----------------------------------------------------------------------------
Public Sub MakeTopmost(pobjForm As Form, pblnMakeTopmost As Boolean)
'-----------------------------------------------------------------------------

    Dim lngParm As Long
    
    lngParm = IIf(pblnMakeTopmost, HWND_TOPMOST, HWND_NOTOPMOST)
    
    SetWindowPos pobjForm.hwnd, _
                 lngParm, _
                 0, _
                 0, _
                 0, _
                 0, _
                 (SWP_NOACTIVATE Or SWP_SHOWWINDOW Or _
                  SWP_NOMOVE Or SWP_NOSIZE)

End Sub



