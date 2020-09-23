Attribute VB_Name = "modGeneral"
Option Explicit

Public gblnPopulating   As Boolean

'************************************************************************
'*          Global Variable and Constant Declarations                   *
'************************************************************************

'=============================================================================
'                     General Routines
'=============================================================================

'-----------------------------------------------------------------------------
Public Function GetAppPath() As String
'-----------------------------------------------------------------------------

    Dim strAppPath As String
    
    strAppPath = App.Path
    
    If Right$(strAppPath, 1) <> "\" Then
        strAppPath = strAppPath & "\"
    End If
    
    GetAppPath = strAppPath

End Function

'-----------------------------------------------------------------------------
Public Sub SelectTextboxText(pobjTextbox As TextBox)
'-----------------------------------------------------------------------------
    With pobjTextbox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

'-----------------------------------------------------------------------------
Public Sub TabToNextTextBox(pobjTextBox1 As TextBox, pobjTextBox2 As TextBox)
'-----------------------------------------------------------------------------

    If gblnPopulating Then Exit Sub
    
    If pobjTextBox2.Enabled = False Then Exit Sub
    
    If Len(pobjTextBox1.Text) = pobjTextBox1.MaxLength Then
        pobjTextBox2.SetFocus
    End If

End Sub

'-----------------------------------------------------------------------------
Public Function DigitOnly(pintKeyAscii As Integer) As Integer
'-----------------------------------------------------------------------------

    If (Chr$(pintKeyAscii) >= "0" And Chr$(pintKeyAscii) <= "9") _
    Or (pintKeyAscii < 32) Then
        ' it's OK
        DigitOnly = pintKeyAscii
    Else
        DigitOnly = 0
    End If
    
End Function
