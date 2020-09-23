Attribute VB_Name = "modValidate"
Option Explicit

'-----------------------------------------------------------------------------
Public Function ValidateRequiredField(pobjCtl As Control, _
                                      pstrFieldDesc As String) _
As Boolean
'-----------------------------------------------------------------------------

    If Trim$(pobjCtl.Text) = "" Then
        MsgBox pstrFieldDesc & " must not be blank.", _
               vbExclamation, _
               pstrFieldDesc
        pobjCtl.SetFocus
        ValidateRequiredField = False
    Else
        ValidateRequiredField = True
    End If

End Function

'-----------------------------------------------------------------------------
Public Function ValidateZipCode(pobjZip5 As Control, _
                                Optional pobjZip4 As Control = Nothing) _
As Boolean
'-----------------------------------------------------------------------------
    
    Dim strErrorMsg     As String
    Dim objFocusControl As Control
    
    If Not pobjZip4 Is Nothing Then
        If Trim$(pobjZip5.Text) = "" _
        And Trim$(pobjZip4.Text) <> "" _
        Then
            strErrorMsg = "First part of Zip must be valued when '+4' part " _
                        & "is valued."
            Set objFocusControl = pobjZip5
            GoTo ValidateZipCode_Error
        End If
    End If
    
    If Len(pobjZip5.Text) = 0 _
    Or Len(pobjZip5.Text) = 5 _
    Then
        ' it's OK
    Else
        strErrorMsg = "Invalid length for Zip."
        Set objFocusControl = pobjZip5
        GoTo ValidateZipCode_Error
    End If
    
    If Not pobjZip4 Is Nothing Then
        If Len(pobjZip4.Text) = 0 _
        Or Len(pobjZip4.Text) = 4 _
        Then
            ' it's OK
        Else
            strErrorMsg = "Invalid length for Zip '+4' part."
            Set objFocusControl = pobjZip4
            GoTo ValidateZipCode_Error
        End If
    End If
    
    ValidateZipCode = True
    Exit Function
    
ValidateZipCode_Error:
    MsgBox strErrorMsg, _
           vbExclamation, _
           "Zip"
    objFocusControl.SetFocus
    ValidateZipCode = False

End Function
   
'-----------------------------------------------------------------------------
Public Function ValidatePhoneNumber(pobjArea As Control, _
                                    pobjPrfx As Control, _
                                    pobjLine As Control, _
                                    Optional pobjExt As Control = Nothing, _
                                    Optional pblnBlankOK As Boolean = True) _
As Boolean
'-----------------------------------------------------------------------------

    Dim strErrorMsg             As String
    Dim blnIncompletePhoneNbr   As Boolean
    Dim objFocusControl         As Control

    If pobjArea.Text = "" _
    And pobjPrfx.Text = "" _
    And pobjLine.Text = "" Then
        If pblnBlankOK Then
            If pobjExt Is Nothing Then
                ValidatePhoneNumber = True
                Exit Function
            Else
                If pobjExt.Text = "" Then
                    ValidatePhoneNumber = True
                    Exit Function
                Else
                    strErrorMsg = "Phone Number must be valued when extension is valued."
                    Set objFocusControl = pobjArea
                    GoTo ValidatePhoneNumber_Error
                End If
            End If
        Else
            strErrorMsg = "Phone Number must not be blank."
            Set objFocusControl = pobjArea
            GoTo ValidatePhoneNumber_Error
        End If
    End If

    blnIncompletePhoneNbr = False
    If Len(pobjArea.Text) <> 3 Then
        Set objFocusControl = pobjArea
        blnIncompletePhoneNbr = True
    ElseIf Len(pobjPrfx.Text) <> 3 Then
        Set objFocusControl = pobjPrfx
        blnIncompletePhoneNbr = True
    ElseIf Len(pobjLine.Text) <> 4 Then
        Set objFocusControl = pobjLine
        blnIncompletePhoneNbr = True
    End If
    If blnIncompletePhoneNbr Then
        strErrorMsg = "Phone Number entry is incomplete."
        GoTo ValidatePhoneNumber_Error
    End If

    ValidatePhoneNumber = True
    Exit Function
    
ValidatePhoneNumber_Error:
    MsgBox strErrorMsg, _
           vbExclamation, _
           "Phone Number"
    objFocusControl.SetFocus
    ValidatePhoneNumber = False

End Function

'-----------------------------------------------------------------------------
Public Function ValidateDate(pobjMonth As Control, _
                             pobjDay As Control, _
                             pobjYear As Control, _
                             Optional pblnBlankOK As Boolean = True) _
As Boolean
'-----------------------------------------------------------------------------

    Dim strErrorMsg     As String
    Dim objFocusControl As Control

    If pobjMonth.Text = "" _
    And pobjDay.Text = "" _
    And pobjYear.Text = "" Then
        If pblnBlankOK Then
            ValidateDate = True
            Exit Function
        Else
            strErrorMsg = "Date must not be blank."
            Set objFocusControl = pobjMonth
            GoTo ValidateDate_Error
        End If
    End If

    If pobjYear.Text <> "" And Len(pobjYear.Text) <> 4 Then
        strErrorMsg = "Four digits must be entered for the year."
        Set objFocusControl = pobjYear
        GoTo ValidateDate_Error
    End If

    If Not IsDate(pobjMonth.Text & "/" _
                & pobjDay.Text & "/" _
                & pobjYear.Text) Then
        strErrorMsg = "Date must be a valid date" & IIf(pblnBlankOK, " or blank", "") & "."
        Set objFocusControl = pobjMonth
        GoTo ValidateDate_Error
    End If

    ValidateDate = True
    Exit Function

ValidateDate_Error:
    MsgBox strErrorMsg, _
           vbExclamation, _
           "Date Error"
    objFocusControl.SetFocus
    ValidateDate = False

End Function

'-----------------------------------------------------------------------------
Public Function ValidateSSN(pobjSSN1 As Control, _
                            pobjSSN2 As Control, _
                            pobjSSN3 As Control) _
As Boolean
'-----------------------------------------------------------------------------

    Dim strErrorMsg         As String
    Dim strErrorField       As String
    Dim blnIncompleteSSN    As Boolean
    Dim objFocusControl     As Control
    
    If (Trim$(pobjSSN1.Text) = "" And _
        Trim$(pobjSSN2.Text) = "" And _
        Trim$(pobjSSN3.Text) = "") _
    Or (Trim$(pobjSSN1.Text) <> "" And _
        Trim$(pobjSSN2.Text) <> "" And _
        Trim$(pobjSSN3.Text) <> "") _
    Then
        If Trim$(pobjSSN1.Text) <> "" Then
            blnIncompleteSSN = False
            If Len(pobjSSN1.Text) <> 3 Then
                Set objFocusControl = pobjSSN1
                blnIncompleteSSN = True
            ElseIf Len(pobjSSN2.Text) <> 2 Then
                Set objFocusControl = pobjSSN2
                blnIncompleteSSN = True
            ElseIf Len(pobjSSN3.Text) <> 4 Then
                Set objFocusControl = pobjSSN3
                blnIncompleteSSN = True
            End If
            If blnIncompleteSSN Then
                strErrorMsg = "SSN entry is incomplete."
                strErrorField = "SSN"
                GoTo ValidateSSN_Error
            End If
        End If
    Else
        strErrorMsg = "A partial SSN is not valid. " _
                    & "Either fill all parts, or leave all parts blank."
        strErrorField = "SSN"
        Set objFocusControl = pobjSSN1
        GoTo ValidateSSN_Error
    End If
    
    ValidateSSN = True
    Exit Function

ValidateSSN_Error:
    MsgBox strErrorMsg, _
           vbExclamation, _
           "SSN Error"
    objFocusControl.SetFocus
    ValidateSSN = False

End Function
