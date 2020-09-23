VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCustMaint 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Table Maintenance"
   ClientHeight    =   7200
   ClientLeft      =   3660
   ClientTop       =   1695
   ClientWidth     =   10485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustMaint.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10485
   Begin VB.Frame fraCurrentRec 
      Caption         =   "Current Record"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2475
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   8895
      Begin VB.TextBox txtLast 
         Height          =   375
         Left            =   4560
         TabIndex        =   10
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox txtFirst 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   4215
      End
      Begin VB.TextBox txtAddr 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   1200
         Width           =   8655
      End
      Begin VB.TextBox txtZip 
         Height          =   375
         Left            =   6120
         MaxLength       =   5
         TabIndex        =   19
         Text            =   "99999"
         Top             =   1920
         Width           =   675
      End
      Begin VB.TextBox txtCity 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   5115
      End
      Begin VB.TextBox txtState 
         Height          =   375
         Left            =   5340
         MaxLength       =   2
         TabIndex        =   18
         Top             =   1920
         Width           =   435
      End
      Begin VB.TextBox txtPrfx 
         Height          =   375
         Left            =   7620
         MaxLength       =   3
         TabIndex        =   24
         Text            =   "999"
         Top             =   1920
         Width           =   435
      End
      Begin VB.TextBox txtArea 
         Height          =   375
         Left            =   7020
         MaxLength       =   3
         TabIndex        =   22
         Text            =   "999"
         Top             =   1920
         Width           =   435
      End
      Begin VB.TextBox txtLine 
         Height          =   375
         Left            =   8220
         MaxLength       =   4
         TabIndex        =   25
         Text            =   "9999"
         Top             =   1920
         Width           =   555
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         Height          =   195
         Index           =   0
         Left            =   4620
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "("
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   6900
         TabIndex        =   21
         Top             =   1980
         Width           =   90
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   900
         Width           =   675
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   13
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   5340
         TabIndex        =   14
         Top             =   1620
         Width           =   495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Zip"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   6120
         TabIndex        =   15
         Top             =   1620
         Width           =   435
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone Number"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   6900
         TabIndex        =   16
         Top             =   1620
         Width           =   1695
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ")"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   7500
         TabIndex        =   23
         Top             =   1980
         Width           =   90
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   8100
         TabIndex        =   20
         Top             =   1920
         Width           =   75
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      DisabledPicture =   "frmCustMaint.frx":08CA
      Height          =   855
      Left            =   9180
      Picture         =   "frmCustMaint.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3540
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      DisabledPicture =   "frmCustMaint.frx":1A5E
      Height          =   855
      Left            =   9180
      Picture         =   "frmCustMaint.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   4860
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      DisabledPicture =   "frmCustMaint.frx":2BF2
      Height          =   855
      Left            =   9180
      Picture         =   "frmCustMaint.frx":34BC
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   5940
      Width           =   1155
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      DisabledPicture =   "frmCustMaint.frx":3D86
      Height          =   855
      Left            =   9180
      Picture         =   "frmCustMaint.frx":4650
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1620
      Width           =   1155
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      DisabledPicture =   "frmCustMaint.frx":4F1A
      Height          =   855
      Left            =   9180
      Picture         =   "frmCustMaint.frx":57E4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   660
      Width           =   1155
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      DisabledPicture =   "frmCustMaint.frx":60AE
      Height          =   855
      Left            =   9180
      Picture         =   "frmCustMaint.frx":6978
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2580
      Width           =   1155
   End
   Begin MSComctlLib.ImageList imlLVIcons 
      Left            =   9780
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustMaint.frx":7242
            Key             =   "Custs"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwCustomer 
      Height          =   3855
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   6800
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "imlLVIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label lblPhysMaint 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Table Maintenance"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   420
      Left            =   900
      TabIndex        =   0
      Top             =   60
      Width           =   9345
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmCustMaint.frx":75DC
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frmCustMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjConn                As ADODB.Connection
Private mobjCmd                 As ADODB.Command
Private mobjRst                 As ADODB.Recordset

Private mstrMaintMode           As String
Private mblnFormActivated       As Boolean
Private mblnUpdateInProgress    As Boolean

' Customer LV SubItem Indexes ...
Private Const mlngCUST_LAST_IDX          As Long = 1
Private Const mlngCUST_ADDR_IDX          As Long = 2
Private Const mlngCUST_CITY_IDX          As Long = 3
Private Const mlngCUST_ST_IDX            As Long = 4
Private Const mlngCUST_ZIP_IDX           As Long = 5
Private Const mlngCUST_PHONE_IDX         As Long = 6
Private Const mlngCUST_ID_IDX            As Long = 7

'*****************************************************************************
'*                          General Form Events                              *
'*****************************************************************************

'-----------------------------------------------------------------------------
Private Sub Form_Load()
'-----------------------------------------------------------------------------
    
    CenterForm Me
    
    ConnectToDB
    
    SetupCustLVCols
    
    LoadCustomerListView
    
End Sub

'-----------------------------------------------------------------------------
Private Sub Form_Activate()
'-----------------------------------------------------------------------------
    
    If mblnFormActivated Then Exit Sub
    
    Refresh
    
    SetFormState True
    
    mblnFormActivated = True

End Sub

'-----------------------------------------------------------------------------
Private Sub Form_Unload(Cancel As Integer)
'-----------------------------------------------------------------------------
    
    Dim objRst  As ADODB.Recordset
    
    If mblnUpdateInProgress Then
        MsgBox "You must save or cancel the current action before " _
             & "closing this window.", _
               vbInformation, _
               "Cannot Close"
        Cancel = 1
        Exit Sub
    End If
        
    DisconnectFromDB
    
    Set frmCustMaint = Nothing

End Sub

'*****************************************************************************
'*                        Command Button Events                              *
'*****************************************************************************

'-----------------------------------------------------------------------------
Private Sub cmdAdd_Click()
'-----------------------------------------------------------------------------

    mstrMaintMode = "ADD"
    mblnUpdateInProgress = True
    
    ClearCurrRecControls
    
    SetFormState False
    
    txtFirst.SetFocus

End Sub


'-----------------------------------------------------------------------------
Private Sub cmdUpdate_Click()
'-----------------------------------------------------------------------------
    
    If lvwCustomer.SelectedItem Is Nothing Then
        MsgBox "No Customer selected to update.", _
               vbExclamation, _
               "Update"
        Exit Sub
    End If
    
    mstrMaintMode = "EDIT"
    mblnUpdateInProgress = True
    
    SetFormState False
    
    txtFirst.SetFocus
    
End Sub


'-----------------------------------------------------------------------------
Private Sub cmdDelete_Click()
'-----------------------------------------------------------------------------

    Dim strFirstName    As String
    Dim strLastName     As String
    Dim lngCustID       As Long
    Dim lngNewSelIndex  As Long
    
    If lvwCustomer.SelectedItem Is Nothing Then
        MsgBox "No Customer selected to delete.", _
               vbExclamation, _
               "Delete"
        Exit Sub
    End If
    
    With lvwCustomer.SelectedItem
        strFirstName = .Text
        strLastName = .SubItems(mlngCUST_LAST_IDX)
        lngCustID = CLng(.SubItems(mlngCUST_ID_IDX))
    End With
    
    If MsgBox("Are you sure that you want to delete Customer '" _
            & strFirstName & " " & strLastName & "'?", _
              vbYesNo + vbQuestion, _
              "Confirm Delete") = vbNo Then
        Exit Sub
    End If
    
    mobjCmd.CommandText = "DELETE FROM Customer WHERE CustID = " & lngCustID
    mobjCmd.Execute
    
    With lvwCustomer
        If .SelectedItem.Index = .ListItems.Count Then
            lngNewSelIndex = .ListItems.Count - 1
        Else
            lngNewSelIndex = .SelectedItem.Index
        End If
        .ListItems.Remove .SelectedItem.Index
        If .ListItems.Count > 0 Then
            Set .SelectedItem = .ListItems(lngNewSelIndex)
            lvwCustomer_ItemClick .SelectedItem
        Else
            ClearCurrRecControls
        End If
    End With

End Sub

'-----------------------------------------------------------------------------
Private Sub cmdClose_Click()
'-----------------------------------------------------------------------------
    
    Unload Me

End Sub


'-----------------------------------------------------------------------------
Private Sub cmdSave_Click()
'-----------------------------------------------------------------------------

    Dim strPhone        As String
    Dim objNewListItem  As ListItem
    Dim lngIDField      As Long
    Dim strSQL          As String

    If Not ValidateFormFields Then Exit Sub
    
    strPhone = txtArea.Text & txtPrfx.Text & txtLine.Text
    
    If mstrMaintMode = "ADD" Then
    
        lngIDField = GetNextCustID()
        
        strSQL = "INSERT INTO Customer(  CustID"
        strSQL = strSQL & "            , FirstName"
        strSQL = strSQL & "            , LastName"
        strSQL = strSQL & "            , Address"
        strSQL = strSQL & "            , City"
        strSQL = strSQL & "            , State"
        strSQL = strSQL & "            , Zip"
        strSQL = strSQL & "            , PhoneNumber"
        strSQL = strSQL & "         ) VALUES ("
        strSQL = strSQL & lngIDField
        strSQL = strSQL & ", '" & Replace$(txtFirst.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(txtLast.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(txtAddr.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & Replace$(txtCity.Text, "'", "''") & "'"
        strSQL = strSQL & ", '" & txtState.Text & "'"
        strSQL = strSQL & ", '" & txtZip.Text & "'"
        strSQL = strSQL & ", '" & strPhone & "'"
        strSQL = strSQL & ")"
        
        Set objNewListItem = lvwCustomer.ListItems.Add(, , txtFirst.Text, , "Custs")
        PopulateListItem objNewListItem
        With objNewListItem
            .SubItems(mlngCUST_ID_IDX) = CStr(lngIDField)
            .EnsureVisible
        End With
        Set lvwCustomer.SelectedItem = objNewListItem
        Set objNewListItem = Nothing
    Else
        lngIDField = CLng(lvwCustomer.SelectedItem.SubItems(mlngCUST_ID_IDX))
        
        strSQL = "UPDATE Customer SET "
        strSQL = strSQL & "  FirstName   = '" & Replace$(txtFirst.Text, "'", "''") & "'"
        strSQL = strSQL & ", LastName    = '" & Replace$(txtLast.Text, "'", "''") & "'"
        strSQL = strSQL & ", Address     = '" & Replace$(txtAddr.Text, "'", "''") & "'"
        strSQL = strSQL & ", City        = '" & Replace$(txtCity.Text, "'", "''") & "'"
        strSQL = strSQL & ", State       = '" & txtState.Text & "'"
        strSQL = strSQL & ", Zip         = '" & txtZip.Text & "'"
        strSQL = strSQL & ", PhoneNumber = '" & strPhone & "'"
        strSQL = strSQL & " WHERE CustID = " & lngIDField
        
        lvwCustomer.SelectedItem.Text = txtFirst.Text
        PopulateListItem lvwCustomer.SelectedItem
    End If
    
    mobjCmd.CommandText = strSQL
    mobjCmd.Execute
    
    SetFormState True

    mblnUpdateInProgress = False

End Sub


'-----------------------------------------------------------------------------
Private Sub cmdCancel_Click()
'-----------------------------------------------------------------------------
    
    mblnUpdateInProgress = False
    SetFormState True
    lvwCustomer_ItemClick lvwCustomer.SelectedItem
    
End Sub


'*****************************************************************************
'*                          ListView Events                                  *
'*****************************************************************************

'-------------------------------------------------------------------------
Private Sub lvwCustomer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'-------------------------------------------------------------------------
    
    ' sort the listview on the column clicked
    With lvwCustomer
        If (.Sorted) And (ColumnHeader.SubItemIndex = .SortKey) Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .Sorted = True
            .SortKey = ColumnHeader.SubItemIndex
            .SortOrder = lvwAscending
        End If
        .Refresh
    End With
        
    ' If an item was selected prior to the sort,
    ' make sure it is still visible now that the sort is done.
    If Not lvwCustomer.SelectedItem Is Nothing Then
        lvwCustomer.SelectedItem.EnsureVisible
    End If

End Sub

'-----------------------------------------------------------------------------
Private Sub lvwCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
'-----------------------------------------------------------------------------

    gblnPopulating = True
    
    With Item
        txtFirst.Text = .Text
        txtLast.Text = .SubItems(mlngCUST_LAST_IDX)
        txtAddr.Text = .SubItems(mlngCUST_ADDR_IDX)
        txtCity.Text = .SubItems(mlngCUST_CITY_IDX)
        txtState.Text = .SubItems(mlngCUST_ST_IDX)
        txtZip.Text = .SubItems(mlngCUST_ZIP_IDX)
        If .SubItems(mlngCUST_PHONE_IDX) = "" Then
            txtArea.Text = ""
            txtPrfx.Text = ""
            txtLine.Text = ""
        Else
            txtArea.Text = Mid$(.SubItems(mlngCUST_PHONE_IDX), 2, 3)
            txtPrfx.Text = Mid$(.SubItems(mlngCUST_PHONE_IDX), 7, 3)
            txtLine.Text = Right$(.SubItems(mlngCUST_PHONE_IDX), 4)
        End If
    End With
    
    gblnPopulating = False
    
End Sub

'*****************************************************************************
'*                      Other Control Events                                 *
'*****************************************************************************

Private Sub txtFirst_GotFocus()
    SelectTextboxText txtFirst
End Sub
Private Sub txtLast_GotFocus()
    SelectTextboxText txtLast
End Sub
Private Sub txtAddr_GotFocus()
    SelectTextboxText txtAddr
End Sub
Private Sub txtCity_GotFocus()
    SelectTextboxText txtCity
End Sub
Private Sub txtState_GotFocus()
    SelectTextboxText txtState
End Sub
Private Sub txtState_KeyPress(KeyAscii As Integer)
    If KeyAscii < 32 Then Exit Sub
    If Chr$(KeyAscii) >= "a" And Chr$(KeyAscii) <= "z" Then
        KeyAscii = KeyAscii - 32
    ElseIf Chr$(KeyAscii) >= "A" And Chr$(KeyAscii) <= "Z" Then
        ' OK
    Else
        KeyAscii = 0
    End If
End Sub
Private Sub txtState_Change()
    TabToNextTextBox txtState, txtZip
End Sub

Private Sub txtZip_GotFocus()
    SelectTextboxText txtZip
End Sub
Private Sub txtZip_KeyPress(KeyAscii As Integer)
    KeyAscii = DigitOnly(KeyAscii)
End Sub
Private Sub txtZip_Change()
    TabToNextTextBox txtZip, txtArea
End Sub

Private Sub txtArea_GotFocus()
    SelectTextboxText txtArea
End Sub
Private Sub txtArea_KeyPress(KeyAscii As Integer)
    KeyAscii = DigitOnly(KeyAscii)
End Sub
Private Sub txtArea_Change()
    TabToNextTextBox txtArea, txtPrfx
End Sub

Private Sub txtPrfx_GotFocus()
    SelectTextboxText txtPrfx
End Sub
Private Sub txtPrfx_KeyPress(KeyAscii As Integer)
    KeyAscii = DigitOnly(KeyAscii)
End Sub
Private Sub txtPrfx_Change()
    TabToNextTextBox txtPrfx, txtLine
End Sub

Private Sub txtLine_GotFocus()
    SelectTextboxText txtLine
End Sub
Private Sub txtLine_KeyPress(KeyAscii As Integer)
    KeyAscii = DigitOnly(KeyAscii)
End Sub


'*****************************************************************************
'*               Programmer-Defined Subs & Functions                         *
'*****************************************************************************

'-----------------------------------------------------------------------------
Private Sub ConnectToDB()
'-----------------------------------------------------------------------------

    Set mobjConn = New ADODB.Connection
    mobjConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" _
                              & "Data Source=" _
                              & GetAppPath _
                              & "Cust.mdb"
    mobjConn.Open

    Set mobjCmd = New ADODB.Command
    Set mobjCmd.ActiveConnection = mobjConn
    mobjCmd.CommandType = adCmdText

End Sub

'-----------------------------------------------------------------------------
Private Sub DisconnectFromDB()
'-----------------------------------------------------------------------------

    Set mobjCmd = Nothing
    
    mobjConn.Close
    Set mobjConn = Nothing

End Sub


'-----------------------------------------------------------------------------
Private Sub ClearCurrRecControls()
'-----------------------------------------------------------------------------
    
    gblnPopulating = True
    
    txtFirst.Text = ""
    txtLast.Text = ""
    txtAddr.Text = ""
    txtCity.Text = ""
    txtState.Text = ""
    txtZip.Text = ""
    txtArea.Text = ""
    txtPrfx.Text = ""
    txtLine.Text = ""

    gblnPopulating = False
    
End Sub

'-----------------------------------------------------------------------------
Private Sub SetFormState(pblnEnabled As Boolean)
'-----------------------------------------------------------------------------

    lvwCustomer.Enabled = pblnEnabled
    cmdAdd.Enabled = pblnEnabled
    cmdUpdate.Enabled = pblnEnabled
    cmdDelete.Enabled = pblnEnabled
    cmdClose.Enabled = pblnEnabled
    
    txtFirst.Enabled = Not pblnEnabled
    txtLast.Enabled = Not pblnEnabled
    txtAddr.Enabled = Not pblnEnabled
    txtCity.Enabled = Not pblnEnabled
    txtState.Enabled = Not pblnEnabled
    txtZip.Enabled = Not pblnEnabled
    txtArea.Enabled = Not pblnEnabled
    txtPrfx.Enabled = Not pblnEnabled
    txtLine.Enabled = Not pblnEnabled
    
    cmdSave.Enabled = Not pblnEnabled
    cmdCancel.Enabled = Not pblnEnabled

End Sub

'-----------------------------------------------------------------------------
Private Function ValidateFormFields() As Boolean
'-----------------------------------------------------------------------------
    
    If Not ValidateRequiredField(txtFirst, "First Name") Then
        ValidateFormFields = False
        Exit Function
    End If
    
    If Not ValidateRequiredField(txtLast, "Last Name") Then
        ValidateFormFields = False
        Exit Function
    End If
   
    If Not ValidateZipCode(txtZip) Then
        ValidateFormFields = False
        Exit Function
    End If
    
    If Not ValidatePhoneNumber(txtArea, txtPrfx, txtLine) Then
        ValidateFormFields = False
        Exit Function
    End If
        
    ValidateFormFields = True
    
End Function

'-----------------------------------------------------------------------------
Private Sub PopulateListItem(pobjListItem As ListItem)
'-----------------------------------------------------------------------------

    With pobjListItem
        .SubItems(mlngCUST_LAST_IDX) = txtLast.Text
        .SubItems(mlngCUST_ADDR_IDX) = txtAddr.Text
        .SubItems(mlngCUST_CITY_IDX) = txtCity.Text
        .SubItems(mlngCUST_ST_IDX) = txtState.Text
        .SubItems(mlngCUST_ZIP_IDX) = txtZip.Text
        .SubItems(mlngCUST_PHONE_IDX) _
            = IIf(txtArea.Text = "", _
                  "", _
                  "(" & txtArea.Text & ") " & txtPrfx.Text & "-" & txtLine.Text)

    End With

End Sub

'-----------------------------------------------------------------------------
Private Sub SetupCustLVCols()
'-----------------------------------------------------------------------------
                                 
    With lvwCustomer
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "First Name", .Width * 0.15
        .ColumnHeaders.Add , , "Last Name", .Width * 0.12
        .ColumnHeaders.Add , , "Address", .Width * 0.2
        .ColumnHeaders.Add , , "City", .Width * 0.15
        .ColumnHeaders.Add , , "St", .Width * 0.06
        .ColumnHeaders.Add , , "Zip", .Width * 0.1
        .ColumnHeaders.Add , , "Phone #", .Width * 0.2
        .ColumnHeaders.Add , , "ID", 0
    End With

End Sub

'-----------------------------------------------------------------------------
Private Sub LoadCustomerListView()
'-----------------------------------------------------------------------------
                                 
    Dim strSQL      As String
    Dim objCurrLI   As ListItem
    Dim strZip      As String
    Dim strPhone    As String
                                 
    strSQL = "SELECT FirstName" _
           & "     , LastName" _
           & "     , Address" _
           & "     , City" _
           & "     , State" _
           & "     , Zip" _
           & "     , PhoneNumber" _
           & "     , CustID" _
           & "  FROM Customer " _
           & " ORDER BY LastName" _
           & "        , FirstName"
    
    mobjCmd.CommandText = strSQL
    Set mobjRst = mobjCmd.Execute
    
    lvwCustomer.ListItems.Clear
    
    With mobjRst
        Do Until .EOF
            strPhone = !PhoneNumber & ""
            If Len(strPhone) > 0 Then
                strPhone = "(" & Left$(strPhone, 3) & ") " _
                         & Mid$(strPhone, 4, 3) & "-" _
                         & Right$(strPhone, 4)
            End If
            Set objCurrLI = lvwCustomer.ListItems.Add(, , !FirstName & "", , "Custs")
            objCurrLI.SubItems(mlngCUST_LAST_IDX) = !LastName & ""
            objCurrLI.SubItems(mlngCUST_ADDR_IDX) = !Address & ""
            objCurrLI.SubItems(mlngCUST_CITY_IDX) = !City & ""
            objCurrLI.SubItems(mlngCUST_ST_IDX) = !State & ""
            objCurrLI.SubItems(mlngCUST_ZIP_IDX) = !Zip & ""
            objCurrLI.SubItems(mlngCUST_PHONE_IDX) = strPhone
            objCurrLI.SubItems(mlngCUST_ID_IDX) = CStr(!CustID)
            .MoveNext
        Loop
    End With
    
    With lvwCustomer
        If .ListItems.Count > 0 Then
            Set .SelectedItem = .ListItems(1)
            lvwCustomer_ItemClick .SelectedItem
        End If
    End With
    
    Set objCurrLI = Nothing
    Set mobjRst = Nothing

End Sub

'------------------------------------------------------------------------
Private Function GetNextCustID() As Long
'------------------------------------------------------------------------

    mobjCmd.CommandText = "SELECT MAX(CustID) AS MaxID FROM Customer"
    Set mobjRst = mobjCmd.Execute

    If mobjRst.EOF Then
        GetNextCustID = 1
    ElseIf IsNull(mobjRst!MaxID) Then
        GetNextCustID = 1
    Else
        GetNextCustID = mobjRst!MaxID + 1
    End If

    Set mobjRst = Nothing

End Function


