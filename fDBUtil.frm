VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form fDBUtil 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Database Utility"
   ClientHeight    =   5940
   ClientLeft      =   6780
   ClientTop       =   5865
   ClientWidth     =   6240
   Icon            =   "fDBUtil.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   396
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   416
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Data dcNav 
      BackColor       =   &H00C0FFFF&
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'Standard-Cursor
      DefaultType     =   2  'ODBC verwenden
      Exclusive       =   0   'False
      Height          =   345
      Left            =   105
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   750
      Width           =   1740
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'Kein
      Height          =   1155
      Left            =   120
      ScaleHeight     =   77
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2670
      Width           =   2835
      Begin VB.CheckBox ckTop 
         BackColor       =   &H00C0C0C0&
         Caption         =   "On Top"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1890
         Style           =   1  'Grafisch
         TabIndex        =   13
         Top             =   750
         Width           =   810
      End
      Begin VB.ComboBox cbAttri 
         BackColor       =   &H00C0E0FF&
         Height          =   315
         Left            =   15
         Sorted          =   -1  'True
         Style           =   2  'Dropdown-Liste
         TabIndex        =   5
         ToolTipText     =   "Current Field Attributes"
         Top             =   330
         Width           =   2220
      End
      Begin VB.CommandButton btSQL 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SQL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   765
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   8
         ToolTipText     =   "Query the Database"
         Top             =   750
         UseMaskColor    =   -1  'True
         Width           =   555
      End
      Begin VB.CommandButton btDelete 
         Height          =   315
         Left            =   390
         Picture         =   "fDBUtil.frx":0442
         Style           =   1  'Grafisch
         TabIndex        =   7
         ToolTipText     =   "Delete the Current Record"
         Top             =   750
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton btAdd 
         Height          =   315
         Left            =   15
         MaskColor       =   &H00FFFFFF&
         Picture         =   "fDBUtil.frx":0544
         Style           =   1  'Grafisch
         TabIndex        =   6
         ToolTipText     =   "Add a New Record"
         Top             =   750
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "&Field Attributes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   15
         TabIndex        =   4
         Top             =   90
         Width           =   1350
      End
   End
   Begin VB.CheckBox ckTrueFalse 
      BackColor       =   &H00C0E0FF&
      Height          =   270
      Left            =   2565
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1410
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.TextBox txtValue 
      Appearance      =   0  '2D
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2565
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1065
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2925
      Top             =   1665
   End
   Begin VB.CommandButton btOpen 
      Caption         =   "&Open"
      Height          =   555
      Left            =   120
      Picture         =   "fDBUtil.frx":062E
      Style           =   1  'Grafisch
      TabIndex        =   0
      ToolTipText     =   "Open a Database"
      Top             =   75
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid grdData 
      Height          =   1215
      Left            =   105
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1485
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   2143
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   315
      BackColor       =   8438015
      BackColorFixed  =   8421504
      ForeColorFixed  =   12648447
      BackColorBkg    =   12632256
      GridColorFixed  =   4210752
      AllowBigSelection=   0   'False
      Enabled         =   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLines       =   3
      GridLinesFixed  =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdl 
      Left            =   2910
      Top             =   2100
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Access Databases|*.mdb"
   End
   Begin VB.ComboBox cbTables 
      BackColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   690
      Sorted          =   -1  'True
      Style           =   2  'Dropdown-Liste
      TabIndex        =   2
      ToolTipText     =   "Table Names"
      Top             =   330
      Width           =   2175
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Fields    "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   1245
      Width           =   2175
   End
   Begin VB.Label lbl 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Tables:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   720
      TabIndex        =   1
      Top             =   75
      Width           =   645
   End
End
Attribute VB_Name = "fDBUtil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This program is based on a software called DataReader found
'at Planet Source Code some time ago

'Thanks to the original author

DefLng A-Z

Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd, ByVal wMsg, ByVal wParam, lParam As Any)

Private Const HWND_NOTTOPMOST   As Long = -2
Private Const HWND_TOPMOST      As Long = -1
Private Const SWP_NOMOVE        As Long = 2
Private Const SWP_NOSIZE        As Long = 1
Private Const WM_USER           As Long = &H400
Private Const EM_GETLINECOUNT   As Long = WM_USER + &H19
Private Const CB_ISDROPPED      As Long = WM_USER + &H17

Private Const TrueText          As String = "Wahr"   ' "True"  ######### Localize!
Private Const FalseText         As String = "Falsch" ' "False"

Private Margin                  As Long
Private CurrentFieldIndex       As Long
Private i                       As Long
Private FieldType               As String
Private TTT                     As String
Private LastSQL                 As String
Private NewSQL                  As String
Private Bookmark                As String

Private Sub cbAttri_DropDown()

    tmr.Enabled = Len(LastSQL)

End Sub

Private Sub ckTop_Click()

    If ckTop = vbChecked Then
        SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
      Else
        SetWindowPos hWnd, HWND_NOTTOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    End If

End Sub

Private Sub ckTrueFalse_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyCode = 0 'prevent change by Keyboard; update happens with MouseUp

End Sub

Private Sub ckTrueFalse_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    With dcNav.Recordset
        If Not .BOF And Not .EOF Then
            .Edit
            .Fields(grdData.Row - 1).Value = (ckTrueFalse = vbChecked)
            .Update
            grdData.Text = IIf(ckTrueFalse = vbChecked, TrueText, FalseText)
        End If
    End With 'dcNav.RECORDSET
    
End Sub

Private Sub cbTables_Click()

    MousePointer = vbHourglass
    grdData.Visible = False
    With dcNav
        LastSQL = "Select * From [" & cbTables.Text & "]"
        .RecordSource = LastSQL
        .Refresh
        If Not .Recordset.EOF Then
            .Recordset.MoveLast
        End If
        dcNav.Caption = .Recordset.RecordCount & " Records"
        dcNav.ToolTipText = "Navigate in " & cbTables.Text
        If Not .Recordset.BOF Then
            .Recordset.MoveFirst
        End If
    End With 'dcNav
    RefreshGrid
    Form_Resize
    grdData.Visible = True
    grdData_Click
    MousePointer = vbDefault

End Sub

Private Sub btAdd_Click()

    If Not dcNav.Recordset Is Nothing Then
        With dcNav.Recordset
            Bookmark = .Bookmark
            On Error Resume Next
              .AddNew
              For i = 0 To .Fields.Count - 1
                  With .Fields(i)
                      If Not .Attributes And dbAutoIncrField Then
                          Select Case True
                            Case .AllowZeroLength
                              .Value = ""
                            Case .Type = dbBoolean
                              .Value = False
                            Case .Type = dbInteger
                              .Value = 0
                            Case .Type = dbLong
                              .Value = 0
                            Case .Type = dbCurrency
                              .Value = 0
                            Case .Type = dbSingle
                              .Value = 0
                            Case .Type = dbDouble
                              .Value = 0
                            Case .Type = dbDate
                              .Value = Now
                            Case .Type = dbText
                              .Value = " "
                          End Select
                      End If
                  End With '.FIELDS(I)
              Next i
              .Update
              If Err Then
                  MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
                  .Bookmark = Bookmark
                Else
                  .MoveLast
              End If
              RefreshGrid
            On Error GoTo 0
        End With 'dcNav.RECORDSET'I
    End If

End Sub

Private Sub btDelete_Click()

    If Not dcNav.Recordset Is Nothing Then
        With dcNav.Recordset
            If Not .BOF And Not .EOF Then
                .Delete
                .MoveFirst
                RefreshGrid
            End If
        End With 'dcNav.RECORDSET
    End If

End Sub

Private Sub btOpen_Click()
        
    txtValue.Visible = False
    ckTrueFalse.Visible = False
    On Error Resume Next
      cdl.ShowOpen
      i = Err
    On Error GoTo 0
    DoEvents
    If i = 0 Then
        With dcNav
            .DatabaseName = cdl.FileName
            .RecordSource = ""
            .Refresh
            cbTables.Clear
            For i = 0 To .Database.TableDefs.Count - 1
                If InStr(1, .Database.TableDefs(i).Name, "MSys") = 0 Then
                    cbTables.AddItem .Database.TableDefs(i).Name
                End If
            Next i
            cbTables.ListIndex = 0
            grdData_Click
        End With 'dcNav
      Else
        If cbTables.ListCount Then
            cbTables_Click
        End If
    End If
    
End Sub

Private Sub btSQL_Click()

    If Not dcNav.Recordset Is Nothing Then
        On Error Resume Next
          With dcNav
              NewSQL = InputBox("Enter the SQL Statement to use for the Query:", "SQL", LastSQL & " ")
              If NewSQL <> "" Then
                  grdData.Enabled = False
                  grdData.Clear
                  txtValue.Visible = False
                  ckTrueFalse.Visible = False
                  cbAttri.Clear
                  cbAttri.Enabled = False
                  dcNav.Caption = ""
                  LastSQL = NewSQL
                  .RecordSource = LastSQL
                  .Refresh
                  With .Recordset
                      If Not .EOF Then
                          .MoveLast
                          dcNav.Caption = LastSQL & ": " & .RecordCount & " Records"
                          dcNav.ToolTipText = "Navigate in Resultset"
                          .MoveFirst
                      End If
                  End With '.RECORDSET
                  RefreshGrid
                  grdData_Click
              End If
          End With 'dcNav
        On Error GoTo 0
    End If

End Sub

Private Sub dcNav_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    txtValue_LostFocus

End Sub

Private Sub dcNav_Reposition()
    
    RefreshGrid
    
End Sub

Private Sub Form_Load()

    Margin = btOpen.Left
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
      i = ScaleWidth - Margin - Margin
      If WindowState <> vbMinimized Then
          cbTables.Width = i - cbTables.Left + Margin
          dcNav.Width = i + 2
          pic.Width = i + 1
          ckTop.Left = pic.ScaleWidth - ckTop.Width
          lbl(2).Width = i
          pic.Top = ScaleHeight - pic.Height - 4
          cbAttri.Width = pic.ScaleWidth
          With grdData
              .Width = i + 2
              .Height = pic.Top - .Top + 1
              If .Rows * (.CellHeight + 15) / 15 > .Height - 5 Then
                  .ColWidth(0) = .Width * 7.5 - 180
                  lbl(2) = "Fields    "
                Else
                  .ColWidth(0) = .Width * 7.5 - 60
                  lbl(2) = "Fields"
              End If
              .ColWidth(1) = .ColWidth(0)
          End With 'GRDDATA
          MoveOverlays
          AdjustBackcolor
      End If
    On Error GoTo 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    txtValue_LostFocus
     
End Sub

Private Sub grdData_Click()
        
    Select Case dcNav.Recordset.Fields(grdData.Row - 1).Type
      Case dbBoolean
        FieldType = "Boolean"
      Case dbInteger
        FieldType = "Integer"
      Case dbLong
        FieldType = "Long"
      Case dbSingle
        FieldType = "Single"
      Case dbDouble
        FieldType = "Double"
      Case dbCurrency
        FieldType = "Currency"
      Case dbDate
        FieldType = "Date/Time"
      Case dbText
        FieldType = "Text"
      Case 12
        FieldType = "Memo"
      Case Else
        FieldType = "Other"
    End Select
    With cbAttri
        .Enabled = True
        .Clear
        .AddItem " " & dcNav.Recordset.Fields(grdData.Row - 1).Name & " (" & FieldType & ")"
        For i = 1 To dcNav.Recordset.Fields(grdData.Row - 1).Properties.Count - 1
            On Error Resume Next
              Select Case dcNav.Recordset.Fields(grdData.Row - 1).Properties(i).Name
                Case "Attributes"
                  .AddItem " " & dcNav.Recordset.Fields(grdData.Row - 1).Properties(i).Name & ": x'" & Hex$(dcNav.Recordset.Fields(grdData.Row - 1).Properties(i).Value) & "'"
                Case "Size"
                  .AddItem " " & dcNav.Recordset.Fields(grdData.Row - 1).Properties(i).Name & ": " & dcNav.Recordset.Fields(grdData.Row - 1).Properties(i).Value
                  Select Case FieldType
                    Case "Text"
                      txtValue.MaxLength = dcNav.Recordset.Fields(grdData.Row - 1).Properties(i).Value
                    Case "Date/Time"
                      txtValue.MaxLength = 19
                    Case Else
                      txtValue.MaxLength = 0
                  End Select
                Case Else
                  .AddItem " " & dcNav.Recordset.Fields(grdData.Row - 1).Properties(i).Name & ": " & dcNav.Recordset.Fields(grdData.Row - 1).Properties(i).Value
              End Select
            On Error GoTo 0
        Next i
        .ListIndex = 0
        .Refresh
        DoEvents
    End With 'cbAttri
    MoveOverlays
    RefreshOverlays
    On Error Resume Next
      txtValue.SetFocus
    On Error GoTo 0
    
End Sub

Private Sub grdData_Scroll()

    MoveOverlays

End Sub

Private Sub tmr_Timer()

    If SendMessage(cbAttri.hWnd, CB_ISDROPPED, 0, 0) = False Then
        tmr.Enabled = False
        On Error Resume Next
          cbAttri.ListIndex = 0
        On Error GoTo 0
        RefreshOverlays
        grdData_Click
    End If

End Sub

Private Sub txtValue_GotFocus()

    CurrentFieldIndex = grdData.Row - 1

End Sub

Private Sub txtValue_LostFocus()

    If txtValue.DataChanged Then
        With dcNav.Recordset
            .Edit
            If LCase$(txtValue.Text) = "{null}" Or LCase$(txtValue.Text) = "null" Then
                .Fields(CurrentFieldIndex).Value = Null
              Else
                On Error Resume Next
                  .Fields(CurrentFieldIndex).Value = CVar(txtValue.Text)
                On Error GoTo 0
            End If
            .Update
        End With 'dcNav.RECORDSET
        txtValue.DataChanged = False
    End If
    
End Sub

Private Sub RefreshGrid()
  
    With grdData
        .Enabled = (dcNav.Recordset.RecordCount > 0)
        .Clear
        .Rows = dcNav.Recordset.Fields.Count + 1
        .TextMatrix(0, 0) = "Field Name"
        .TextMatrix(0, 1) = "Value"
        .ColAlignment(1) = flexAlignLeftCenter
        For i = 0 To dcNav.Recordset.Fields.Count - 1
            .TextMatrix(i + 1, 0) = dcNav.Recordset.Fields(i).Name
            If Not dcNav.Recordset.EOF And dcNav.Recordset.RecordCount > 0 Then
                With dcNav.Recordset.Fields(i)
                    Select Case .Type
                      Case dbBoolean, _
                           dbInteger, _
                           dbLong, _
                           dbSingle, _
                           dbDouble, _
                           dbCurrency, _
                           dbDate, _
                           dbText
                        If IsNull(.Value) Then
                            grdData.TextMatrix(i + 1, 1) = "{Null}"
                          Else
                            grdData.TextMatrix(i + 1, 1) = .Value
                        End If
                    End Select
                End With 'dcNav.RECORDSET.FIELDS(I)
            End If
        Next i
    End With 'GRDDATA
    MoveOverlays
    RefreshOverlays
          
End Sub

Private Sub MoveOverlays()

    If dcNav.RecordSource <> "" Then
        With grdData
            .Col = 1
            If dcNav.Recordset.Fields(.Row - 1).Type = dbBoolean Then
                'strange things happen: querying .CellTop scrolls
                'when it is outside the boundaries of the grid
                ckTrueFalse.Move .CellLeft / 15 + .Left + 2, _
                                 .CellTop / 15 + .Top + 2, _
                                 .CellWidth / 15 - 3, _
                                 .CellHeight / 15 - 3
                ckTrueFalse.Visible = True
                txtValue.Visible = False
              Else
                ckTrueFalse.Visible = False
                txtValue.Move .CellLeft / 15 + .Left + 2, _
                              .CellTop / 15 + .Top + 2, _
                              .CellWidth / 15 - 3, _
                              .CellHeight / 15 - 3
                txtValue.Visible = True
            End If
        End With 'GRDDATA
    End If

End Sub

Private Sub RefreshOverlays()

    TTT = "Alter Value of Field: " & dcNav.Recordset.Fields(grdData.Row - 1).Name
    If txtValue.Visible Then
        With txtValue
            .Text = grdData.Text
            .DataChanged = False
            .SelStart = 0
            .ToolTipText = TTT
        End With 'TXTVALUE
        AdjustBackcolor
    End If
    If ckTrueFalse.Visible Then
        With ckTrueFalse
            .Value = IIf(grdData.Text = TrueText, vbChecked, vbUnchecked)
            .SetFocus
            .ToolTipText = TTT
        End With 'CKTRUEFALSE
    End If

End Sub

Private Sub AdjustBackcolor()

    With txtValue
        If SendMessage(.hWnd, EM_GETLINECOUNT, Len(.Text), 0) Then
            .BackColor = &HC0F0E0
          Else
            .BackColor = &HC0E0FF
        End If
    End With 'TXTVALUE

End Sub

':) Ulli's Code Formatter V2.3 (02.09.2001 16:11:01) 31 + 445 = 476 Lines
