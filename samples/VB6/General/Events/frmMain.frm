VERSION 5.00
Object = "{B6CC61F6-3F1A-4B00-9918-13F66F185263}#1.2#0"; "LblCtlsU.ocx"
Object = "{F6099CD4-9726-4AFE-B3A0-6C00C5ECC3A3}#1.2#0"; "LblCtlsA.ocx"
Begin VB.Form frmMain 
   Caption         =   "LabelControls 1.2 - Events Sample"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   7650
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   387
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   510
   StartUpPosition =   2  'Bildschirmmitte
   Begin LblCtlsLibACtl.WindowedLabel wlblA 
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   2400
      Width           =   2175
      _cx             =   3836
      _cy             =   873
      Appearance      =   0
      AutoSize        =   0
      BackColor       =   -2147483633
      BackStyle       =   1
      BorderStyle     =   0
      ClipLastLine    =   -1  'True
      DisabledEvents  =   0
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      HAlignment      =   0
      HoverTime       =   -1
      MousePointer    =   0
      OwnerDrawn      =   0   'False
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      SupportOLEDragImages=   -1  'True
      TextTruncationStyle=   0
      UseMnemonic     =   -1  'True
      UseSystemFont   =   -1  'True
      Text            =   "frmMain.frx":0000
   End
   Begin LblCtlsLibUCtl.WindowedLabel wlblU 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2175
      _cx             =   3836
      _cy             =   873
      Appearance      =   0
      AutoSize        =   0
      BackColor       =   -2147483633
      BackStyle       =   1
      BorderStyle     =   0
      ClipLastLine    =   -1  'True
      DisabledEvents  =   0
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      HAlignment      =   0
      HoverTime       =   -1
      MousePointer    =   0
      OwnerDrawn      =   0   'False
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      SupportOLEDragImages=   -1  'True
      TextTruncationStyle=   0
      UseMnemonic     =   -1  'True
      UseSystemFont   =   -1  'True
      Text            =   "frmMain.frx":006E
   End
   Begin LblCtlsLibUCtl.SysLink slU 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2175
      _cx             =   3836
      _cy             =   873
      Appearance      =   0
      AutomaticallyMarkLinksAsVisited=   -1  'True
      BackColor       =   -2147483633
      BackStyle       =   1
      BorderStyle     =   0
      DetectDoubleClicks=   -1  'True
      DisabledEvents  =   0
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      HAlignment      =   0
      HotTracking     =   -1  'True
      HoverTime       =   -1
      MousePointer    =   0
      ProcessContextMenuKeys=   -1  'True
      ProcessReturnKey=   -1  'True
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      ShowToolTips    =   -1  'True
      SupportOLEDragImages=   -1  'True
      UseMnemonic     =   -1  'True
      UseSystemFont   =   -1  'True
      UseVisualStyle  =   0   'False
      Text            =   "frmMain.frx":00E0
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "&Log"
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4230
      TabIndex        =   2
      Top             =   5280
      Width           =   2415
   End
   Begin VB.TextBox txtLog 
      Height          =   4815
      Left            =   3120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin LblCtlsLibACtl.WindowlessLabel lblA 
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   2175
      _cx             =   3836
      _cy             =   873
      Appearance      =   0
      AutoSize        =   0
      BackColor       =   -2147483633
      BackStyle       =   0
      BorderStyle     =   0
      ClipLastLine    =   -1  'True
      DisabledEvents  =   0
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      HAlignment      =   0
      HoverTime       =   -1
      MousePointer    =   0
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      SupportOLEDragImages=   -1  'True
      TextTruncationStyle=   0
      UseMnemonic     =   -1  'True
      UseSystemFont   =   -1  'True
      VAlignment      =   0
      WordWrapping    =   1
      Text            =   "frmMain.frx":0268
   End
   Begin LblCtlsLibUCtl.WindowlessLabel lblU 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   2175
      _cx             =   3836
      _cy             =   873
      Appearance      =   0
      AutoSize        =   0
      BackColor       =   -2147483633
      BackStyle       =   0
      BorderStyle     =   0
      ClipLastLine    =   -1  'True
      DisabledEvents  =   0
      DontRedraw      =   0   'False
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483630
      HAlignment      =   0
      HoverTime       =   -1
      MousePointer    =   0
      RegisterForOLEDragDrop=   -1  'True
      RightToLeft     =   0
      SupportOLEDragImages=   -1  'True
      TextTruncationStyle=   0
      UseMnemonic     =   -1  'True
      UseSystemFont   =   -1  'True
      VAlignment      =   0
      WordWrapping    =   1
      Text            =   "frmMain.frx":02DA
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements IHook


  Private bLog As Boolean
  Private objActiveCtl As Object


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Private Function IHook_KeyboardProcAfter(ByVal hookCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  '
End Function

Private Function IHook_KeyboardProcBefore(ByVal hookCode As Long, ByVal wParam As Long, ByVal lParam As Long, eatIt As Boolean) As Long
  Const KF_ALTDOWN = &H2000
  Const UIS_CLEAR = 2
  Const UISF_HIDEACCEL = &H2
  Const UISF_HIDEFOCUS = &H1
  Const VK_DOWN = &H28
  Const VK_LEFT = &H25
  Const VK_RIGHT = &H27
  Const VK_TAB = &H9
  Const VK_UP = &H26
  Const WM_CHANGEUISTATE = &H127

  If HiWord(lParam) And KF_ALTDOWN Then
    ' this will make keyboard and focus cues work
    SendMessageAsLong Me.hWnd, WM_CHANGEUISTATE, MakeDWord(UIS_CLEAR, UISF_HIDEACCEL Or UISF_HIDEFOCUS), 0
  Else
    Select Case wParam
      Case VK_LEFT, VK_RIGHT, VK_UP, VK_DOWN, VK_TAB
        ' this will make focus cues work
        SendMessageAsLong Me.hWnd, WM_CHANGEUISTATE, MakeDWord(UIS_CLEAR, UISF_HIDEFOCUS), 0
    End Select
  End If
End Function


Private Sub chkLog_Click()
  bLog = (chkLog.Value = CheckBoxConstants.vbChecked)
End Sub

Private Sub cmdAbout_Click()
  objActiveCtl.About
End Sub

Private Sub Form_Initialize()
  InitCommonControls
End Sub

Private Sub Form_Load()
  chkLog.Value = CheckBoxConstants.vbChecked
  Set objActiveCtl = slU

  InstallKeyboardHook Me
End Sub

Private Sub Form_Resize()
  Dim cx As Long

  If Me.WindowState <> vbMinimized Then
    slU.Move 8, 8
    wlblU.Move slU.Left, slU.Top + slU.Height + 3, slU.Width
    lblU.Move slU.Left, wlblU.Top + wlblU.Height + 3, slU.Width
    wlblA.Move slU.Left, lblU.Top + lblU.Height + 12, slU.Width
    lblA.Move slU.Left, wlblA.Top + wlblA.Height + 3, slU.Width
    txtLog.Move slU.Left + slU.Width + 10, 0, Me.ScaleWidth - slU.Width - slU.Left - 10, Me.ScaleHeight - cmdAbout.Height - 10
    chkLog.Move txtLog.Left, txtLog.Top + txtLog.Height + (Me.ScaleHeight - (txtLog.Top + txtLog.Height) - chkLog.Height) / 2
    cmdAbout.Move chkLog.Left + chkLog.Width + ((txtLog.Width - chkLog.Width) - cmdAbout.Width) / 2, txtLog.Top + txtLog.Height + (Me.ScaleHeight - (txtLog.Top + txtLog.Height) - cmdAbout.Height) / 2
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  RemoveKeyboardHook Me
End Sub

Private Sub lblA_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblA_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblA_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblA_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblA_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblA_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "lblA_DragDrop"
End Sub

Private Sub lblA_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "lblA_DragOver"
End Sub

Private Sub lblA_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblA_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblA_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblA_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblA_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblA_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblA_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblA_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblA_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblA_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblA_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblA_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblA_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "lblA_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblA_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblA_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblA_OLEDragDrop(ByVal data As LblCtlsLibACtl.IOLEDataObject, effect As LblCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "lblA_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=0x" & Hex(effect) & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    lblA.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub lblA_OLEDragEnter(ByVal data As LblCtlsLibACtl.IOLEDataObject, effect As LblCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "lblA_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=0x" & Hex(effect) & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub lblA_OLEDragLeave(ByVal data As LblCtlsLibACtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "lblA_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub lblA_OLEDragMouseMove(ByVal data As LblCtlsLibACtl.IOLEDataObject, effect As LblCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "lblA_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=0x" & Hex(effect) & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub lblA_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblA_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblA_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblA_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblA_ResizedControlWindow()
  AddLogEntry "lblA_ResizedControlWindow"
End Sub

Private Sub lblA_TextChanged()
  AddLogEntry "lblA_TextChanged"
End Sub

Private Sub lblA_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblA_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblA_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblA_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblU_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblU_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblU_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblU_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblU_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblU_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "lblU_DragDrop"
End Sub

Private Sub lblU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "lblU_DragOver"
End Sub

Private Sub lblU_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblU_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblU_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblU_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblU_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblU_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblU_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblU_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblU_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblU_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblU_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblU_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblU_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "lblU_MouseMove: button=" & Button & ", shift=" & Shift & ", x=" & X & ", y=" & Y
End Sub

Private Sub lblU_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblU_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblU_OLEDragDrop(ByVal data As LblCtlsLibUCtl.IOLEDataObject, effect As LblCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "lblU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=0x" & Hex(effect) & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    lblU.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub lblU_OLEDragEnter(ByVal data As LblCtlsLibUCtl.IOLEDataObject, effect As LblCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "lblU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=0x" & Hex(effect) & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub lblU_OLEDragLeave(ByVal data As LblCtlsLibUCtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "lblU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub lblU_OLEDragMouseMove(ByVal data As LblCtlsLibUCtl.IOLEDataObject, effect As LblCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "lblU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=0x" & Hex(effect) & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub lblU_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblU_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblU_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblU_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblU_ResizedControlWindow()
  AddLogEntry "lblU_ResizedControlWindow"
End Sub

Private Sub lblU_TextChanged()
  AddLogEntry "lblU_TextChanged"
End Sub

Private Sub lblU_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblU_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub lblU_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "lblU_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub slU_CaretChanged(ByVal previousCaretLink As LblCtlsLibUCtl.ILink, ByVal newCaretLink As LblCtlsLibUCtl.ILink)
  Dim str As String

  str = "slU_CaretChanged: previousCaretLink="
  If previousCaretLink Is Nothing Then
    str = str & "Nothing, newCaretLink="
  Else
    str = str & previousCaretLink.URL & ", newCaretLink="
  End If
  If newCaretLink Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & newCaretLink.URL
  End If

  AddLogEntry str
End Sub

Private Sub slU_Click(ByVal Link As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "slU_Click: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_ContextMenu(ByVal Link As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "slU_ContextMenu: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_CustomDraw(ByVal Link As LblCtlsLibUCtl.ILink, ByVal nonLinkPartIndex As Long, textColor As stdole.OLE_COLOR, TextBackColor As stdole.OLE_COLOR, ByVal drawStage As LblCtlsLibUCtl.CustomDrawStageConstants, ByVal linkState As LblCtlsLibUCtl.CustomDrawLinkStateConstants, ByVal hDC As Long, drawingRectangle As LblCtlsLibUCtl.RECTANGLE, furtherProcessing As LblCtlsLibUCtl.CustomDrawReturnValuesConstants)
'  Dim str As String
'
'  str = "slU_CustomDraw: link="
'  If Link Is Nothing Then
'    str = str & "Nothing"
'  Else
'    str = str & Link.URL
'  End If
'  str = str & ", nonLinkPartIndex=" & nonLinkPartIndex & ", textColor=0x" & Hex(textColor) & ", textBackColor=0x" & Hex(TextBackColor) & ", drawStage=" & drawStage & ", linkState=" & linkState & ", hDC=0x" & Hex(hDC) & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), furtherProcessing=" & furtherProcessing
'
'  AddLogEntry str
End Sub

Private Sub slU_CustomizeTextDrawing(ByVal isLink As Boolean, ByVal maxTextLength As Long, Text As String, ByVal hDC As Long, drawingRectangle As LblCtlsLibUCtl.RECTANGLE, ByVal drawTextFlags As Long, drawText As Boolean)
'  AddLogEntry "slU_CustomizeTextDrawing: isLink=" & isLink & ", maxTextLength=" & maxTextLength & ", text=" & Text & ", hDC=0x" & Hex(hDC) & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & "), drawTextFlags=0x" & Hex(drawTextFlags) & ", drawText=" & drawText
End Sub

Private Sub slU_DblClick(ByVal Link As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "slU_DblClick: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "slU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub slU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "slU_DragDrop"
End Sub

Private Sub slU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "slU_DragOver"
End Sub

Private Sub slU_GotFocus()
  AddLogEntry "slU_GotFocus"
  Set objActiveCtl = slU
End Sub

Private Sub slU_KeyDown(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "slU_KeyDown: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub slU_KeyPress(keyAscii As Integer)
  AddLogEntry "slU_KeyPress: keyAscii=" & keyAscii
End Sub

Private Sub slU_KeyUp(keyCode As Integer, ByVal shift As Integer)
  AddLogEntry "slU_KeyUp: keyCode=" & keyCode & ", shift=" & shift
End Sub

Private Sub slU_LinkGetInfoTipText(ByVal Link As LblCtlsLibUCtl.ILink, ByVal maxInfoTipLength As Long, infoTipText As String, abortToolTip As Boolean)
  If Link Is Nothing Then
    AddLogEntry "slU_LinkGetInfoTipText: Link=Nothing, maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
  Else
    AddLogEntry "slU_LinkGetInfoTipText: Link=" & Link.URL & ", maxInfoTipLength=" & maxInfoTipLength & ", infoTipText=" & infoTipText & ", abortToolTip=" & abortToolTip
    If infoTipText <> "" Then
      infoTipText = infoTipText & vbNewLine & "Index: " & Link.Index & vbNewLine & "Text: " & Link.Text
    Else
      infoTipText = "Index: " & Link.Index & vbNewLine & "Text: " & Link.Text
    End If
  End If
End Sub

Private Sub slU_LinkMouseEnter(ByVal Link As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "slU_LinkMouseEnter: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_LinkMouseLeave(ByVal Link As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "slU_LinkMouseLeave: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_LostFocus()
  AddLogEntry "slU_LostFocus"
End Sub

Private Sub slU_MClick(ByVal Link As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "slU_MClick: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_MDblClick(ByVal Link As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "slU_MDblClick: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_MouseDown(ByVal Link As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "slU_MouseDown: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_MouseEnter(ByVal Link As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "slU_MouseEnter: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_MouseHover(ByVal Link As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "slU_MouseHover: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_MouseLeave(ByVal Link As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "slU_MouseLeave: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_MouseMove(ByVal Link As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "slU_MouseMove: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  'AddLogEntry str
End Sub

Private Sub slU_MouseUp(ByVal Link As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "slU_MouseUp: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_OLEDragDrop(ByVal data As LblCtlsLibUCtl.IOLEDataObject, effect As LblCtlsLibUCtl.OLEDropEffectConstants, ByVal dropTarget As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "slU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=0x" & Hex(effect) & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    slU.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub slU_OLEDragEnter(ByVal data As LblCtlsLibUCtl.IOLEDataObject, effect As LblCtlsLibUCtl.OLEDropEffectConstants, ByVal dropTarget As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "slU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=0x" & Hex(effect) & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_OLEDragLeave(ByVal data As LblCtlsLibUCtl.IOLEDataObject, ByVal dropTarget As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "slU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_OLEDragMouseMove(ByVal data As LblCtlsLibUCtl.IOLEDataObject, effect As LblCtlsLibUCtl.OLEDropEffectConstants, ByVal dropTarget As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim files() As String
  Dim str As String

  str = "slU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=0x" & Hex(effect) & ", dropTarget="
  If dropTarget Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & dropTarget.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_OpenLink(ByVal Link As LblCtlsLibUCtl.ILink, ByVal causedBy As LblCtlsLibUCtl.EventCausedByConstants)
  Dim str As String

  str = "slU_OpenLink: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
    Link.Visited = True
  End If
  str = str & ", causedBy=" & causedBy

  AddLogEntry str
End Sub

Private Sub slU_RClick(ByVal Link As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "slU_RClick: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_RDblClick(ByVal Link As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "slU_RDblClick: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "slU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub slU_ResizedControlWindow()
  AddLogEntry "slU_ResizedControlWindow"
End Sub

Private Sub slU_TextChanged()
  AddLogEntry "slU_TextChanged"
End Sub

Private Sub slU_Validate(Cancel As Boolean)
  AddLogEntry "slU_Validate"
End Sub

Private Sub slU_XClick(ByVal Link As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "slU_XClick: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub slU_XDblClick(ByVal Link As LblCtlsLibUCtl.ILink, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As LblCtlsLibUCtl.HitTestConstants)
  Dim str As String

  str = "slU_XDblClick: link="
  If Link Is Nothing Then
    str = str & "Nothing"
  Else
    str = str & Link.URL
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y & ", hitTestDetails=0x" & Hex(hitTestDetails)

  AddLogEntry str
End Sub

Private Sub wlblA_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblA_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblA_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblA_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblA_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblA_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblA_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "wlblA_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub wlblA_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "wlblA_DragDrop"
End Sub

Private Sub wlblA_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "wlblA_DragOver"
End Sub

Private Sub wlblA_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblA_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblA_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblA_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblA_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblA_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblA_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblA_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblA_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblA_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblA_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblA_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblA_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "wlblA_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblA_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblA_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblA_OLEDragDrop(ByVal data As LblCtlsLibACtl.IOLEDataObject, effect As LblCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "wlblA_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=0x" & Hex(effect) & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    wlblA.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub wlblA_OLEDragEnter(ByVal data As LblCtlsLibACtl.IOLEDataObject, effect As LblCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "wlblA_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=0x" & Hex(effect) & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub wlblA_OLEDragLeave(ByVal data As LblCtlsLibACtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "wlblA_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub wlblA_OLEDragMouseMove(ByVal data As LblCtlsLibACtl.IOLEDataObject, effect As LblCtlsLibACtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "wlblA_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=0x" & Hex(effect) & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub wlblA_OwnerDraw(ByVal requiredAction As LblCtlsLibACtl.OwnerDrawActionConstants, ByVal controlState As LblCtlsLibACtl.OwnerDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As LblCtlsLibACtl.RECTANGLE)
  AddLogEntry "wlblA_OwnerDraw: requiredAction=" & requiredAction & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"
End Sub

Private Sub wlblA_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblA_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblA_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblA_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblA_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "wlblA_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub wlblA_ResizedControlWindow()
  AddLogEntry "wlblA_ResizedControlWindow"
End Sub

Private Sub wlblA_TextChanged()
  AddLogEntry "wlblA_TextChanged"
End Sub

Private Sub wlblA_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblA_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblA_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblA_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblU_Click(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblU_Click: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblU_ContextMenu(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblU_ContextMenu: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblU_DblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblU_DblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblU_DestroyedControlWindow(ByVal hWnd As Long)
  AddLogEntry "wlblU_DestroyedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub wlblU_DragDrop(Source As Control, x As Single, y As Single)
  AddLogEntry "wlblU_DragDrop"
End Sub

Private Sub wlblU_DragOver(Source As Control, x As Single, y As Single, State As Integer)
  AddLogEntry "wlblU_DragOver"
End Sub

Private Sub wlblU_MClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblU_MClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblU_MDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblU_MDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblU_MouseDown(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblU_MouseDown: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblU_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblU_MouseEnter: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblU_MouseHover(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblU_MouseHover: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblU_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblU_MouseLeave: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblU_MouseMove(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  'AddLogEntry "wlblU_MouseMove: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblU_MouseUp(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblU_MouseUp: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblU_OLEDragDrop(ByVal data As LblCtlsLibUCtl.IOLEDataObject, effect As LblCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "wlblU_OLEDragDrop: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=0x" & Hex(effect) & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str

  If data.GetFormat(vbCFFiles) Then
    files = data.GetData(vbCFFiles)
    str = Join(files, vbNewLine)
    wlblU.FinishOLEDragDrop
    MsgBox "Dropped files:" & vbNewLine & str
  End If
End Sub

Private Sub wlblU_OLEDragEnter(ByVal data As LblCtlsLibUCtl.IOLEDataObject, effect As LblCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "wlblU_OLEDragEnter: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=0x" & Hex(effect) & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub wlblU_OLEDragLeave(ByVal data As LblCtlsLibUCtl.IOLEDataObject, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "wlblU_OLEDragLeave: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub wlblU_OLEDragMouseMove(ByVal data As LblCtlsLibUCtl.IOLEDataObject, effect As LblCtlsLibUCtl.OLEDropEffectConstants, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  Dim files() As String
  Dim str As String

  str = "wlblU_OLEDragMouseMove: data="
  If data Is Nothing Then
    str = str & "Nothing"
  Else
    On Error Resume Next
    files = data.GetData(vbCFFiles)
    If Err Then
      str = str & "0 files"
    Else
      str = str & (UBound(files) - LBound(files) + 1) & " files"
    End If
  End If
  str = str & ", effect=0x" & Hex(effect) & ", button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y

  AddLogEntry str
End Sub

Private Sub wlblU_OwnerDraw(ByVal requiredAction As LblCtlsLibUCtl.OwnerDrawActionConstants, ByVal controlState As LblCtlsLibUCtl.OwnerDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As LblCtlsLibUCtl.RECTANGLE)
  AddLogEntry "wlblU_OwnerDraw: requiredAction=" & requiredAction & ", controlState=" & controlState & ", hDC=" & hDC & ", drawingRectangle=(" & drawingRectangle.Left & "," & drawingRectangle.Top & ")-(" & drawingRectangle.Right & "," & drawingRectangle.Bottom & ")"
End Sub

Private Sub wlblU_RClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblU_RClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblU_RDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblU_RDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblU_RecreatedControlWindow(ByVal hWnd As Long)
  AddLogEntry "wlblU_RecreatedControlWindow: hWnd=0x" & Hex(hWnd)
End Sub

Private Sub wlblU_ResizedControlWindow()
  AddLogEntry "wlblU_ResizedControlWindow"
End Sub

Private Sub wlblU_TextChanged()
  AddLogEntry "wlblU_TextChanged"
End Sub

Private Sub wlblU_XClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblU_XClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub

Private Sub wlblU_XDblClick(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  AddLogEntry "wlblU_XDblClick: button=" & button & ", shift=" & shift & ", x=" & x & ", y=" & y
End Sub


Private Sub AddLogEntry(ByVal txt As String)
  Dim pos As Long
  Static cLines As Long
  Static oldtxt As String

  If bLog Then
    cLines = cLines + 1
    If cLines > 50 Then
      ' delete the first line
      pos = InStr(oldtxt, vbNewLine)
      oldtxt = Mid(oldtxt, pos + Len(vbNewLine))
      cLines = cLines - 1
    End If
    oldtxt = oldtxt & (txt & vbNewLine)

    txtLog.Text = oldtxt
    txtLog.SelStart = Len(oldtxt)
  End If
End Sub

' returns the higher 16 bits of <value>
Private Function HiWord(ByVal Value As Long) As Integer
  Dim ret As Integer

  CopyMemory VarPtr(ret), VarPtr(Value) + LenB(ret), LenB(ret)

  HiWord = ret
End Function

' makes a 32 bits number out of two 16 bits parts
Private Function MakeDWord(ByVal Lo As Integer, ByVal hi As Integer) As Long
  Dim ret As Long

  CopyMemory VarPtr(ret), VarPtr(Lo), LenB(Lo)
  CopyMemory VarPtr(ret) + LenB(Lo), VarPtr(hi), LenB(hi)

  MakeDWord = ret
End Function
