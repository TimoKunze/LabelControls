VERSION 5.00
Object = "{B6CC61F6-3F1A-4B00-9918-13F66F185263}#1.2#0"; "LblCtlsU.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "LabelControls 1.2 - Hover Effects Sample"
   ClientHeight    =   1320
   ClientLeft      =   -15
   ClientTop       =   390
   ClientWidth     =   4620
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   88
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   308
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin LblCtlsLibUCtl.WindowedLabel lblButtonEffect 
      Height          =   255
      Left            =   2468
      TabIndex        =   2
      Top             =   360
      Width           =   1935
      _cx             =   3413
      _cy             =   450
      Appearance      =   0
      AutoSize        =   0
      BackColor       =   -2147483633
      BackStyle       =   1
      BorderStyle     =   0
      ClipLastLine    =   -1  'True
      DisabledEvents  =   4098
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
      HAlignment      =   1
      HoverTime       =   -1
      MousePointer    =   0
      OwnerDrawn      =   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      SupportOLEDragImages=   -1  'True
      TextTruncationStyle=   0
      UseMnemonic     =   -1  'True
      UseSystemFont   =   -1  'True
      Text            =   "frmMain.frx":0000
   End
   Begin LblCtlsLibUCtl.WindowedLabel lblBackColor 
      Height          =   255
      Left            =   218
      TabIndex        =   1
      Top             =   360
      Width           =   1935
      _cx             =   3413
      _cy             =   450
      Appearance      =   0
      AutoSize        =   0
      BackColor       =   -2147483633
      BackStyle       =   1
      BorderStyle     =   0
      ClipLastLine    =   -1  'True
      DisabledEvents  =   4098
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
      HAlignment      =   1
      HoverTime       =   -1
      MousePointer    =   0
      OwnerDrawn      =   0   'False
      RegisterForOLEDragDrop=   0   'False
      RightToLeft     =   0
      SupportOLEDragImages=   -1  'True
      TextTruncationStyle=   0
      UseMnemonic     =   -1  'True
      UseSystemFont   =   -1  'True
      Text            =   "frmMain.frx":003A
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
      Left            =   1103
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

  Implements IHook


  Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
  End Type


  Private themeableOS As Boolean


  Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal sz As Long)
  Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
  Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
  Private Declare Function DrawEdge Lib "user32.dll" (ByVal hDC As Long, ByRef qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
  Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextW" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
  Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
  Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
  Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
  Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryW" (ByVal lpLibFileName As Long) As Long
  Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, ByRef lColorRef As Long) As Long
  Private Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long


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


Private Sub cmdAbout_Click()
  lblBackColor.About
End Sub

Private Sub Form_Initialize()
  Dim hMod As Long

  InitCommonControls

  hMod = LoadLibrary(StrPtr("uxtheme.dll"))
  If hMod Then
    themeableOS = True
    FreeLibrary hMod
  End If
End Sub

Private Sub Form_Load()
  InstallKeyboardHook Me
  ' use a custom hover time so that the 'hot' state is entered immediately
  lblButtonEffect.HoverTime = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  RemoveKeyboardHook Me
End Sub

Private Sub lblBackColor_MouseEnter(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  lblBackColor.BackColor = vbHighlight
  lblBackColor.ForeColor = vbHighlightText
End Sub

Private Sub lblBackColor_MouseLeave(ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single)
  lblBackColor.BackColor = vbButtonFace
  lblBackColor.ForeColor = vbButtonText
End Sub

Private Sub lblButtonEffect_OwnerDraw(ByVal requiredAction As LblCtlsLibUCtl.OwnerDrawActionConstants, ByVal controlState As LblCtlsLibUCtl.OwnerDrawControlStateConstants, ByVal hDC As Long, drawingRectangle As LblCtlsLibUCtl.RECTANGLE)
  Const BDR_RAISEDINNER = &H4
  Const BF_BOTTOM = &H8
  Const BF_LEFT = &H1
  Const BF_RIGHT = &H4
  Const BF_TOP = &H2
  Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
  Const DT_CENTER = &H1
  Const DT_EDITCONTROL = &H2000
  Const DT_HIDEPREFIX = &H100000
  Const DT_RTLREADING = &H20000
  Const DT_SINGLELINE = &H20
  Const DT_VCENTER = &H4
  Const TRANSPARENT = 1
  Dim cx As Long
  Dim cy As Long
  Dim flags As Long
  Dim hBrush As Long
  Dim rc As RECT
  Dim txt As String

  If requiredAction And OwnerDrawActionConstants.odaDrawEntire Then
    ' draw background
    LSet rc = drawingRectangle
    hBrush = CreateSolidBrush(TranslateColor(lblButtonEffect.BackColor))
    If hBrush Then
      FillRect hDC, rc, hBrush
      DeleteObject hBrush
    End If

    ' draw border
    If controlState And OwnerDrawControlStateConstants.odcsHot Then
      DrawEdge hDC, rc, BDR_RAISEDINNER, BF_RECT
    End If

    ' draw text
    txt = lblButtonEffect.Text
    flags = DT_CENTER Or DT_VCENTER Or DT_EDITCONTROL Or DT_SINGLELINE
    If lblButtonEffect.RightToLeft And RightToLeftConstants.rtlText Then
      flags = flags Or DT_RTLREADING
    End If
    If controlState And OwnerDrawControlStateConstants.odcsNoAccelerator Then
      flags = flags Or DT_HIDEPREFIX
    End If
    SetBkMode hDC, TRANSPARENT
    DrawText hDC, StrPtr(txt), Len(txt), rc, flags
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

Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional ByVal hPal As Long = 0) As Long
  Const CLR_INVALID = &HFFFF&

  If OleTranslateColor(clr, hPal, TranslateColor) Then TranslateColor = CLR_INVALID
End Function
