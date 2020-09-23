VERSION 5.00
Begin VB.UserControl SysIcon 
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1470
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   1110
   ScaleWidth      =   1470
   Begin VB.PictureBox picTray 
      Height          =   495
      Left            =   840
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   4
      Top             =   1200
      Width           =   495
   End
   Begin VB.PictureBox picNormal 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   120
      Top             =   1200
   End
   Begin VB.PictureBox picAnimate 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   840
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   480
      Width           =   495
   End
   Begin VB.Line Line2 
      X1              =   1440
      X2              =   0
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   720
      Y1              =   0
      Y2              =   1080
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   1095
      Left            =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Animated Icon"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   0
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Normal Icon"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   735
   End
End
Attribute VB_Name = "SysIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2

Const WM_LBUTTONDBLCLK = &H203
Const WM_LBUTTONDOWN = &H201
Const WM_RBUTTONDBLCLK = &H206
Const WM_RBUTTONDOWN = &H204
Const WM_RBUTTONUP = &H205
Const WM_LBUTTONUP = &H202
Const WM_MOUSEMOVE = &H200

Const NIM_ADD = &H0&
Const NIM_MODIFY = &H1
Const NIM_DELETE = &H2
Const NIF_MESSAGE = &H1
Const NIF_ICON = &H2
Const NIF_TIP = &H4

Private NI As NOTIFYICONDATA

Private Type NOTIFYICONDATA
     cbSize As Long
     hWnd As Long
     uID As Long
     uFlags As Long
     uCallbackMessage As Long
     hIcon As Long
     szTip As String * 64
End Type

Dim ShownFlag As Boolean
Dim ShownNormal As Boolean
Dim ShownAnim As Boolean
Dim DeleteAtEnd As Boolean

Dim IconDisplayText As String
Public Event IconMouseMove()
Public Event IconLeftDown()
Public Event IconLeftUp()
Public Event IconLeftDouble()
Public Event IconRightDown()
Public Event IconRightUp()
Public Event IconRightDouble()


'This procedure triggers the events
Private Sub picTray_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Msg
Msg = (X And &HFF) * &H100
    Select Case Msg
        Case 0
            RaiseEvent IconMouseMove
        Case &HF00
            RaiseEvent IconLeftDown
        Case &H1E00
            RaiseEvent IconLeftUp
        Case &H2D00
            RaiseEvent IconLeftDouble
        Case &H3C00
            RaiseEvent IconRightDown
        Case &H4B00
            RaiseEvent IconRightUp
        Case &H5A00
            RaiseEvent IconRightDouble
    End Select


End Sub

'This procedure set the System Tray Icon to the normal picture
Private Sub SetNormalIcon()
    picTray.Picture = picNormal.Picture
    NI.cbSize = Len(NI)
    NI.hWnd = picTray.hWnd
    NI.uID = 0
    NI.uID = NI.uID + 1
    NI.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    NI.uCallbackMessage = WM_MOUSEMOVE
    NI.hIcon = picTray.Picture
    'NI.szTip = Parameter set usin IconText Property
End Sub

'This procedure sets the System Tray Icon to the icon that is used
'to animate i.e. change between normal and this one.
Private Sub SetAnimIcon()
    picTray.Picture = picAnimate.Picture
    NI.cbSize = Len(NI)
    NI.hWnd = picTray.hWnd
    NI.uID = 0
    NI.uID = NI.uID + 1
    NI.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    NI.uCallbackMessage = WM_MOUSEMOVE
    NI.hIcon = picTray.Picture
    'NI.szTip = Parameter set usin IconText Property
End Sub

'Self explanatory
Public Sub ShowNormalIcon()
    Dim Result As Integer
    SetNormalIcon
    If ShownFlag = False Then
        Result = Shell_NotifyIconA(NIM_ADD, NI)
        ShownFlag = True
        ShownNormal = True
        ShownAnim = False
    ElseIf ShownFlag = True Then
        Result = Shell_NotifyIconA(NIM_MODIFY, NI)
        ShownNormal = True
        ShownAnim = False
    End If
End Sub

'Self explanatory
Public Sub ShowAnimIcon()
    Dim Result As Integer
    SetAnimIcon
    If ShownFlag = False Then
        Result = Shell_NotifyIconA(NIM_ADD, NI)
        ShownFlag = True
        ShownNormal = False
        ShownAnim = True
    ElseIf ShownFlag = True Then
        Result = Shell_NotifyIconA(NIM_MODIFY, NI)
        ShownNormal = False
        ShownAnim = True
    End If
End Sub

'Self explanatory
Public Sub DeleteIcon()
    Dim Result
    Result = Shell_NotifyIconA(NIM_DELETE, NI)
    ShownFlag = False
End Sub

'Time event to alternate the two icons
Private Sub Timer1_Timer()
    If ShownFlag = False Then
        ShowNormalIcon
        DeleteAtEnd = True
        Exit Sub
    End If
    
    If ShownNormal = True Then
        ShowAnimIcon
    ElseIf ShownAnim = True Then
        ShowNormalIcon
    End If
End Sub

'Initialize the flag variables
Private Sub UserControl_Initialize()
    ShownFlag = False
    ShownNormal = False
    ShownAnim = False
    DeleteAtEnd = False
End Sub

'MAke sure icone is deleted when the program closes
Private Sub UserControl_Terminate()
    If ShownFlag = True Then DeleteIcon
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Set NormalPicture = PropBag.ReadProperty("NormalPicture", Nothing)
    Set AnimPicture = PropBag.ReadProperty("AnimPicture", Nothing)
    Timer1.Interval = PropBag.ReadProperty("AnimDelay", 0)
    Timer1.Enabled = PropBag.ReadProperty("Animating", False)
    Let IconText = PropBag.ReadProperty("IconText", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("NormalPicture", NormalPicture, Nothing)
    Call PropBag.WriteProperty("AnimPicture", AnimPicture, Nothing)
    Call PropBag.WriteProperty("AnimDelay", Timer1.Interval, 0)
    Call PropBag.WriteProperty("Animating", Timer1.Enabled, False)
    Call PropBag.WriteProperty("IconText", IconText, Nothing)
End Sub

'+-------------------------------------------------------------+
'| Here are all of the properties that are used in the control |
'|            They are all pretty self explanatory             |
'+-------------------------------------------------------------+
 
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Timer1,Timer1,-1,Interval
Public Property Get AnimDelay() As Long
Attribute AnimDelay.VB_Description = "Returns/sets the number of milliseconds between calls to a Timer control's Timer event."
    AnimDelay = Timer1.Interval
End Property

Public Property Let AnimDelay(ByVal New_AnimDelay As Long)
    Timer1.Interval() = New_AnimDelay
    PropertyChanged "AnimDelay"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Timer1,Timer1,-1,Enabled
Public Property Get Animating() As Boolean
Attribute Animating.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Animating = Timer1.Enabled
End Property

Public Property Let Animating(ByVal New_Animating As Boolean)
    Timer1.Enabled() = New_Animating
    If (New_Animating = False) And (DeleteAtEnd = False) Then
        ShowNormalIcon
    ElseIf DeleteAtEnd = True Then
        DeleteIcon
        DeleteAtEnd = False
    End If
    PropertyChanged "Animating"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picAnimate,picAnimate,-1,Picture
Public Property Get AnimPicture() As Picture
    Set AnimPicture = picAnimate.Picture
End Property

Public Property Set AnimPicture(ByVal New_AnimPicture As Picture)
    Set picAnimate.Picture = New_AnimPicture
    PropertyChanged "AnimPicture"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picNormal,picNormal,-1,Picture
Public Property Get NormalPicture() As Picture
    Set NormalPicture = picNormal.Picture
End Property

Public Property Set NormalPicture(ByVal New_NormalPicture As Picture)
    Set picNormal.Picture = New_NormalPicture
    PropertyChanged "NormalPicture"
End Property

'Icontext Property
Public Property Get IconText() As String
    IconText = Mid(NI.szTip, 1, Len(NI.szTip) - 1)
End Property

Public Property Let IconText(New_IconText As String)
    NI.szTip = New_IconText & Chr$(0)
    
    If ShownNormal Then
        ShowNormalIcon
    ElseIf ShownAnim Then
        ShowAnimIcon
    End If
    
    PropertyChanged "IconText"
End Property


