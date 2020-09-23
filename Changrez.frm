VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Resolutions"
   ClientHeight    =   1920
   ClientLeft      =   2265
   ClientTop       =   2910
   ClientWidth     =   4005
   Icon            =   "Changrez.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1920
   ScaleWidth      =   4005
   Begin ChangeRes.SysIcon SysIcon1 
      Left            =   3120
      Top             =   0
      _ExtentX        =   1508
      _ExtentY        =   450
      IconText        =   ""
   End
   Begin VB.Frame fraChange 
      Caption         =   "Quick Change"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton cmdChange10 
         Caption         =   "Change to 1024 x 768"
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdChange64 
         Caption         =   "Change to 640 x 480"
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton cmdChange80 
         Caption         =   "Change to 800 x 600"
         Height          =   375
         Left            =   1080
         TabIndex        =   1
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Menu mnuTray 
      Caption         =   "mnuTray"
      Begin VB.Menu c1024768 
         Caption         =   "1024x768"
      End
      Begin VB.Menu c800600 
         Caption         =   "800x600"
      End
      Begin VB.Menu c640480 
         Caption         =   "640x480"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu popexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean


Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
    Private Const CCDEVICENAME = 32
    Private Const CCFORMNAME = 32
    Private Const DM_BITSPERPEL = &H60000
    Private Const DM_PELSWIDTH = &H80000
    Private Const DM_PELSHEIGHT = &H100000


Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    End Type

Private Sub about_Click()
MsgBox "ChangeRes 1.0" & vbCrLf & "Copyright (C) 2000 Justin Hinerman", vbInformation + vbOKOnly + vbDefaultButton1, "About..."
End Sub

Private Sub c1024768_Click()
Change1024768
End Sub

Private Sub c640480_Click()
Change640480
End Sub

Private Sub c800600_Click()
Change800600
End Sub

Private Sub cmdChange10_Click()
RetValue = ChangeRes(1024, 768, 32)
End Sub

Private Sub cmdChange64_Click()
RetValue = ChangeRes(640, 480, 32)
End Sub

Private Sub cmdChange80_Click()
RetValue = ChangeRes(800, 600, 32)
End Sub

Function ChangeRes(Width As Single, Height As Single, BPP As Integer) As Integer
    On Error GoTo ERROR_HANDLER
    Dim DevM As DEVMODE, I As Integer, ReturnVal As Boolean, _
    RetValue, OldWidth As Single, OldHeight As Single, _
    OldBPP As Integer
    Call EnumDisplaySettings(0&, -1, DevM)
    OldWidth = DevM.dmPelsWidth
    OldHeight = DevM.dmPelsHeight
    OldBPP = DevM.dmBitsPerPel
    I = 0


    Do
        ReturnVal = EnumDisplaySettings(0&, I, DevM)
        I = I + 1
    Loop Until (ReturnVal = False)
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    DevM.dmPelsWidth = Width
    DevM.dmPelsHeight = Height
    DevM.dmBitsPerPel = BPP
    Call ChangeDisplaySettings(DevM, 1)
    'RetValue = MsgBox("Do You Wish To Keep Your Screen Resolution To " & Width & "x" & Height & " - " & BPP & " BPP?", vbQuestion + vbOKCancel, "Change Resolution Confirm:")


    If RetValue = vbCancel Then
        DevM.dmPelsWidth = OldWidth
        DevM.dmPelsHeight = OldHeight
        DevM.dmBitsPerPel = OldBPP
        Call ChangeDisplaySettings(DevM, 1)
   '     MsgBox "Old Resolution(" & OldWidth & " x " & OldHeight & ", " & OldBPP & " Bit) Successfully Restored!", vbInformation + vbOKOnly, "Resolution Confirm:"
        ChangeRes = 0
    Else
        ChangeRes = 1
    End If
    Exit Function
ERROR_HANDLER:
    ChangeRes = 0
End Function

Private Sub exit_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set SysIcon1.NormalPicture = LoadResPicture(101, 1)
Me.Icon = LoadResPicture(101, 1)
SysIcon1.ShowNormalIcon
Me.Hide
UpdateRes
End Sub
Public Sub Change1024768()
RetValue = ChangeRes(1024, 768, 32)
UpdateRes
End Sub
Public Sub Change800600()
RetValue = ChangeRes(800, 600, 32)
UpdateRes
End Sub
Public Sub Change640480()
RetValue = ChangeRes(640, 480, 32)
UpdateRes
End Sub
Private Sub Form_Unload(Cancel As Integer)
SysIcon1.DeleteIcon
End Sub
Private Sub popexit_Click()
Unload Me
End
End Sub
Private Sub SysIcon1_IconRightUp()
PopupMenu mnuTray
End Sub
Public Sub UpdateRes()
   Dim DevM As DEVMODE, I As Integer, ReturnVal As Boolean, _
    RetValue, OldWidth As Single, OldHeight As Single, _
    OldBPP As Integer
    Call EnumDisplaySettings(0&, -1, DevM)
    OldWidth = DevM.dmPelsWidth
    OldHeight = DevM.dmPelsHeight
    OldBPP = DevM.dmBitsPerPel
        If OldWidth = 1024 Then
            c1024768.Checked = True
            c800600.Checked = False
            c640480.Checked = False
            Exit Sub
        ElseIf OldWidth = 800 Then
            c1024768.Checked = False
            c800600.Checked = True
            c640480.Checked = False
            Exit Sub
        ElseIf OldWidth = 640 Then
            c1024768.Checked = False
            c800600.Checked = False
            c640480.Checked = True
            Exit Sub
        End If
   
End Sub
